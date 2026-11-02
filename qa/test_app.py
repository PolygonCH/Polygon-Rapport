import io
import json
import os
import tempfile
import unittest

os.environ["DATA_DIR"] = tempfile.mkdtemp(prefix="rapport-test-")

import app


def request(path, method="GET", payload=None, headers=None):
    path, _, query = path.partition("?")
    raw = json.dumps(payload).encode("utf-8") if payload is not None else b""
    environ = {
        "REQUEST_METHOD": method,
        "PATH_INFO": path,
        "QUERY_STRING": query,
        "CONTENT_LENGTH": str(len(raw)),
        "wsgi.input": io.BytesIO(raw),
        "HTTP_HOST": "example.test",
        "wsgi.url_scheme": "https",
    }
    for key, value in (headers or {}).items():
        environ["HTTP_" + key.upper().replace("-", "_")] = value
    captured = {}

    def start_response(status, response_headers):
        captured["status"] = status
        captured["headers"] = dict(response_headers)

    body = b"".join(app.application(environ, start_response))
    return captured["status"], captured["headers"], body


def sample_report(language="de"):
    return {
        "language": language,
        "date": "2026-10-27",
        "reporterName": "Anna Beispiel",
        "region": "Bern / Wallis",
        "projectLead": "Imer Abazi",
        "damageType": "Wasser",
        "projectNumber": "P-26001",
        "projectName": "Sanierung Muster AG",
        "projectStatus": "In Arbeit",
        "location": {"name": "Sanierung Muster AG", "address": "Testweg 1", "city": "3000 Bern"},
        "areas": ["fh", "tr"],
        "work": {
            "fh": [{"role": "Facharbeiter", "employeeType": "Polygon", "firstName": "Max", "lastName": "Muster", "hours": 8}],
            "tr": [{"role": "Leckortung", "employeeType": "Temporär", "firstName": "Luca", "lastName": "Test", "hours": 2, "temp": "Tempo AG"}],
        },
        "vehicles": {"fh": [{"type": "PW", "count": 1, "billing": "Kilometer", "amount": 20, "unitPrice": 1.9}]},
        "equipment": {"tr": [{"item": "TR202 Kondenstrockner 40", "quantity": 3, "unit": "Stk./Tag", "unitPrice": 20}]},
        "products": {"fh": [{"item": "Neutralreiniger", "amount": 2, "unitPrice": 22.5}]},
        "consumables": {},
        "disposal": {"fh": {"method": "Entsorgung Lager Polygon", "amount": 1, "unit": "m³"}},
        "thirdParty": [{"description": "Elektriker", "company": "Strom AG", "price": 250}],
        "costs": {"insurance": "Muster Versicherung", "inspectorName": "Sara Sicher", "inspectorEmail": "sara@example.ch", "claimNumber": "S-123", "invoiceAmount": 2000},
        "notes": "Testbemerkung",
        "status": "Entwurf",
    }


class AppTests(unittest.TestCase):
    def test_costs_and_exports(self):
        report = app.normalize_report(sample_report())
        costs = app.calculate_costs(report)
        self.assertAlmostEqual(costs["labour"], 830)
        self.assertAlmostEqual(costs["vehicles"], 38)
        self.assertAlmostEqual(costs["equipment"], 60)
        self.assertAlmostEqual(costs["materials"], 175)
        self.assertAlmostEqual(costs["thirdParty"], 250)
        self.assertAlmostEqual(costs["total"], 1353)
        self.assertAlmostEqual(costs["margin"], 647)
        for language in ("de", "fr", "it"):
            report["language"] = language
            self.assertGreater(len(app.build_excel(report)), 10_000)
        pdf = app.build_pdf(report)
        self.assertTrue(pdf.startswith(b"%PDF"))
        self.assertGreater(len(pdf), 4_000)

    def test_api_save_search_fields_and_hours(self):
        status, _, body = request("/api/reports", "POST", sample_report())
        self.assertTrue(status.startswith("201"), body)
        report_id = json.loads(body)["id"]
        status, _, body = request("/api/reports")
        self.assertTrue(status.startswith("200"))
        saved = json.loads(body)["reports"][0]
        self.assertEqual(saved["projectNumber"], "P-26001")
        self.assertEqual(saved["projectName"], "Sanierung Muster AG")
        self.assertEqual(saved["projectStatus"], "In Arbeit")
        status, _, body = request("/api/hour-summary?date=2026-10-27")
        self.assertTrue(status.startswith("200"))
        names = {item["name"] for item in json.loads(body)["summary"]}
        self.assertEqual(names, {"Max Muster", "Luca Test"})
        status, _, body = request(f"/api/reports/{report_id}/excel")
        self.assertTrue(status.startswith("200"), body)
        self.assertTrue(body.startswith(b"PK"))
        status, _, body = request(f"/api/reports/{report_id}/workflow", "POST", {"status": "Zur Kontrolle"})
        self.assertTrue(status.startswith("200"), body)
        self.assertTrue(json.loads(body)["pdfCreated"])
        status, _, body = request(f"/api/reports/{report_id}/workflow", "POST", {"status": "Frontoffice"})
        self.assertTrue(status.startswith("200"), body)
        with app.db_connection() as connection:
            payload = json.loads(connection.execute("SELECT payload FROM reports WHERE id = ?", (report_id,)).fetchone()["payload"])
        self.assertTrue((app.DATA_DIR / payload["_pdfPath"]).exists())
        self.assertTrue((app.DATA_DIR / payload["_frontofficePath"]).exists())


if __name__ == "__main__":
    unittest.main()
