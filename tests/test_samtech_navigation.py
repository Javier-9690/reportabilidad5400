"""La navegación normal utiliza QR y conserva el histórico sin contarlo."""

from datetime import date
import json
import re
import unittest
from unittest.mock import patch

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from tests.test_integration import app
from tests.test_form_structure import PageStructure
from gestion5s import web
from gestion5s.user_reports import build_user_report


class SamtechNavigationTest(unittest.TestCase):
    def setUp(self):
        self.engine = create_engine("sqlite://")
        web.Base.metadata.create_all(self.engine)
        self.sessions = sessionmaker(bind=self.engine)
        self.session_patch = patch.object(web, "SessionLocal", self.sessions)
        self.session_patch.start()
        self.client = app.test_client()

    def tearDown(self):
        self.session_patch.stop()
        self.engine.dispose()

    def test_active_navigation_only_offers_qr_and_keeps_template_import(self):
        active = set(web.ENTITY_MODEL) - {"samtech_usuarios"}
        for entity in active:
            with self.subTest(entity=entity):
                response = self.client.get("/gestion-5s/panel", query_string={"tab": entity})
                self.assertEqual(response.status_code, 200)
                page = response.get_data(as_text=True)
                links = PageStructure(page).links
                self.assertNotIn("/gestion-5s/panel?tab=samtech_usuarios", links)
                self.assertIn("/gestion-5s/panel?tab=samtech_qr", links)
                self.assertNotIn("Samtech usuarios", page)
        page = self.client.get("/gestion-5s/panel?tab=samtech_qr").get_data(as_text=True)
        self.assertIn('action="/gestion-5s/import/samtech_qr"', page)
        self.assertIn('/gestion-5s/template/samtech_qr.xlsx', page)
        self.assertIn('name="csrf_token"', page)
        self.assertIn('Ver registros', page)
        listing = self.client.get("/gestion-5s/registros?vista=samtech_qr").get_data(as_text=True)
        selector = re.search(r'<select id="recordCategory".*?</select>', listing, re.S).group(0)
        self.assertEqual(set(re.findall(r'<option value="([^"]+)"', selector)), active)

    def test_dashboard_and_report_only_count_qr_and_keep_legacy_data(self):
        with self.sessions() as db:
            db.add(web.SamtechUsuarioEntry(ticket="HISTORICO", fecha_creacion=date(2026, 9, 1), estado="Abierto"))
            db.commit()
        page = self.client.get("/gestion-5s/dashboard").get_data(as_text=True)
        self.assertNotIn("const series =", page)
        with self.sessions() as db:
            db.add_all([
                web.SamtechQRUsuarioEntry(ticket="QR-1", fecha_creacion=date(2026, 9, 29), estado="No Iniciada"),
                web.SamtechQRUsuarioEntry(ticket="QR-2", fecha_inicio=date(2026, 9, 30), estado="Aprobada"),
                web.SamtechQRUsuarioEntry(ticket="QR-SIN-FECHA"),
            ])
            db.commit()
        page = self.client.get("/gestion-5s/dashboard").get_data(as_text=True)
        labels = json.loads(re.search(r"const labels = (.*?);", page).group(1))
        series = json.loads(re.search(r"const series = (.*?);", page).group(1))
        self.assertEqual(labels, ["2026-09-29", "2026-09-30"])
        self.assertEqual(series["samtech_qr"], [1, 1])
        self.assertNotIn("samtech_usuarios", series)
        self.assertNotIn("Samtech usuarios", page)
        self.assertNotIn("samtechUsuariosChart", page)
        self.assertIn('id="samtechQRChart"', page)
        canvases = set(re.findall(r'<canvas id="([^"]+)"', page))
        chart_targets = set(re.findall(r"new Chart\(document.getElementById\('([^']+)'\)", page))
        self.assertEqual(chart_targets, canvases)
        with self.sessions() as db:
            report = build_user_report(db, web.ENTITY_MODEL, date(2026, 9, 1), date(2026, 9, 30))
            self.assertEqual(db.query(web.SamtechUsuarioEntry).one().ticket, "HISTORICO")
            self.assertEqual(db.query(web.SamtechQRUsuarioEntry).count(), 3)
        source = next(row for row in report["categories"] if row["label"] == "Solicitudes de usuarios")
        self.assertEqual(source["totals"]["records"], 2)
        self.assertEqual(source["totals"]["open"], 1)
        self.assertEqual(source["totals"]["closed"], 1)
        historical = self.client.get("/gestion-5s/registros?vista=samtech_usuarios").get_data(as_text=True)
        self.assertIn("Samtech usuarios (histórico)", historical)
        self.assertIn("HISTORICO", historical)


if __name__ == "__main__":
    unittest.main()
