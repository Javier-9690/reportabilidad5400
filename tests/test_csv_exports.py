"""El CSV de cada categoría mantiene las columnas y el contenido al descargar."""

import codecs
import unittest
from unittest.mock import patch

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from tests.csv_helpers import excel_csv_reader
from tests.test_integration import FORM_DATA, app
from gestion5s import web


EXPECTED_COLUMNS = {
    "censo": 4, "eventos": 5, "duplicidades": 15, "encuestas": 14,
    "atencion": 3, "robos": 11, "miscelaneo": 15, "desviaciones": 14,
    "solicitud_ot": 29, "reclamos": 15, "alarmas": 17, "extensiones": 12,
    "onboarding": 6, "apertura": 5, "cumplimiento": 7,
    "entradas_salidas": 16, "habitaciones_bloqueadas": 12,
    "ordenamiento": 10, "habitaciones_liberadas": 7,
    "samtech_usuarios": 15, "samtech_qr": 15,
}


class CsvExportsTest(unittest.TestCase):
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

    def download(self, entity):
        response = self.client.get(f"/gestion-5s/download/{entity}.csv")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "text/csv")
        self.assertIn(f"filename={entity}.csv", response.headers["Content-Disposition"])
        self.assertTrue(response.data.startswith(codecs.BOM_UTF8 + b"sep=;\r\n"))
        self.assertEqual(response.data.count(codecs.BOM_UTF8), 1)
        reader = excel_csv_reader(response.data)
        self.assertEqual(len(reader.fieldnames), EXPECTED_COLUMNS[entity])
        rows = list(reader)
        for row in rows:
            self.assertEqual(set(row), set(reader.fieldnames))
            self.assertNotIn(None, row.values())
        return rows

    def test_all_categories_keep_separate_columns_and_empty_values(self):
        self.assertEqual(set(EXPECTED_COLUMNS), set(web.ENTITY_MODEL))
        for entity in web.ENTITY_MODEL:
            with self.subTest(entity=entity):
                tab = "encuesta" if entity == "encuestas" else entity
                response = self.client.post(
                    "/gestion-5s/panel", query_string={"tab": tab}, data=FORM_DATA[tab]
                )
                self.assertEqual(response.status_code, 302)
                self.assertEqual(len(self.download(entity)), 1)

    def test_empty_categories_still_download_column_headers(self):
        for entity in web.ENTITY_MODEL:
            with self.subTest(entity=entity):
                self.assertEqual(self.download(entity), [])

    def test_delimiters_quotes_newlines_accents_and_blank_dates_keep_their_cells(self):
        comment = 'Revisión de baño; puerta, chapa y "llave".\r\nSegunda línea: señalización.\nÚltima línea.'
        response = self.client.post(
            "/gestion-5s/panel?tab=samtech_qr",
            data={"ticket": "000123", "division": "División Ñandú",
                  "area": "Mantención; hotelería", "falla": 'Puerta, "chapa"',
                  "comentario": comment, "estado": "En progreso"},
        )
        self.assertEqual(response.status_code, 302)
        # Store both kinds of line endings: this export must not rewrite text.
        with self.sessions() as db:
            record = db.query(web.SamtechQRUsuarioEntry).one()
            record.comentario = comment
            db.commit()
        rows = self.download("samtech_qr")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Ticket"], "000123")
        self.assertEqual(rows[0]["División"], "División Ñandú")
        self.assertEqual(rows[0]["Área"], "Mantención; hotelería")
        self.assertEqual(rows[0]["Falla"], 'Puerta, "chapa"')
        self.assertEqual(rows[0]["Comentario"], comment)
        for field in ("Fecha creación", "Fecha inicio", "Fecha término", "Fecha aprobación", "Empresa"):
            self.assertEqual(rows[0][field], "")


if __name__ == "__main__":
    unittest.main()
