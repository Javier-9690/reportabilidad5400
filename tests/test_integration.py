import os
import tempfile
import unittest
from io import BytesIO

from openpyxl import Workbook


TEST_DIR = tempfile.TemporaryDirectory(prefix="reportabilidad5400-tests-")
os.environ["DATABASE_URL"] = "sqlite:///" + os.path.join(TEST_DIR.name, "test.db")
os.environ["SECRET_KEY"] = "integration-test-secret"

from app import app
from gestion5s.web import ENTITY_MODEL, SessionLocal, TEMPLATES


TABS = [
    "censo", "eventos", "duplicidades", "encuesta", "atencion",
    "robos", "miscelaneo", "desviaciones", "solicitud_ot", "reclamos",
    "alarmas", "extensiones", "onboarding", "apertura", "cumplimiento", "entradas_salidas", "habitaciones_bloqueadas",
    "ordenamiento", "habitaciones_liberadas",
]

VIEWS = ["encuestas" if tab == "encuesta" else tab for tab in TABS]

FORM_DATA = {
    "censo": {"fecha": "2026-09-01", "censo_dia": "10", "censo_noche": "12"},
    "eventos": {"fecha": "2026-09-01", "horario": "Día", "que_ocurrio": "Prueba"},
    "duplicidades": {"semana": "88", "fecha": "2026-09-01"},
    "encuesta": {"fecha_hora": "2026-09-01T12:30", "q1_puntaje": "5"},
    "atencion": {"fecha": "2026-09-01", "tiempo_promedio": "03:45", "cantidad": "2"},
    "robos": {"fecha": "2026-09-01", "hora": "13:15"},
    "miscelaneo": {"ot": "OT-TEST", "fecha_creacion": "2026-09-01"},
    "desviaciones": {"fecha": "2026-09-01", "acciones": "Solicitar revisión a mantención."},
    "solicitud_ot": {"n_solicitud": "TEST", "fecha_inicio": "2026-09-01", "tiempo_respuesta": "04:20"},
    "reclamos": {"fecha": "2026-09-01"},
    "alarmas": {"fecha": "2026-09-01", "hora_reporte_salfa": "14:00"},
    "extensiones": {"fecha_solicitud": "2026-09-01", "centro_costos": "CC-300", "cant_clientes": "3", "tipo_solicitud": "Extensión"},
    "onboarding": {"fecha_hora": "2026-09-01T10:00", "nombre": "Prueba"},
    "apertura": {"fecha": "2026-09-01", "hora": "08:00"},
    "cumplimiento": {"fecha": "2026-09-01", "empresa": "Prueba"},
    "entradas_salidas": {"fecha_ingreso": "2026-09-01", "hora_entrada": "08:15", "nombre": "Prueba", "n_tarjeta": "000123"},
    "habitaciones_bloqueadas": {"fecha_bloqueo": "2026-09-01", "habitacion": "001", "motivo": "Reparación"},
    "ordenamiento": {"fecha_ejecucion": "2026-09-01", "habitacion": "001", "motivo_cambio": "Cambio de turno"},
    "habitaciones_liberadas": {"fecha_devolucion": "2026-09-01", "habitacion": "001", "observacion": "Revisada"},
}

IMPORT_ROWS = {
    "censo": ["2026-09-02", 5, 6, 11],
    "eventos": ["2026-09-02", "Día", "Prueba", "", ""],
    "duplicidades": [88, "2026-09-02"] + [""] * 13,
    "encuesta": ["2026-09-02 10:00", "Bien", 5, "Bien", 5, "Bien", 5, "Bien", 5, "Bien", 5, 25, 5.0, ""],
    "atencion": ["2026-09-02", "03:30", 2],
    "robos": ["2026-09-02", "08:30"] + [""] * 9,
    "miscelaneo": ["OT", "", "", "", "", "", "", "", "", "2026-09-02", "", "", "", "", ""],
    "desviaciones": ["", "2026-09-02"] + [""] * 11 + ["Señalizar y programar reparación."],
    "solicitud_ot": ["", "", "", "", "", "", "", "", "", "", "", "2026-09-02", "", "04:20", "", "", ""],
    "reclamos": ["", "2026-09-02"] + [""] * 13,
    "alarmas": ["", "", "", "2026-09-02", "", "", "", None, None, None, None, "08:00", "", "", "2026-09-02", "", ""],
    "extensiones": ["2026-09-02", "", "", "", "", "CC-400", 2, "Excepción", "2026-09-02", "2026-09-03", "", ""],
    "onboarding": ["2026-09-02 10:00", "", "", "", "", ""],
    "apertura": ["2026-09-02", "", "08:00", "", ""],
    "cumplimiento": ["2026-09-02", "", "", "", "", "", ""],
    "entradas_salidas": ["2026-09-02", "2026-09-03", "23:15", "07:45", "Empresa", "00042", "Persona", "12.345.678-9", "Noche", "A", "001", "Trabajo", "Recepción", "Sí", "000123", "Pendiente"],
    "habitaciones_bloqueadas": ["2026-09-02", "001", "000123", "Empresa", "00042", "Mantención", "Recepción informada", "2026-09-03", "", "", "", "Seguimiento"],
    "ordenamiento": ["2026-09-02", "Empresa", "001", "Persona", "00123456-7", "Noche", "002", "Cambio de turno", "Sí", "Revisión"],
    "habitaciones_liberadas": ["001", "Empresa", "Devolución", "2026-09-02", "Llaves recibidas", "2026-09-03", "Revisada"],
}


class IntegratedApplicationTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        app.config.update(TESTING=True)
        cls.client = app.test_client()

    def test_main_and_5s_pages_render(self):
        paths = [
            "/", "/imports", "/censos", "/no-match", "/curva", "/curva/nuevo",
            "/reports/dotacion-gerencia", "/reports/ocupabilidad", "/reports/egp", "/reports/fa",
            "/gestion-5s/", "/gestion-5s/registros", "/gestion-5s/dashboard", "/gestion-5s/health",
        ]
        for path in paths:
            with self.subTest(path=path):
                response = self.client.get(path, follow_redirects=True)
                self.assertEqual(response.status_code, 200)

        for tab in TABS:
            with self.subTest(tab=tab):
                response = self.client.get("/gestion-5s/panel", query_string={"tab": tab})
                self.assertEqual(response.status_code, 200)

    def test_all_manual_forms_lists_csv_and_delete(self):
        for tab, data in FORM_DATA.items():
            with self.subTest(tab=tab):
                response = self.client.post(
                    "/gestion-5s/panel",
                    query_string={"tab": tab},
                    data=data,
                    follow_redirects=False,
                )
                self.assertEqual(response.status_code, 302)

        for view in VIEWS:
            with self.subTest(view=view):
                response = self.client.get("/gestion-5s/registros", query_string={"vista": view})
                self.assertEqual(response.status_code, 200)

                response = self.client.get(f"/gestion-5s/download/{view}.csv")
                self.assertEqual(response.status_code, 200)

        session = SessionLocal()
        try:
            record_id = session.query(ENTITY_MODEL["censo"]).first().id
        finally:
            session.close()
        response = self.client.post(f"/gestion-5s/delete/censo/{record_id}")
        self.assertEqual(response.status_code, 302)

    def test_all_excel_templates_and_imports(self):
        for entity, headers in TEMPLATES.items():
            with self.subTest(entity=entity):
                response = self.client.get(f"/gestion-5s/template/{entity}.xlsx")
                self.assertEqual(response.status_code, 200)
                self.assertTrue(response.data.startswith(b"PK"))

                workbook = Workbook()
                sheet = workbook.active
                sheet.append(headers)
                sheet.append(IMPORT_ROWS[entity])
                payload = BytesIO()
                workbook.save(payload)
                payload.seek(0)

                response = self.client.post(
                    f"/gestion-5s/import/{entity}",
                    data={"file": (payload, f"{entity}.xlsx")},
                    content_type="multipart/form-data",
                    follow_redirects=True,
                )
                self.assertEqual(response.status_code, 200)
                self.assertNotIn(b"Error importando", response.data)
                self.assertNotIn(b"No se pudo leer", response.data)


if __name__ == "__main__":
    unittest.main()
