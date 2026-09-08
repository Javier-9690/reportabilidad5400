import unittest
from datetime import date, time
from html.parser import HTMLParser
from urllib.parse import parse_qs, urlsplit

from tests.test_integration import FORM_DATA, app
from gestion5s.web import ENTITY_MODEL, SessionLocal


class FormValues(HTMLParser):
    """Lee los valores que enviaría el formulario mostrado al usuario."""
    def __init__(self, page):
        super().__init__(convert_charrefs=True)
        self.values = {}
        self.textarea = None
        self.feed(page)

    def handle_starttag(self, tag, attributes):
        attributes = dict(attributes)
        name = attributes.get("name")
        if tag == "input" and name and "disabled" not in attributes:
            if attributes.get("type") != "checkbox" or "checked" in attributes:
                self.values[name] = attributes.get("value", "")
        elif tag == "textarea" and name:
            self.textarea = name
            self.values[name] = ""

    def handle_data(self, data):
        if self.textarea:
            self.values[self.textarea] += data

    def handle_endtag(self, tag):
        if tag == "textarea":
            self.textarea = None


class EditHotelRecordsTest(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()
        self.records = []

    def tearDown(self):
        with SessionLocal() as db:
            for entity, rid in self.records:
                record = db.get(ENTITY_MODEL[entity], rid)
                if record is not None:
                    db.delete(record)
            db.commit()

    def create_record(self, entity):
        tab = "encuesta" if entity == "encuestas" else entity
        response = self.client.post("/gestion-5s/panel", query_string={"tab": tab}, data=FORM_DATA[tab])
        self.assertEqual(response.status_code, 302)
        with SessionLocal() as db:
            model = ENTITY_MODEL[entity]
            rid = db.query(model).order_by(model.id.desc()).first().id
        self.records.append((entity, rid))
        return rid

    def snapshot(self, entity, rid):
        with SessionLocal() as db:
            record = db.get(ENTITY_MODEL[entity], rid)
            return {column.name: getattr(record, column.name) for column in record.__table__.columns}

    def open_editor(self, entity, rid, **query):
        response = self.client.get(f"/gestion-5s/edit/{entity}/{rid}", query_string=query)
        self.assertEqual(response.status_code, 200)
        return FormValues(response.get_data(as_text=True)).values

    def test_all_fifteen_types_have_edit_icon_and_update_original_record(self):
        changes = {
            "censo": ("censo_dia", "23", 23),
            "eventos": ("que_ocurrio", 'Cambio con "comillas" & <texto>', 'Cambio con "comillas" & <texto>'),
            "duplicidades": ("id_interno", "ID-EDITADO", "ID-EDITADO"),
            "encuestas": ("comentarios", "Nueva evaluación", "Nueva evaluación"),
            "atencion": ("tiempo_promedio_sec", "125:07", 7507),
            "robos": ("hora", "17:25:36", time(17, 25, 36)),
            "miscelaneo": ("comentario", "Seguimiento\nSegunda línea", "Seguimiento\nSegunda línea"),
            "desviaciones": ("acciones", "Reparar y verificar el cierre.", "Reparar y verificar el cierre."),
            "solicitud_ot": ("tiempo_respuesta_sec", "00:00", 0),
            "reclamos": ("estatus", "Cerrado", "Cerrado"),
            "alarmas": ("aviso_mantencion_h", "1.25", 1.25),
            "extensiones": ("cant_clientes", "0", 0),
            "onboarding": ("nombre", "Nombre corregido", "Nombre corregido"),
            "apertura": ("estado_chapa", "Reparada", "Reparada"),
            "cumplimiento": ("fecha", "2026-09-03", date(2026, 9, 3)),
        }
        for entity, (name, value, expected) in changes.items():
            with self.subTest(entity=entity):
                rid = self.create_record(entity)
                before = self.snapshot(entity, rid)
                with SessionLocal() as db:
                    count = db.query(ENTITY_MODEL[entity]).count()
                listing = self.client.get("/gestion-5s/registros", query_string={"vista": entity}).get_data(as_text=True)
                self.assertIn(f'/gestion-5s/edit/{entity}/{rid}?', listing)
                self.assertIn('class="fas fa-pen"', listing)
                self.assertIn(f'/gestion-5s/delete/{entity}/{rid}', listing)
                data = self.open_editor(entity, rid)
                self.assertEqual(self.snapshot(entity, rid), before, "Abrir el editor no modifica el registro")
                data.update({name: value, "id": "987654321", "creado": "1900-01-01"})
                response = self.client.post(f"/gestion-5s/edit/{entity}/{rid}", data=data)
                self.assertEqual(response.status_code, 302)
                self.assertEqual(response.location, f"/gestion-5s/registros?vista={entity}")
                after = self.snapshot(entity, rid)
                self.assertEqual(after[name], expected)
                self.assertEqual(after["id"], before["id"])
                self.assertEqual(after["creado"], before["creado"])
                with SessionLocal() as db:
                    self.assertEqual(db.query(ENTITY_MODEL[entity]).count(), count)
                # Los datos guardados vuelven a aparecer en el formulario.
                reopened = self.open_editor(entity, rid)
                if isinstance(expected, str):
                    self.assertEqual(reopened[name], expected)

    def test_totals_and_optional_scores_recalculate(self):
        rid = self.create_record("censo")
        data = self.open_editor("censo", rid)
        data.update(censo_dia="7", censo_noche="8", total="999")
        self.assertEqual(self.client.post(f"/gestion-5s/edit/censo/{rid}", data=data).status_code, 302)
        self.assertEqual(self.snapshot("censo", rid)["total"], 15)
        data = self.open_editor("censo", rid)
        data.pop("censo_total_auto")
        data["total"] = "20"
        self.assertEqual(self.client.post(f"/gestion-5s/edit/censo/{rid}", data=data).status_code, 302)
        self.assertEqual(self.snapshot("censo", rid)["total"], 20)
        self.assertNotIn("censo_total_auto", self.open_editor("censo", rid))

        rid = self.create_record("encuestas")
        data = self.open_editor("encuestas", rid)
        data.update(q1_puntaje="3", q2_puntaje="4", total="999", promedio="999")
        self.assertEqual(self.client.post(f"/gestion-5s/edit/encuestas/{rid}", data=data).status_code, 302)
        record = self.snapshot("encuestas", rid)
        self.assertEqual((record["total"], record["promedio"]), (7, 3.5))
        data = self.open_editor("encuestas", rid)
        data.update(q1_puntaje="", q2_puntaje="")
        self.assertEqual(self.client.post(f"/gestion-5s/edit/encuestas/{rid}", data=data).status_code, 302)
        record = self.snapshot("encuestas", rid)
        self.assertIsNone(record["total"])
        self.assertIsNone(record["promedio"])

    def test_invalid_values_do_not_change_any_stored_field(self):
        cases = [
            ("censo", "censo_dia", "-1"), ("censo", "censo_dia", "2.5"),
            ("censo", "fecha", "2026-02-30"), ("censo", "fecha", ""),
            ("atencion", "tiempo_promedio_sec", "03:75"),
            ("robos", "hora", "25:00"), ("encuestas", "q1_puntaje", "6"),
            ("alarmas", "aviso_mantencion_h", "NaN"),
            ("extensiones", "cant_clientes", "2147483648"),
            ("onboarding", "nombre", "a" * 201),
        ]
        for entity, name, invalid in cases:
            with self.subTest(entity=entity, field=name, value=invalid):
                rid = self.create_record(entity)
                before = self.snapshot(entity, rid)
                data = self.open_editor(entity, rid)
                data[name] = invalid
                response = self.client.post(f"/gestion-5s/edit/{entity}/{rid}", data=data)
                self.assertEqual(response.status_code, 422)
                self.assertEqual(self.snapshot(entity, rid), before)
                self.assertEqual(FormValues(response.get_data(as_text=True)).values[name], invalid)

    def test_optional_fields_clear_and_unchanged_times_keep_seconds(self):
        rid = self.create_record("apertura")
        with SessionLocal() as db:
            db.get(ENTITY_MODEL["apertura"], rid).hora = time(8, 30, 45)
            db.commit()
        data = self.open_editor("apertura", rid)
        self.assertEqual(data["hora"], "08:30:45")
        data["responsable"] = "Recepción"
        self.assertEqual(self.client.post(f"/gestion-5s/edit/apertura/{rid}", data=data).status_code, 302)
        self.assertEqual(self.snapshot("apertura", rid)["hora"], time(8, 30, 45))
        data = self.open_editor("apertura", rid)
        data.update(hora="", responsable="")
        self.assertEqual(self.client.post(f"/gestion-5s/edit/apertura/{rid}", data=data).status_code, 302)
        record = self.snapshot("apertura", rid)
        self.assertIsNone(record["hora"])
        self.assertIsNone(record["responsable"])

    def test_return_preserves_filters_and_rejects_external_targets(self):
        rid = self.create_record("censo")
        target = "/gestion-5s/registros?vista=censo&from=2026-09-01&to=2026-09-10&semana=88"
        data = self.open_editor("censo", rid, next=target)
        response = self.client.post(f"/gestion-5s/edit/censo/{rid}", data=data)
        self.assertEqual(parse_qs(urlsplit(response.location).query), parse_qs(urlsplit(target).query))
        self.assertEqual(urlsplit(response.location).path, "/gestion-5s/registros")
        data = self.open_editor("censo", rid)
        data["next"] = "https://example.org/gestion-5s/registros"
        response = self.client.post(f"/gestion-5s/edit/censo/{rid}", data=data)
        self.assertEqual(response.location, "/gestion-5s/registros?vista=censo")

    def test_stale_form_cannot_overwrite_a_newer_change(self):
        rid = self.create_record("eventos")
        first = self.open_editor("eventos", rid)
        second = self.open_editor("eventos", rid)
        first["accion"] = "Actualización del primer editor"
        self.assertEqual(self.client.post(f"/gestion-5s/edit/eventos/{rid}", data=first).status_code, 302)
        second["accion"] = "Actualización antigua"
        response = self.client.post(f"/gestion-5s/edit/eventos/{rid}", data=second)
        self.assertEqual(response.status_code, 409)
        self.assertEqual(self.snapshot("eventos", rid)["accion"], first["accion"])

    def test_invalid_entities_missing_records_and_unverified_forms_do_not_write(self):
        self.assertEqual(self.client.get("/gestion-5s/edit/desconocido/1").status_code, 404)
        self.assertEqual(self.client.get("/gestion-5s/edit/censo/987654321").status_code, 404)
        rid = self.create_record("censo")
        before = self.snapshot("censo", rid)
        data = self.open_editor("censo", rid)
        data.pop("csrf_token")
        self.assertEqual(self.client.post(f"/gestion-5s/edit/censo/{rid}", data=data).status_code, 400)
        self.assertEqual(self.snapshot("censo", rid), before)


if __name__ == "__main__":
    unittest.main()
