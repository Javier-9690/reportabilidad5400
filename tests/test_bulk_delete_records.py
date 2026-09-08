import unittest
import time
from datetime import date, timedelta
from unittest.mock import patch
from urllib.parse import parse_qs, urlsplit

from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker

from tests.test_integration import FORM_DATA, app
from tests.test_edit_records import FormValues
from gestion5s import web


class BulkDeleteHotelRecordsTest(unittest.TestCase):
    def setUp(self):
        # Cada prueba usa su propia base, incluso para eliminar un módulo completo.
        self.engine = create_engine("sqlite://")
        web.Base.metadata.create_all(self.engine)
        self.sessions = sessionmaker(bind=self.engine)
        self.session_patch = patch.object(web, "SessionLocal", self.sessions)
        self.session_patch.start()
        self.client = app.test_client()

    def tearDown(self):
        self.session_patch.stop()
        self.engine.dispose()

    def create_record(self, entity, day="2026-09-01", clock="12:00:00"):
        tab = "encuesta" if entity == "encuestas" else entity
        data = dict(FORM_DATA[tab])
        field = web.ENTITY_DATE_FIELD.get(entity, "fecha")
        data[field] = f"{day}T{clock}" if field == "fecha_hora" else day
        response = self.client.post("/gestion-5s/panel", query_string={"tab": tab}, data=data)
        self.assertEqual(response.status_code, 302)
        with self.sessions() as db:
            model = web.ENTITY_MODEL[entity]
            return db.query(model).order_by(model.id.desc()).first().id

    def ids(self, entity):
        with self.sessions() as db:
            return {row.id for row in db.query(web.ENTITY_MODEL[entity]).all()}

    def preview(self, entity, mode, ids=(), filters=None):
        response = self.client.get("/gestion-5s/registros", query_string={"vista": entity, **(filters or {})})
        self.assertEqual(response.status_code, 200)
        data = FormValues(response.get_data(as_text=True)).values
        data.update(mode=mode, ids=[str(rid) for rid in ids])
        return self.client.post(f"/gestion-5s/delete/{entity}/bulk/confirm", data=data)

    def confirm(self, entity, preview):
        self.assertEqual(preview.status_code, 200)
        data = FormValues(preview.get_data(as_text=True)).values
        self.assertIn("confirmation_token", data)
        return self.client.post(f"/gestion-5s/delete/{entity}/bulk", data=data, follow_redirects=True)

    def test_selected_deletes_only_marked_rows_in_all_fifteen_modules(self):
        for entity in web.ENTITY_MODEL:
            with self.subTest(entity=entity):
                chosen = [self.create_record(entity), self.create_record(entity)]
                remaining = self.create_record(entity)
                other = "eventos" if entity == "censo" else "censo"
                sentinel = self.create_record(other, "2027-01-01")
                filters = {"from": "2026-09-01", "to": "2026-09-10"}
                before = self.ids(entity)
                preview = self.preview(entity, "selected", chosen, filters)
                self.assertEqual(self.ids(entity), before, "La confirmación previa no elimina datos")
                self.assertIn("Confirmar eliminación de 2 registros", preview.get_data(as_text=True))
                response = self.confirm(entity, preview)
                self.assertEqual(response.status_code, 200)
                self.assertEqual(self.ids(entity), before - set(chosen))
                self.assertIn(remaining, self.ids(entity))
                self.assertIn(sentinel, self.ids(other))
                query = parse_qs(urlsplit(response.history[0].location).query)
                self.assertEqual(query, {"vista": [entity], **{key: [value] for key, value in filters.items()}})

    def test_all_respects_date_boundaries_in_all_fifteen_modules(self):
        for entity in web.ENTITY_MODEL:
            with self.subTest(entity=entity):
                chosen = [self.create_record(entity, "2026-09-01", "00:00:00"), self.create_record(entity, "2026-09-02", "23:59:59")]
                outside = self.create_record(entity, "2026-09-03", "00:00:00")
                other = "eventos" if entity == "censo" else "censo"
                sentinel = self.create_record(other, "2027-01-01")
                preview = self.preview(entity, "all", filters={"from": "2026-09-01", "to": "2026-09-02"})
                self.assertIn("Confirmar eliminación de 2 registros", preview.get_data(as_text=True))
                response = self.confirm(entity, preview)
                self.assertEqual(response.status_code, 200)
                self.assertTrue(set(chosen).isdisjoint(self.ids(entity)))
                self.assertIn(outside, self.ids(entity))
                self.assertIn(sentinel, self.ids(other))

    def test_all_without_filters_clears_only_current_module_including_undated_rows(self):
        for entity, model in web.ENTITY_MODEL.items():
            with self.subTest(entity=entity):
                self.create_record(entity)
                rid = self.create_record(entity, "2027-01-01")
                date_name = web.ENTITY_DATE_FIELD.get(entity, "fecha")
                if model.__table__.columns[date_name].nullable:
                    with self.sessions() as db:
                        setattr(db.get(model, rid), date_name, None)
                        db.commit()
                other = "eventos" if entity == "censo" else "censo"
                self.create_record(other)
                before_other = {key: self.ids(key) for key in web.ENTITY_MODEL if key != entity}
                preview = self.preview(entity, "all")
                self.assertIn("Todo el historial de este módulo", preview.get_data(as_text=True))
                self.assertEqual(self.confirm(entity, preview).status_code, 200)
                self.assertEqual(self.ids(entity), set())
                for key, expected in before_other.items():
                    self.assertEqual(self.ids(key), expected)

    def test_week_filter_and_datetime_end_of_day_are_preserved(self):
        start, end = web.week_range(88)
        first = self.create_record("encuestas", start.isoformat(), "00:00:00")
        last = self.create_record("encuestas", end.isoformat(), "23:59:59")
        outside = self.create_record("encuestas", (end + timedelta(days=1)).isoformat(), "00:00:00")
        preview = self.preview("encuestas", "all", filters={"semana": "88"})
        response = self.confirm("encuestas", preview)
        self.assertEqual(self.ids("encuestas"), {outside})
        self.assertTrue({first, last}.isdisjoint(self.ids("encuestas")))
        query = parse_qs(urlsplit(response.history[0].location).query)
        self.assertEqual(query["semana"], ["88"])
        self.assertEqual(query["from"], [start.isoformat()])
        self.assertEqual(query["to"], [end.isoformat()])

    def test_empty_invalid_and_out_of_scope_selections_never_become_delete_all(self):
        visible = self.create_record("censo")
        outside = self.create_record("censo", "2027-01-01")
        cases = [("selected", []), ("selected", [outside]), ("selected", ["bad-id"]), ("unexpected", [visible])]
        for mode, selected in cases:
            with self.subTest(mode=mode, ids=selected):
                response = self.preview("censo", mode, selected, {"from": "2026-09-01", "to": "2026-09-10"})
                self.assertIn(response.status_code, (302, 400))
                self.assertEqual(self.ids("censo"), {visible, outside})
        page = self.client.get("/gestion-5s/registros?vista=censo")
        data = FormValues(page.get_data(as_text=True)).values
        data.update(mode="all", **{"from": "fecha-inválida"})
        self.assertEqual(self.client.post("/gestion-5s/delete/censo/bulk/confirm", data=data).status_code, 400)
        self.assertEqual(self.ids("censo"), {visible, outside})

    def test_confirmation_cancel_and_get_requests_do_not_delete(self):
        rid = self.create_record("desviaciones")
        preview = self.preview("desviaciones", "selected", [rid])
        self.assertIn("Cancelar", preview.get_data(as_text=True))
        self.assertEqual(self.client.get("/gestion-5s/registros?vista=desviaciones").status_code, 200)
        self.assertEqual(self.client.get("/gestion-5s/delete/desviaciones/bulk").status_code, 405)
        self.assertEqual(self.ids("desviaciones"), {rid})

    def test_new_records_after_preview_are_not_deleted_and_replay_does_not_delete_them(self):
        first = self.create_record("censo")
        preview = self.preview("censo", "all")
        later = self.create_record("censo")
        self.assertEqual(self.confirm("censo", preview).status_code, 200)
        self.assertEqual(self.ids("censo"), {later})
        self.assertNotIn(first, self.ids("censo"))
        self.assertEqual(self.confirm("censo", preview).status_code, 200)
        self.assertEqual(self.ids("censo"), {later})

    def test_a_changed_record_cancels_the_whole_pending_deletion(self):
        ids = [self.create_record("desviaciones"), self.create_record("desviaciones")]
        preview = self.preview("desviaciones", "all")
        with self.sessions() as db:
            db.get(web.DesviacionEntry, ids[0]).acciones = "Dato actualizado después de revisar el listado"
            db.commit()
        response = self.confirm("desviaciones", preview)
        self.assertEqual(response.status_code, 200)
        self.assertIn("No se eliminó ninguno", response.get_data(as_text=True))
        self.assertEqual(self.ids("desviaciones"), set(ids))

    def test_bulk_failure_rolls_back_every_batch(self):
        with self.sessions() as db:
            db.add_all([web.CensusEntry(fecha=date(2026, 9, 1)) for _ in range(501)])
            db.commit()
        with self.engine.begin() as conn:
            conn.execute(text("CREATE TRIGGER refuse_last_census BEFORE DELETE ON census_entries WHEN OLD.id = 501 BEGIN SELECT RAISE(ABORT, 'test rollback'); END"))
        preview = self.preview("censo", "all")
        with self.assertLogs(web.app.logger, level="ERROR"):
            response = self.confirm("censo", preview)
        self.assertEqual(response.status_code, 200)
        self.assertIn("No se guardó ningún borrado", response.get_data(as_text=True))
        self.assertEqual(len(self.ids("censo")), 501)

    def test_csrf_invalid_tokens_expiration_and_wrong_module_cannot_delete(self):
        rid = self.create_record("censo")
        preview = self.preview("censo", "all")
        data = FormValues(preview.get_data(as_text=True)).values
        no_csrf = {key: value for key, value in data.items() if key != "csrf_token"}
        self.assertEqual(self.client.post("/gestion-5s/delete/censo/bulk", data=no_csrf).status_code, 400)
        self.assertEqual(self.client.post("/gestion-5s/delete/censo/bulk/confirm", data={"mode": "all"}).status_code, 400)
        self.assertEqual(self.client.post("/gestion-5s/delete/eventos/bulk", data=data).status_code, 400)
        tampered = {**data, "confirmation_token": data["confirmation_token"] + "alterado"}
        self.assertEqual(self.client.post("/gestion-5s/delete/censo/bulk", data=tampered).status_code, 302)
        serializer = web.bulk_delete_serializer()
        payload = serializer.loads(data["confirmation_token"])
        with patch("itsdangerous.timed.TimestampSigner.get_timestamp", return_value=int(time.time()) - 901):
            expired = {**data, "confirmation_token": serializer.dumps(payload)}
        self.assertEqual(self.client.post("/gestion-5s/delete/censo/bulk", data=expired).status_code, 302)
        self.assertEqual(self.ids("censo"), {rid})


if __name__ == "__main__":
    unittest.main()
