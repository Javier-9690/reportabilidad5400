import unittest
from urllib.parse import parse_qs, urlsplit

from tests.test_integration import FORM_DATA, app
from tests.test_edit_records import FormValues
from gestion5s.web import ENTITY_MODEL, SessionLocal


class DeleteHotelRecordsTest(unittest.TestCase):
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

    def test_each_delete_form_returns_to_its_filtered_listing(self):
        for entity in ENTITY_MODEL:
            with self.subTest(entity=entity):
                rid = self.create_record(entity)
                filters = {"vista": entity, "from": "2026-09-01", "to": "2026-09-30"}
                page = self.client.get("/gestion-5s/registros", query_string=filters)
                self.assertEqual(page.status_code, 200)
                values = FormValues(page.get_data(as_text=True)).values
                self.assertEqual(urlsplit(values["next"]).path, "/gestion-5s/registros")

                response = self.client.post(
                    f"/gestion-5s/delete/{entity}/{rid}",
                    data={"next": values["next"]}, follow_redirects=True,
                )
                self.assertEqual(response.status_code, 200)
                self.assertEqual(len(response.history), 1)
                destination = urlsplit(response.history[0].location)
                self.assertEqual(destination.path, "/gestion-5s/registros")
                self.assertEqual(parse_qs(destination.query), {key: [value] for key, value in filters.items()})
                self.assertIn("Registro eliminado.", response.get_data(as_text=True))
                with SessionLocal() as db:
                    self.assertIsNone(db.get(ENTITY_MODEL[entity], rid))

    def test_legacy_unprefixed_form_keeps_filters_and_handles_repeated_delete(self):
        entity = "desviaciones"
        rid = self.create_record(entity)
        next_url = "/registros?vista=desviaciones&from=2026-09-01&to=2026-09-10&semana=88"
        for attempt in range(2):
            with self.subTest(attempt=attempt):
                response = self.client.post(
                    f"/gestion-5s/delete/{entity}/{rid}", data={"next": next_url}, follow_redirects=True,
                )
                self.assertEqual(response.status_code, 200)
                destination = urlsplit(response.history[0].location)
                self.assertEqual(destination.path, "/gestion-5s/registros")
                self.assertEqual(parse_qs(destination.query), parse_qs(urlsplit(next_url).query))
                message = "Registro eliminado." if attempt == 0 else "Registro no encontrado."
                self.assertIn(message, response.get_data(as_text=True))

    def test_missing_or_invalid_return_target_stays_in_the_selected_module(self):
        for target in (None, "https://example.org/registros", "//example.org/registros", "/ruta-inexistente"):
            with self.subTest(target=target):
                rid = self.create_record("reclamos")
                response = self.client.post(
                    f"/gestion-5s/delete/reclamos/{rid}",
                    data={} if target is None else {"next": target}, follow_redirects=True,
                )
                self.assertEqual(response.status_code, 200)
                self.assertEqual(response.history[0].location, "/gestion-5s/registros?vista=reclamos")
                with SessionLocal() as db:
                    self.assertIsNone(db.get(ENTITY_MODEL["reclamos"], rid))


if __name__ == "__main__":
    unittest.main()
