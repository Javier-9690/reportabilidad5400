import csv
import html
import re
import time
import unicodedata
import unittest
from datetime import timedelta
from io import StringIO
from unittest.mock import patch
from urllib.parse import parse_qs, urlencode, urlsplit

from sqlalchemy import String, create_engine
from sqlalchemy.dialects import postgresql
from sqlalchemy.orm import sessionmaker

from tests.test_integration import FORM_DATA, app
from tests.test_edit_records import FormValues
from tests.test_form_structure import PageStructure
from gestion5s import web
from gestion5s.editing import list_fields
from gestion5s.searching import record_search_condition


class RecordSearchTest(unittest.TestCase):
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

    def create(self, entity, **values):
        tab = "encuesta" if entity == "encuestas" else entity
        response = self.client.post(f"/gestion-5s/panel?tab={tab}", data={**FORM_DATA[tab], **values})
        self.assertEqual(response.status_code, 302, response.get_data(as_text=True))
        with self.sessions() as db:
            model = web.ENTITY_MODEL[entity]
            return db.query(model).order_by(model.id.desc()).first().id

    def stored_ids(self, entity):
        with self.sessions() as db:
            return {record.id for record in db.query(web.ENTITY_MODEL[entity]).all()}

    def listing(self, entity, q="", **filters):
        response = self.client.get("/gestion-5s/registros", query_string={"vista": entity, "q": q, **filters})
        self.assertEqual(response.status_code, 200, response.get_data(as_text=True))
        return response.get_data(as_text=True)

    def listed_ids(self, entity, q="", **filters):
        return {int(rid) for rid in re.findall(
            r'class="form-check-input record-select" name="ids" value="(\d+)"',
            self.listing(entity, q, **filters),
        )}

    def exported(self, entity, q="", **filters):
        response = self.client.get(f"/gestion-5s/download/{entity}.csv", query_string={"q": q, **filters})
        self.assertEqual(response.status_code, 200, response.get_data(as_text=True))
        reader = csv.DictReader(StringIO(response.data.decode("utf-8-sig")))
        self.assertTrue(reader.fieldnames)
        return list(reader)

    def preview(self, entity, q, mode="all", ids=(), **filters):
        data = FormValues(self.listing(entity, q, **filters)).values
        self.assertEqual(data["q"], q.strip())
        data.update(mode=mode, ids=[str(rid) for rid in ids])
        return self.client.post(f"/gestion-5s/delete/{entity}/bulk/confirm", data=data)

    def marker_field(self, entity):
        if entity == "censo":
            return "censo_dia"
        if entity == "atencion":
            return "cantidad"
        model = web.ENTITY_MODEL[entity]
        return next(field["name"] for field in list_fields(entity, model())
                    if isinstance(getattr(model, field["name"]).type, String))

    def test_partial_numbers_match_only_current_category_in_list_and_csv_for_all_twenty(self):
        for entity in web.ENTITY_MODEL:
            with self.subTest(entity=entity):
                field = self.marker_field(entity)
                expected = {self.create(entity, **{field: marker}) for marker in ("730019", "7300199")}
                self.create(entity, **{field: "640029"})
                self.assertEqual(self.listed_ids(entity, "30019"), expected)
                rows = self.exported(entity, "30019")
                self.assertEqual(len(rows), 2)
                self.assertEqual({value for row in rows for value in row.values()
                                  if value in ("730019", "7300199", "640029")}, {"730019", "7300199"})
                page = self.listing(entity, "30019")
                self.assertIn('id="recordSearch"', page)
                self.assertIn('for="recordSearch"', page)
                self.assertIn('type="search"', page)
                self.assertIn('2 coincidencias', page)
                download = next(link for link in PageStructure(page).links if "/download/" in link)
                self.assertEqual(parse_qs(urlsplit(download).query)["q"], ["30019"])

    def test_words_ignore_case_and_spanish_accents_including_decomposed_text(self):
        phrase = "ÁRBOL ÉL ÍNDIGO ÓPERA ÚTIL PINGÜINO ÑANDÚ"
        expected = {self.create("samtech_usuarios", comentario=phrase),
                    self.create("samtech_usuarios", falla=unicodedata.normalize("NFD", phrase))}
        self.create("samtech_usuarios", comentario="Sin coincidencia")
        self.create("miscelaneo", comentario=phrase)
        for term in ("arbol", "él índigo", "oPeRa", "util", "PINGUINO", "ñandú", "nandu",
                     unicodedata.normalize("NFD", "ÑANDÚ")):
            with self.subTest(term=term):
                self.assertEqual(self.listed_ids("samtech_usuarios", term), expected)
                self.assertEqual(len(self.exported("samtech_usuarios", term)), 2)

    def test_symbols_are_literal_and_search_text_is_escaped_in_html(self):
        cases = ("50%", "A_B", "ruta/archivo", "' OR 1=1 --", '<script>alert("x")</script>')
        ids = {term: self.create("samtech_usuarios", comentario=term) for term in cases}
        self.create("samtech_usuarios", comentario="50X A-B rutaXarchivo")
        for term, rid in ids.items():
            with self.subTest(term=term):
                self.assertEqual(self.listed_ids("samtech_usuarios", term), {rid})
                self.assertEqual(len(self.exported("samtech_usuarios", term)), 1)
        self.assertEqual(self.listed_ids("samtech_usuarios", "%"), {ids["50%"]})
        self.assertEqual(self.listed_ids("samtech_usuarios", "_"), {ids["A_B"]})
        page = self.listing("samtech_usuarios", cases[-1])
        self.assertNotIn(cases[-1], page)
        self.assertEqual(FormValues(page).values["q"], cases[-1])

    def test_each_samtech_field_and_deviation_actions_are_searchable_with_blank_dates(self):
        for field in list_fields("samtech_usuarios", web.SamtechUsuarioEntry()):
            name = field["name"]
            data = {key: "" for key in FORM_DATA["samtech_usuarios"]}
            term = "2037-12-23" if field["kind"] == "date" else f"Dato exclusivo {name}"
            rid = self.create("samtech_usuarios", **{**data, name: term})
            with self.subTest(field=name):
                self.assertIn(rid, self.listed_ids("samtech_usuarios", term))
                self.assertTrue(self.exported("samtech_usuarios", term))
        rid = self.create("desviaciones", acciones="Coordinar reparación eléctrica")
        self.assertEqual(self.listed_ids("desviaciones", "REPARACION ELECTRICA"), {rid})
        self.assertEqual(len(self.exported("desviaciones", "reparacion")), 1)

    def test_dates_times_durations_zero_and_computed_numbers_use_displayed_values(self):
        rid = self.create("samtech_usuarios", fecha_aprobacion="2037-12-23")
        for term in ("2037-12-23", "23/12/2037", "23/12"):
            self.assertEqual(self.listed_ids("samtech_usuarios", term), {rid})
        rid = self.create("entradas_salidas", hora_salida="19:43:27")
        self.assertEqual(self.listed_ids("entradas_salidas", "19:43"), {rid})
        for entity, field in (("atencion", "tiempo_promedio_sec"), ("solicitud_ot", "tiempo_respuesta_sec")):
            for term in ("00:00", "03:54", "125:59"):
                rid = self.create(entity, **{field: term})
                self.assertEqual(self.listed_ids(entity, term), {rid})
                self.assertEqual(len(self.exported(entity, term)), 1)
        rid = self.create("alarmas", aviso_mantencion_h="0", llegada_lider_h="1.25")
        for term in ("0.0", "1.25", "1,25"):
            self.assertEqual(self.listed_ids("alarmas", term), {rid})
        rid = self.create("encuestas", q1_puntaje="3", q2_puntaje="4")
        for term in ("7", "3.5", "3,5"):
            self.assertEqual(self.listed_ids("encuestas", term), {rid})
        # Los identificadores técnicos y las columnas retiradas no son datos buscables.
        rid = self.create("extensiones")
        with self.sessions() as db:
            row = db.get(web.ExtensionExcepcionEntry, rid)
            row.id = 1987654321
            row.proyecto = "Proyecto histórico exclusivo"
            db.commit()
        self.assertEqual(self.listed_ids("extensiones", "1987654321"), set())
        self.assertEqual(self.listed_ids("extensiones", "historico exclusivo"), set())

    def test_search_combines_with_dates_and_weeks_including_datetime_day_boundaries(self):
        start, end = web.week_range(88)
        expected = {self.create("encuestas", comentarios="Busca", fecha_hora=f"{start}T00:00:00"),
                    self.create("encuestas", comentarios="Busca", fecha_hora=f"{end}T23:59:59")}
        self.create("encuestas", comentarios="Busca", fecha_hora=f"{end + timedelta(days=1)}T00:00:00")
        self.create("encuestas", comentarios="Otro", fecha_hora=f"{end}T12:00:00")
        for filters in ({"from": start.isoformat(), "to": end.isoformat()}, {"semana": "88"}):
            self.assertEqual(self.listed_ids("encuestas", "busca", **filters), expected)
            self.assertEqual(len(self.exported("encuestas", "busca", **filters)), 2)
        dated = self.create("samtech_usuarios", ticket="077")
        undated = self.create("samtech_usuarios", ticket="077", fecha_creacion="")
        self.assertEqual(self.listed_ids("samtech_usuarios", "077"), {dated, undated})
        self.assertEqual(self.listed_ids("samtech_usuarios", "077", **{"from": "2026-09-01"}), {dated})

    def test_empty_search_keeps_rows_and_no_match_cannot_become_delete_all(self):
        rid = self.create("samtech_usuarios")
        for q in ("", "  \t  "):
            self.assertEqual(self.listed_ids("samtech_usuarios", q), {rid})
            self.assertEqual(len(self.exported("samtech_usuarios", q)), 1)
        self.assertEqual(self.listed_ids("samtech_usuarios", "inexistente"), set())
        self.assertEqual(self.exported("samtech_usuarios", "inexistente"), [])
        page = self.listing("samtech_usuarios", "inexistente")
        self.assertIn("No se encontraron coincidencias", page)
        self.assertRegex(page, r'<button id="deleteAllRecords"[^>]*disabled')
        self.assertEqual(self.preview("samtech_usuarios", "inexistente").status_code, 302)
        self.assertEqual(self.stored_ids("samtech_usuarios"), {rid})

    def test_clear_search_keeps_category_dates_week_and_edit_delete_return_keeps_query(self):
        entity = "samtech_usuarios"
        term = "Árbol & 007"
        rid = self.create(entity, ticket=term)
        start, end = web.week_range(88)
        filters = {"from": start.isoformat(), "to": end.isoformat(), "semana": "88"}
        page = self.listing(entity, term, **filters)
        clear_link = html.unescape(re.search(r'href="([^"]+)"[^>]*>Quitar búsqueda</a>', page).group(1))
        self.assertEqual(parse_qs(urlsplit(clear_link).query), {"vista": [entity], **{k: [v] for k, v in filters.items()}})
        editor = next(link for link in PageStructure(page).links if "/edit/" in link)
        response = self.client.get(editor)
        self.assertEqual(response.status_code, 200)
        data = FormValues(response.get_data(as_text=True)).values
        expected = {"vista": [entity], "q": [term], **{k: [v] for k, v in filters.items()}}
        self.assertEqual(parse_qs(urlsplit(data["next"]).query), expected)
        self.assertIn(data["next"], PageStructure(response.get_data(as_text=True)).links)
        data["estado"] = "Actualizado"
        response = self.client.post(f"/gestion-5s/edit/{entity}/{rid}", data=data)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(parse_qs(urlsplit(response.location).query), expected)
        response = self.client.post(f"/gestion-5s/delete/{entity}/{rid}", data={"next": response.location})
        self.assertEqual(response.status_code, 302)
        self.assertEqual(urlsplit(response.location).path, "/gestion-5s/registros")
        self.assertEqual(parse_qs(urlsplit(response.location).query), expected)
        self.assertEqual(self.client.get(response.location).status_code, 200)

    def test_delete_all_respects_search_date_and_category_for_all_twenty(self):
        for entity in web.ENTITY_MODEL:
            with self.subTest(entity=entity):
                field = self.marker_field(entity)
                chosen = self.create(entity, **{field: "730019"})
                nonmatch = self.create(entity, **{field: "640029"})
                date_field = web.ENTITY_DATE_FIELD.get(entity, "fecha")
                outside = self.create(entity, **{field: "730019", date_field:
                                      "2027-01-01T12:00:00" if date_field == "fecha_hora" else "2027-01-01"})
                others = {key: self.stored_ids(key) for key in web.ENTITY_MODEL if key != entity}
                preview = self.preview(entity, "730019", **{"from": "2026-09-01", "to": "2026-09-02"})
                self.assertEqual(preview.status_code, 200)
                self.assertIn("Confirmar eliminación de 1 registro", preview.get_data(as_text=True))
                self.assertIn("<strong>Búsqueda:</strong> 730019", preview.get_data(as_text=True))
                data = FormValues(preview.get_data(as_text=True)).values
                self.assertEqual(web.bulk_delete_serializer().loads(data["confirmation_token"])["filters"]["q"], "730019")
                data["q"] = ""  # Solo se utiliza el alcance confirmado y firmado.
                response = self.client.post(f"/gestion-5s/delete/{entity}/bulk", data=data)
                self.assertEqual(response.status_code, 302)
                self.assertEqual(parse_qs(urlsplit(response.location).query)["q"], ["730019"])
                self.assertEqual(self.stored_ids(entity), {nonmatch, outside})
                self.assertNotIn(chosen, self.stored_ids(entity))
                for key, expected in others.items():
                    self.assertEqual(self.stored_ids(key), expected)

    def test_selected_outside_search_is_rejected_and_literal_percent_deletes_only_matches(self):
        chosen = self.create("samtech_usuarios", comentario="50%")
        nonmatch = self.create("samtech_usuarios", comentario="50X")
        response = self.preview("samtech_usuarios", "%", mode="selected", ids=[chosen, nonmatch])
        self.assertEqual(response.status_code, 302)
        self.assertEqual(self.stored_ids("samtech_usuarios"), {chosen, nonmatch})
        preview = self.preview("samtech_usuarios", "%", mode="selected", ids=[chosen])
        self.assertEqual(preview.status_code, 200)
        self.assertIn("Sin filtro de fechas", preview.get_data(as_text=True))
        self.assertNotIn("Todo el historial de este módulo", preview.get_data(as_text=True))
        data = FormValues(preview.get_data(as_text=True)).values
        later = self.create("samtech_usuarios", comentario="50%")
        response = self.client.post("/gestion-5s/delete/samtech_usuarios/bulk", data=data)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(self.stored_ids("samtech_usuarios"), {nonmatch, later})

    def test_changed_search_match_and_expired_confirmation_preserve_records_and_query(self):
        rid = self.create("samtech_usuarios", ticket="077")
        preview = self.preview("samtech_usuarios", "077")
        data = FormValues(preview.get_data(as_text=True)).values
        with self.sessions() as db:
            db.get(web.SamtechUsuarioEntry, rid).ticket = "088"
            db.commit()
        response = self.client.post("/gestion-5s/delete/samtech_usuarios/bulk", data=data, follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("No se eliminó ninguno", response.get_data(as_text=True))
        self.assertEqual(self.stored_ids("samtech_usuarios"), {rid})
        serializer = web.bulk_delete_serializer()
        payload = serializer.loads(data["confirmation_token"])
        with patch("itsdangerous.timed.TimestampSigner.get_timestamp", return_value=int(time.time()) - 901):
            data["confirmation_token"] = serializer.dumps(payload)
        response = self.client.post("/gestion-5s/delete/samtech_usuarios/bulk", data=data)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(parse_qs(urlsplit(response.location).query), {"vista": ["samtech_usuarios"], "q": ["077"]})
        data["next"] = "https://example.org/gestion-5s/registros?" + urlencode({"q": "077"})
        response = self.client.post("/gestion-5s/delete/samtech_usuarios/bulk", data=data)
        self.assertEqual(response.location, "/gestion-5s/registros?vista=samtech_usuarios")

    def test_invalid_queries_are_rejected_without_broadening_or_deleting(self):
        rid = self.create("samtech_usuarios")
        data = FormValues(self.listing("samtech_usuarios")).values
        for q in ("x" * 201, "\x00", "\u0301\u0303\u0308"):
            with self.subTest(q=q):
                for path in ("/gestion-5s/registros?vista=samtech_usuarios", "/gestion-5s/download/samtech_usuarios.csv"):
                    response = self.client.get(path + ("&" if "?" in path else "?") + urlencode({"q": q}))
                    self.assertEqual(response.status_code, 400)
                response = self.client.post("/gestion-5s/delete/samtech_usuarios/bulk/confirm", data={**data, "q": q, "mode": "all"})
                self.assertEqual(response.status_code, 400)
                self.assertEqual(self.stored_ids("samtech_usuarios"), {rid})

    def test_postgresql_search_uses_builtin_functions_and_escaped_bound_parameters(self):
        # Compilación SQL de producción; no sustituye una prueba con PostgreSQL real.
        for entity, model in web.ENTITY_MODEL.items():
            condition = record_search_condition(entity, model, "ñandú%_", "postgresql")
            compiled = condition.compile(dialect=postgresql.dialect())
            sql = str(compiled)
            self.assertIn("to_char(", sql)
            self.assertIn("ESCAPE", sql)
            self.assertIn("nandu/%/_", compiled.params.values())
            self.assertNotIn("unaccent(", sql)
            self.assertNotIn("strftime(", sql)
            if entity not in ("censo", "atencion"):
                self.assertIn("translate(", sql)


if __name__ == "__main__":
    unittest.main()
