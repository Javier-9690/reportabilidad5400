import unittest
from collections import Counter
from datetime import date, datetime, time
from html.parser import HTMLParser
from io import BytesIO
from unittest.mock import patch
from urllib.parse import parse_qs, urlsplit

from sqlalchemy import Date, DateTime, Float, Integer, Text, Time, create_engine
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.orm import sessionmaker

from tests.test_integration import FORM_DATA, app
from tests.test_edit_records import FormValues
from gestion5s import web


class PageStructure(HTMLParser):
    def __init__(self, page):
        super().__init__(convert_charrefs=True)
        self.forms, self.ids, self.labels, self.links, self.cells = [], [], [], [], {}
        self.current_form = None
        self.nested_forms = 0
        self.cell = None
        self.feed(page)

    def handle_starttag(self, tag, attributes):
        attrs = dict(attributes)
        if attrs.get("id"):
            self.ids.append(attrs["id"])
        if tag == "label":
            self.labels.append(attrs)
        if tag == "a":
            self.links.append(attrs.get("href", ""))
        if tag == "form":
            self.nested_forms += self.current_form is not None
            self.current_form = {"action": attrs.get("action", ""), "controls": []}
            self.forms.append(self.current_form)
        elif tag in ("input", "textarea", "select") and self.current_form is not None:
            if attrs.get("name"):
                self.current_form["controls"].append(attrs)
        elif tag == "td" and attrs.get("data-field"):
            self.cell = attrs["data-field"]
            self.cells[self.cell] = ""

    def handle_data(self, data):
        if self.cell:
            self.cells[self.cell] += data

    def handle_endtag(self, tag):
        if tag == "form":
            self.current_form = None
        elif tag == "td":
            self.cell = None


def active_columns(entity):
    excluded = {"id", "creado"}
    if entity == "extensiones":
        excluded.add("proyecto")  # Campo histórico retirado por solicitud del usuario.
    return [column for column in web.ENTITY_MODEL[entity].__table__.columns if column.name not in excluded]


def complete_form(entity):
    data = {}
    for column in active_columns(entity):
        name = column.name
        if entity == "encuestas" and name in ("total", "promedio"):
            continue
        if name in ("tiempo_promedio_sec", "tiempo_respuesta_sec"):
            value = "125:59"
        elif isinstance(column.type, DateTime):
            value = "2026-09-09T10:15:42"
        elif isinstance(column.type, Date):
            value = "2026-09-09"
        elif isinstance(column.type, Time):
            value = "00:00:00"
        elif isinstance(column.type, Integer):
            value = "89" if name == "semana" else ("3" if name.endswith("_puntaje") else "0")
        elif isinstance(column.type, Float):
            value = "0"
        elif isinstance(column.type, Text):
            value = (f"{name}: texto completo con acentos, & y <etiquetas>. " * 5) + "\nÚltima línea visible."
        elif name in ("id_interno", "habitacion", "n_habitacion", "n_tarjeta"):
            value = "000007"
        else:
            value = f"Dato-{name}-á-007"
        data[name] = value
    return data


class FormStructureTest(unittest.TestCase):
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

    def records(self, entity):
        with self.sessions() as db:
            return [{column.name: getattr(record, column.name) for column in record.__table__.columns}
                    for record in db.query(web.ENTITY_MODEL[entity]).all()]

    def create(self, entity, data):
        response = self.client.post("/gestion-5s/panel", query_string={"tab": entity}, data=data)
        self.assertEqual(response.status_code, 302, response.get_data(as_text=True))
        return self.records(entity)[-1]

    def test_all_forms_have_unique_controls_labels_and_a_working_view_records_link(self):
        for entity in web.ENTITY_MODEL:
            with self.subTest(entity=entity):
                response = self.client.get("/gestion-5s/panel", query_string={"tab": entity})
                self.assertEqual(response.status_code, 200)
                page = response.get_data(as_text=True)
                structure = PageStructure(page)
                self.assertEqual(structure.nested_forms, 0)
                self.assertEqual(len(structure.ids), len(set(structure.ids)))
                form = next(form for form in structure.forms if "/panel" in form["action"])
                controls = [control for control in form["controls"] if control["name"] != "censo_total_auto"]
                names = [control["name"] for control in controls]
                self.assertTrue(all(count == 1 for count in Counter(names).values()))
                columns = [column for column in active_columns(entity)
                           if not (entity == "encuestas" and column.name in ("total", "promedio"))]
                self.assertEqual(set(names), {column.name for column in columns})
                for control in controls:
                    column = next(column for column in columns if column.name == control["name"])
                    self.assertTrue(any(label.get("for") == control["id"] for label in structure.labels))
                    self.assertEqual("required" in control, not column.nullable and column.default is None)
                    if getattr(column.type, "length", None):
                        self.assertEqual(control["maxlength"], str(column.type.length))
                    if isinstance(column.type, Integer) and not column.name.endswith("_sec"):
                        self.assertEqual(control["type"], "number")
                        self.assertEqual(control["step"], "1")
                        self.assertIn("max", control)
                target = f"/gestion-5s/registros?vista={entity}"
                self.assertEqual(structure.links.count(target), 1)
                self.assertIn("Ver registros", page)
                self.assertEqual(self.client.get(target).status_code, 200)
                template = next(link for link in structure.links if "/template/" in link)
                self.assertEqual(self.client.get(template).status_code, 200)

    def test_every_business_field_is_saved_and_fully_visible_in_all_modules(self):
        for entity in web.ENTITY_MODEL:
            with self.subTest(entity=entity):
                data = complete_form(entity)
                record = self.create(entity, data)
                listing = self.client.get("/gestion-5s/registros", query_string={"vista": entity})
                self.assertEqual(listing.status_code, 200)
                page = listing.get_data(as_text=True)
                structure = PageStructure(page)
                self.assertEqual(set(structure.cells), {column.name for column in active_columns(entity)})
                for name, value in data.items():
                    stored = record[name]
                    if name.endswith("_sec"):
                        self.assertEqual(stored, 7559)
                        self.assertEqual(structure.cells[name], "125:59")
                    elif isinstance(stored, datetime):
                        self.assertEqual(stored, datetime.fromisoformat(value))
                        self.assertEqual(structure.cells[name], str(stored))
                    elif isinstance(stored, (date, time)):
                        self.assertEqual(str(stored), value)
                        self.assertEqual(structure.cells[name], value)
                    elif isinstance(stored, (int, float)):
                        self.assertEqual(stored, float(value))
                        self.assertEqual(float(structure.cells[name]), float(value))
                    else:
                        self.assertEqual(stored, value)
                        self.assertEqual(structure.cells[name], value)
                        self.assertNotIn("<etiquetas>", page)
                if entity == "encuestas":
                    self.assertEqual(structure.cells["total"], "15")
                    self.assertEqual(float(structure.cells["promedio"]), 3)
                self.assertIn('class="table-responsive records-scroll"', page)
                self.assertIn('data-record-scroll-top hidden', page)
                self.assertIn('tabindex="0" role="region"', page)
                self.assertEqual(structure.nested_forms, 0)
                editor = self.client.get(f"/gestion-5s/edit/{entity}/{record['id']}")
                self.assertEqual(editor.status_code, 200)
                editor_values = FormValues(editor.get_data(as_text=True)).values
                for name, value in data.items():
                    expected = str(float(value)) if isinstance(record[name], float) else value
                    self.assertEqual(editor_values[name], expected)

    def test_missing_required_fields_return_errors_without_saving_in_every_module(self):
        for entity in web.ENTITY_MODEL:
            data = complete_form(entity)
            for column in active_columns(entity):
                if column.nullable or column.default is not None:
                    continue
                with self.subTest(entity=entity, field=column.name):
                    incomplete = {key: value for key, value in data.items() if key != column.name}
                    response = self.client.post(f"/gestion-5s/panel?tab={entity}", data=incomplete)
                    self.assertEqual(response.status_code, 422)
                    self.assertIn(f'entry-{column.name}-error', response.get_data(as_text=True))
                    self.assertEqual(self.records(entity), [])

    def test_invalid_values_keep_the_entered_data_and_never_save(self):
        cases = [
            ("censo", "censo_dia", "1.5"), ("censo", "censo_noche", "-1"),
            ("censo", "total", "2147483648"), ("duplicidades", "semana", "cualquier semana"),
            ("eventos", "horario", "x" * 51), ("encuestas", "q3_puntaje", "6"),
            ("encuestas", "q2_puntaje", "texto"), ("atencion", "tiempo_promedio_sec", "03:99"),
            ("solicitud_ot", "tiempo_respuesta_sec", "sin hora"), ("robos", "hora", "25:10"),
            ("alarmas", "aviso_mantencion_h", "incorrecto"), ("alarmas", "llegada_lider_h", "NaN"),
            ("alarmas", "hora_reporte_salfa", "24:99"), ("extensiones", "cant_clientes", "1.5"),
            ("onboarding", "fecha_hora", "2026-09-09T25:01"),
        ]
        cases += [(entity, web.ENTITY_DATE_FIELD.get(entity, "fecha"), "2026-02-30") for entity in web.ENTITY_MODEL]
        for entity, name, invalid in cases:
            with self.subTest(entity=entity, field=name, invalid=invalid):
                data = complete_form(entity)
                data[name] = invalid
                response = self.client.post(f"/gestion-5s/panel?tab={entity}", data=data)
                self.assertEqual(response.status_code, 422)
                page = response.get_data(as_text=True)
                self.assertIn(f'entry-{name}-error', page)
                self.assertEqual({key: FormValues(page).values[key] for key in data}, data)
                self.assertEqual(self.records(entity), [])

    def test_database_failure_preserves_the_form_and_allows_a_single_retry(self):
        data = complete_form("alarmas")
        session = self.sessions()
        with patch.object(web, "SessionLocal", return_value=session), \
                patch.object(session, "commit", side_effect=SQLAlchemyError("simulated failure")), \
                self.assertLogs(web.app.logger, level="ERROR"):
            response = self.client.post("/gestion-5s/panel?tab=alarmas", data=data)
        self.assertEqual(response.status_code, 503)
        self.assertEqual({key: FormValues(response.get_data(as_text=True)).values[key] for key in data}, data)
        self.assertEqual(self.records("alarmas"), [])
        self.create("alarmas", data)
        self.assertEqual(len(self.records("alarmas")), 1)

    def test_legacy_field_names_and_pdf_names_are_preserved(self):
        for entity in ("duplicidades", "desviaciones", "reclamos"):
            record = self.create(entity, dict(FORM_DATA[entity], id="000099"))
            self.assertEqual(record["id_interno"], "000099")
        for entity, field in (("atencion", "tiempo_promedio"), ("solicitud_ot", "tiempo_respuesta")):
            record = self.create(entity, dict(FORM_DATA[entity], **{field: "03:54"}))
            self.assertEqual(record[field + "_sec"], 234)
        data = dict(FORM_DATA["onboarding"], archivo_pdf=(BytesIO(b"%PDF-1.4"), "Constancia á.pdf"))
        record = self.create("onboarding", data)
        self.assertEqual(record["archivo_pdf"], "Constancia á.pdf")

    def test_zero_durations_and_missing_values_are_distinct_and_filters_stay_in_module(self):
        for entity, name in (("atencion", "tiempo_promedio_sec"), ("solicitud_ot", "tiempo_respuesta_sec")):
            self.create(entity, dict(complete_form(entity), **{name: "00:00"}))
            page = self.client.get(f"/gestion-5s/registros?vista={entity}").get_data(as_text=True)
            self.assertEqual(PageStructure(page).cells[name], "00:00")
        for entity in web.ENTITY_MODEL:
            page = self.client.get(f"/gestion-5s/registros?vista={entity}&from=2026-09-09").get_data(as_text=True)
            links = PageStructure(page).links
            self.assertIn(f"/gestion-5s/registros?vista={entity}", links)
        response = self.client.get("/gestion-5s/panel?tab=encuestas")
        survey = next(form for form in PageStructure(response.get_data(as_text=True)).forms if "/panel" in form["action"])
        self.assertEqual(parse_qs(urlsplit(survey["action"]).query), {"tab": ["encuesta"]})
        self.assertEqual(self.client.get("/gestion-5s/panel?tab=inexistente").status_code, 404)


if __name__ == "__main__":
    unittest.main()
