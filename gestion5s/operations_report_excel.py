"""Descarga del reporte desde Flask, con el motor Excel ya instalado en la app."""

from datetime import date, datetime, time
from memory_utils import disk_workbook

from xlsxwriter.utility import xl_col_to_name
from openpyxl.utils.datetime import to_excel

from gestion5s.operations_reports import STATE_LABELS
from gestion5s.user_report_excel import count_axis_step


@disk_workbook
def export_operations_report(report, book):
    categories = report["categories"]
    for category in categories:
        if len(category["details"]) != category["totals"]["records"]:
            raise ValueError("El reporte no contiene el detalle completo. Vuelve a generar la descarga.")
        if len(category["details"]) > 1048570:
            raise ValueError("El detalle supera las filas de Excel. Selecciona un período más corto.")
        if any(isinstance(value, str) and len(value) > 32767
               for record in category["details"] for value in record.values()):
            raise ValueError(f"Un texto de {category['label']} supera los 32.767 caracteres que admite Excel por celda.")
    book.set_properties({"title": report["title"], "author": "Reportabilidad 5400",
                         "comments": "Generado desde los registros del sistema con los filtros indicados."})
    base = {"font_name": "Calibri", "font_size": 11, "valign": "vcenter"}
    def style(**args):
        return book.add_format({**base, **args})
    f = {
        "title": style(bold=True, font_size=19, font_color="#FFFFFF", bg_color="#9D261E", text_wrap=True),
        "subtitle": style(font_color="#536477", text_wrap=True),
        "header": style(bold=True, font_color="#FFFFFF", bg_color="#33485E", text_wrap=True),
        "section": style(bold=True, font_color="#34475C", font_size=13),
        "text": style(text_wrap=True, valign="top"),
        "source": style(text_wrap=True, valign="top", num_format="@"),
        "number": style(num_format="#,##0"),
        "decimal": style(num_format="0.##"),
        "rate": style(num_format="0.0%"),
        "date": style(num_format="dd/mm/yyyy"),
        "time": style(num_format="hh:mm:ss"),
        "datetime": style(num_format="dd/mm/yyyy hh:mm:ss"),
        "duration": style(num_format="[mm]:ss"),
        "total": style(bold=True, bg_color="#E9EFF5", num_format="#,##0"),
        "total_rate": style(bold=True, bg_color="#E9EFF5", num_format="0.0%"),
    }
    summary = book.add_worksheet("Resumen general" if len(categories) > 1 else "Resumen")
    daily = book.add_worksheet("Evolución diaria")
    summary_name = summary.get_name()
    period = f"Período: {report['start']:%d/%m/%Y} al {report['end']:%d/%m/%Y}"
    stamp = f"Generado: {report['generated_at']:%d/%m/%Y %H:%M} UTC"
    search = report["filters"]["search"]
    filter_text = f"Búsqueda: {search or '(sin filtro de texto)'}. Registros sin fecha: {'incluidos' if report['filters']['include_undated'] else 'excluidos'}."
    refs = {}

    def setup(sheet, title, last_col, color="#9D261E"):
        sheet.hide_gridlines(2)
        sheet.set_tab_color(color)
        sheet.merge_range(0, 0, 0, last_col, title, f["title"])
        sheet.set_row(0, 44)
        sheet.merge_range(1, 0, 1, last_col, period + " · " + stamp, f["subtitle"])
        sheet.set_row(1, 28)
        sheet.merge_range(2, 0, 2, last_col, filter_text, f["subtitle"])
        sheet.set_row(2, 36)
        sheet.set_landscape()
        sheet.set_paper(9)
        sheet.set_margins(.3, .3, .45, .45)
        sheet.set_footer("&LReportabilidad 5400&R&P / &N")

    for category in categories:
        fields = [*category["fields"], dict(name="_report_date", label="Fecha usada en reporte", kind="date"),
                  dict(name="_report_state", label="Estado agrupado para reporte", kind="text")]
        sheet = book.add_worksheet(category["sheet"])
        setup(sheet, category["label"], len(fields) - 1, category["color"])
        sheet.merge_range(3, 0, 3, len(fields) - 1,
                          "Fuente: Registros hotelería > " + category["label"] + ". Fecha utilizada: " + category["date_label"] + ".", f["subtitle"])
        note = ("Estado agrupado usa las equivalencias del reporte; se conserva el estado original."
                if category["status"] else "Este registro no tiene un campo de estado. «No aplica» no significa abierto ni cerrado.")
        if category["entity"] == "habitaciones_bloqueadas":
            note += " Las fechas de liberación por área no confirman la liberación final."
        sheet.merge_range(4, 0, 4, len(fields) - 1, note, f["subtitle"])
        sheet.set_row(4, 32)
        for column, field in enumerate(fields):
            sheet.write_string(5, column, field["label"], f["header"])
            sheet.set_column(column, column, 52 if field["kind"] == "textarea" else 23)
        sheet.set_row(5, 38)
        for row, record in enumerate(category["details"], 6):
            max_lines = 1
            for column, field in enumerate(fields):
                value = record[field["name"]]
                if value is None:
                    sheet.write_blank(row, column, None, f["text"])
                elif field["kind"] == "duration":
                    sheet.write_number(row, column, value / 86400, f["duration"])
                elif isinstance(value, datetime):
                    sheet.write_datetime(row, column, value, f["datetime"])
                elif isinstance(value, date):
                    sheet.write_datetime(row, column, value, f["date"])
                elif isinstance(value, time):
                    sheet.write_number(row, column, (value.hour * 3600 + value.minute * 60 + value.second) / 86400, f["time"])
                elif isinstance(value, (int, float)):
                    sheet.write_number(row, column, value, f["decimal"])
                else:
                    text = str(value)
                    sheet.write_string(row, column, text, f["source"])
                    width = 50 if field["kind"] == "textarea" else 21
                    max_lines = max(max_lines, sum(max(1, (len(line) + width - 1) // width) for line in text.splitlines()))
            sheet.set_row(row, min(240, max(23, max_lines * 15)))
        last_row = max(6, len(category["details"]) + 5)
        sheet.autofilter(5, 0, last_row, len(fields) - 1)
        sheet.freeze_panes(6, 0)
        sheet.repeat_rows(5)
        sheet.set_print_scale(85)
        sheet.print_area(0, 0, last_row, len(fields) - 1)
        prefix = "'" + category["sheet"].replace("'", "''") + "'!"
        date_letter, state_letter = xl_col_to_name(len(fields) - 2), xl_col_to_name(len(fields) - 1)
        refs[category["entity"]] = dict(count=f"{prefix}$A$7:$A${last_row + 1}",
                                        date=f"{prefix}${date_letter}$7:${date_letter}${last_row + 1}",
                                        state=f"{prefix}${state_letter}$7:${state_letter}${last_row + 1}")

    setup(summary, report["title"], 10)
    summary.set_column(0, 0, 34)
    summary.set_column(1, 6, 16)
    summary.set_column(7, 10, 13)
    if len(categories) == 1:
        summary.merge_range("A4:K4", "Categoría: " + categories[0]["label"], f["section"])
    summary.merge_range("A5:K5", report["method_note"], f["subtitle"])
    summary.set_row(4, 32)
    summary.merge_range("A6:K6", report["rate_note"], f["subtitle"])
    summary.set_row(5, 43)
    headers = ["Categoría", "Registros", "Abiertos", "Cerrados", "Por revisar", "Sin campo de estado", "Cierre"]
    summary.write_row(7, 0, headers, f["header"])
    summary.set_row(7, 33)
    for row, category in enumerate(categories, 8):
        totals, ref = category["totals"], refs[category["entity"]]
        summary.write_string(row, 0, category["label"], f["text"])
        summary.set_row(row, 32)
        summary.write_formula(row, 1, f"=ROWS({ref['count']})" if totals["records"] else "=0", f["number"], totals["records"])
        for col, key in enumerate(STATE_LABELS, 2):
            summary.write_formula(row, col, f'=COUNTIF({ref["state"]},"{STATE_LABELS[key]}")', f["number"], totals[key])
        summary.write_formula(row, 6, f'=IF(SUM(C{row+1}:D{row+1})=0,"",D{row+1}/SUM(C{row+1}:D{row+1}))',
                              f["rate"], totals["rate"] if totals["rate"] is not None else "")
    total_row = 8 + len(categories)
    summary.write_string(total_row, 0, "Total de registros", f["total"])
    summary.set_row(total_row, 27)
    for col, key in enumerate(("records", *STATE_LABELS), 1):
        letter = xl_col_to_name(col)
        summary.write_formula(total_row, col, f"=SUM({letter}9:{letter}{total_row})", f["total"], report["totals"][key])
    summary.write_formula(total_row, 6,
                          f'=IF(SUM(C{total_row+1}:D{total_row+1})=0,"",D{total_row+1}/SUM(C{total_row+1}:D{total_row+1}))',
                          f["total_rate"], report["totals"]["rate"] if report["totals"]["rate"] is not None else "")
    for col in (2, 3, 4, 5):
        summary.conditional_format(8, col, total_row - 1, col, {"type": "data_bar", "bar_color": {2:"#CF973C",3:"#278269",4:"#8290A0",5:"#C5CDD7"}[col]})
    row = total_row + 2
    summary.merge_range(row, 0, row, 10, "Conclusiones del período", f["section"])
    for note in report["conclusions"]:
        row += 1
        summary.merge_range(row, 0, row, 10, note, f["text"])
        summary.set_row(row, 29)
    row += 2
    summary.merge_range(row, 0, row, 10, "Seguimiento de información", f["section"])
    for category in categories:
        if not category.get("tracking"):
            continue
        row += 1
        summary.merge_range(row, 0, row, 10,
                            f"{category['label']}: {category['tracking_label']}: {category['totals']['tracked']}; "
                            f"{category['missing_label']}: {category['totals']['untracked']}.", f["text"])
        summary.set_row(row, 30)
    row += 1
    summary.merge_range(row, 0, row, 10,
                        f"Registros sin fecha que coinciden con la búsqueda: {report['totals']['undated_available']}. "
                        f"Incluidos en esta descarga: {report['totals']['undated']}.", f["subtitle"])
    summary.set_row(row, 28)

    setup(daily, "Evolución diaria de los registros", max(7, len(categories) + 1))
    daily.set_column(0, 0, 19)
    daily.set_column(1, max(7, len(categories) + 1), 23)
    daily.merge_range(3, 0, 3, max(7, len(categories) + 1),
                      "Las fechas corresponden al origen del registro. Los registros sin fecha se muestran en una fila separada cuando se incluyen.", f["subtitle"])
    daily.set_row(3, 30)
    daily.write_row(5, 0, ["Fecha", *[c["label"] for c in categories], "Total"], f["header"])
    daily.set_row(5, 38)
    for i, day in enumerate(report["days"]):
        row_daily = i + 6
        daily.write_datetime(row_daily, 0, day, f["date"])
        for col, category in enumerate(categories, 1):
            ref = refs[category["entity"]]
            daily.write_formula(row_daily, col, f"=COUNTIF({ref['date']},$A{row_daily+1})", f["number"], category["daily"][i])
        daily.write_formula(row_daily, len(categories) + 1,
                            f"=SUM(B{row_daily+1}:{xl_col_to_name(len(categories))}{row_daily+1})", f["total"], report["daily"][i])
    last_dated = len(report["days"]) + 5
    last_daily = last_dated
    if report["filters"]["include_undated"]:
        last_daily += 1
        daily.write_string(last_daily, 0, "Sin fecha", f["text"])
        for col, category in enumerate(categories, 1):
            ref = refs[category["entity"]]
            # COUNTBLANK sobre un rango vacío contaría una fila ficticia: usar cero.
            formula = f"=COUNTBLANK({ref['date']})" if category["totals"]["records"] else "=0"
            daily.write_formula(last_daily, col, formula, f["number"], category["totals"]["undated"])
        daily.write_formula(last_daily, len(categories) + 1,
                            f"=SUM(B{last_daily+1}:{xl_col_to_name(len(categories))}{last_daily+1})", f["total"], report["totals"]["undated"])
    daily.write_string(last_daily + 1, 0, "Total de registros", f["total"])
    for col, category in enumerate(categories, 1):
        letter = xl_col_to_name(col)
        daily.write_formula(last_daily + 1, col, f"=SUM({letter}7:{letter}{last_daily+1})", f["total"], category["totals"]["records"])
    letter = xl_col_to_name(len(categories) + 1)
    daily.write_formula(last_daily + 1, len(categories) + 1, f"=SUM({letter}7:{letter}{last_daily+1})", f["total"], report["totals"]["records"])
    daily.autofilter(5, 0, last_daily, len(categories) + 1)
    daily.freeze_panes(6, 1)
    daily.repeat_rows(5)
    daily.fit_to_pages(1, 0)
    daily.print_area(0, 0, last_daily + 1, len(categories) + 1)

    chart_row = row + 3
    states = book.add_chart({"type": "bar", "subtype": "stacked"})
    for col, (key, label) in enumerate(STATE_LABELS.items(), 2):
        states.add_series({"name": label, "categories": [summary_name, 8, 0, total_row - 1, 0],
                           "values": [summary_name, 8, col, total_row - 1, col],
                           "categories_data": [c["label"] for c in categories],
                           "values_data": [c["totals"][key] for c in categories],
                           "fill": {"color": {"open":"#CF973C","closed":"#278269","unknown":"#8290A0","not_applicable":"#C5CDD7"}[key]},
                           "line": {"none": True}})
    states.set_title({"name": "Estado actual por categoría"})
    states.set_x_axis({"min": 0, "major_unit": count_axis_step(max((c['totals']['records'] for c in categories), default=0)), "num_format": "0"})
    states.set_y_axis({"reverse": True})
    states.set_legend({"position": "bottom"})
    states.set_size({"width": 960, "height": 350})
    summary.insert_chart(chart_row, 0, states)
    trend = book.add_chart({"type": "line"})
    for col, category in enumerate(categories, 1):
        trend.add_series({"name": category["label"], "categories": ["Evolución diaria", 6, 0, last_dated, 0],
                          "categories_data": [to_excel(day) for day in report["days"]],
                          "values_data": category["daily"],
                          "values": ["Evolución diaria", 6, col, last_dated, col], "line": {"color": category["color"], "width": 1.5}})
    trend.set_title({"name": "Evolución diaria"})
    trend.set_x_axis({"date_axis": True, "num_format": "dd/mm/yyyy"})
    trend.set_y_axis({"min": 0, "major_unit": count_axis_step(max(report["daily"], default=0)), "num_format": "0"})
    trend.set_legend({"position": "bottom"})
    trend.set_size({"width": 960, "height": 350})
    summary.insert_chart(chart_row + 19, 0, trend)
    notes_row = chart_row + 39
    summary.merge_range(notes_row, 0, notes_row, 10, "Fuentes y criterios", f["section"])
    notes = [report["source_note"],
             "Aprobada y Completada se consideran cerradas en Misceláneos y Solicitudes OT. No se infiere un cierre a partir de comentarios ni de fechas de liberación por área.",
             "Los detalles contienen todos los campos disponibles en el sistema, incluidas las columnas históricas de Solicitudes OT. Las columnas de fecha y estado agrupado se añaden para explicar los cálculos.",
             "Este archivo es una consulta de los datos al momento de la descarga. Para incorporar cambios del sistema, genera una nueva exportación."]
    notes += [c["label"] + ": " + c["date_label"] + "." for c in categories]
    for index, note in enumerate(notes, notes_row + 1):
        summary.merge_range(index, 0, index, 10, note, f["subtitle"])
        summary.set_row(index, 32)
    summary.freeze_panes(8, 1)
    summary.set_landscape()
    summary.fit_to_pages(1, 0)
    # Mantener los dos gráficos completos al imprimir el resumen.
    summary.set_h_pagebreaks([chart_row, notes_row])
    summary.repeat_rows(0, 2)
    summary.print_area(0, 0, notes_row + len(notes), 10)
