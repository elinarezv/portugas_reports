from collections import OrderedDict
from tempfile import NamedTemporaryFile
import base64
import datetime
import itertools
import functools
import dataclasses
import copy

import openpyxl as xlsx

from odoo import fields, models, api

#import cierre_diario_ajuste_de_inventario.models.generacion as cierre_diario_gen


def flatten(iter_of_iters):
    for sub in iter_of_iters:
        for res in sub:
            yield res


def groupallby(iterable, *, key, only_groups=False):
    groups = {}
    for el in iterable:
        my_key = key(el)
        group = groups.setdefault(my_key, [])
        group.append(el)
    if only_groups:
        return groups.values()
    else:
        return groups.items()


def copy_cell_style(worksheet, *, source_cell_index, target_cell_index):
    # Taken from
    # https://stackoverflow.com/questions/23332259/copy-cell-style-openpyxl. Biggest
    # thanks to Charlie Clark
    # (https://stackoverflow.com/users/2385133/charlie-clark) and pbarill
    # (https://stackoverflow.com/users/1247325/pbarill) for an useful extensive
    # answer, to Idrg (https://stackoverflow.com/users/296239/ldrg) for the
    # idea of only copying the style field and Avik Samaddar
    # (https://stackoverflow.com/users/4686330/avik-samaddar) for using `style`
    # rather than `_style`.
    trg_cell = worksheet[target_cell_index]
    src_cell = worksheet[source_cell_index]
    #trg_cell.style = copy.copy(src_cell.style) # Issues with the copy. Maybe I need a deep copy?
    for field in {"font", "border", "fill", "number_format", "protection", "alignment"}:
        setattr(trg_cell, field, copy.copy(getattr(src_cell, field)))


def update_row_dims(worksheet, *, source_row, target_row):
    worksheet.row_dimensions[target_row].height = \
        worksheet.row_dimensions[source_row].height


def clone_cell(worksheet, *, source_cell_index, target_cell_index):
    worksheet[target_cell_index] = worksheet[source_cell_index].value
    copy_cell_style(worksheet,
                    source_cell_index=source_cell_index,
                    target_cell_index=target_cell_index)


@dataclasses.dataclass
class ReportSalesPerProduct:
    product_id: object
    qty: float
    subtotal: float


def aggregate_lines_for_a_single_product(lines):
    return functools.reduce(
        lambda state, line: dataclasses.replace(
            state,
            product_id=line.product_id,
            qty=state.qty + line.product_uom_qty,
            subtotal=state.subtotal + line.price_subtotal,
        ),
        lines,
        ReportSalesPerProduct(qty=0, product_id=None, subtotal=0)
    )


def report_sales_by_product(workbook, env, instances):
    worksheet = workbook["Hoja4"]

    sorted_instances = list(instances)
    sorted_instances.sort(key=lambda order: order.date_order)

    if sorted_instances:
        worksheet["E10"] = sorted_instances[0].date_order
        worksheet["G10"] = sorted_instances[-1].date_order

    all_order_lines = flatten(order.order_line for order in sorted_instances)
    by_product_id = groupallby(all_order_lines, key=lambda line: line.product_id.id, only_groups=True)
    rows = (aggregate_lines_for_a_single_product(lines) for lines in by_product_id)

    for row_no, row in enumerate(rows):
        no = row_no + 12
        if no != 12:
            for col in "BCDEFH":
                copy_cell_style(worksheet,
                                source_cell_index=f"{col}12",
                                target_cell_index=f"{col}{no}")
            update_row_dims(worksheet, source_row=12, target_row=no)
        worksheet[f"B{no}"] = row.product_id.default_code
        worksheet[f"C{no}"] = row.product_id.name
        worksheet[f"D{no}"] = str(row.qty) + " " + str(row.product_id.uom_id.name)
        worksheet[f"E{no}"] = row.product_id.standard_price
        worksheet[f"F{no}"] = row.subtotal

    return workbook


def report_sales_by_client(workbook, env, instances):
    worksheet = workbook["Hoja5"]

    sorted_instances = list(instances)
    sorted_instances.sort(key=lambda order: order.date_order)

    if sorted_instances:
        worksheet["F14"] = sorted_instances[0].date_order
        worksheet["H14"] = sorted_instances[-1].date_order

    by_client = groupallby(sorted_instances, key=lambda order: order.partner_id.id, only_groups=True)
    sorted_dates = (sorted(orders, key=lambda order: order.date_order) for orders in by_client)
    order_lines_by_client = ((orders[0], flatten(order.order_line for order in orders)) for orders in sorted_dates)
    by_product_id = ((order, groupallby(lines, key=lambda line: line.product_id.id, only_groups=True))
                     for order, lines in order_lines_by_client)
    rows = ((order, (aggregate_lines_for_a_single_product(product_lines)
                     for product_lines in lines_groups))
            for order, lines_groups in by_product_id)

    no = 15  # 16 (start) - 1 to negate next + 1
    for row_no, (order, product_rows) in enumerate(rows):
        for row in product_rows:
            no += 1
            if no != 16:
                for col in "CDEFGHI":
                    copy_cell_style(worksheet,
                                    source_cell_index=f"{col}16",
                                    target_cell_index=f"{col}{no}")
                update_row_dims(worksheet, source_row=16, target_row=no)
            worksheet[f"C{no}"] = order.partner_id.name
            worksheet[f"D{no}"] = order.date_order
            worksheet[f"E{no}"] = row.product_id.default_code
            worksheet[f"F{no}"] = row.product_id.name
            worksheet[f"G{no}"] = str(row.qty) + " " + str(row.product_id.uom_id.name)
            worksheet[f"H{no}"] = row.product_id.standard_price
            worksheet[f"I{no}"] = row.subtotal

    return workbook


def report_products_by_year(workbook, env, instances):
    worksheet = workbook["Hoja1"]

    sorted_instances = list(instances)
    sorted_instances.sort(key=lambda order: order.date_order)
    by_location = list(groupallby(sorted_instances, key=lambda order: order.fsm_location_id.id, only_groups=True))
    location_start_and_end_dates = {group[0].fsm_location_id.id: (group[0].date_order, group[-1].date_order) for group in by_location}

    by_location_by_month = ((groupallby(group, key=lambda order: order.date_order.month, only_groups=True))
                            for group in by_location)
    by_location_by_month_lines = \
        (((group[0], flatten(order.order_line for order in group)) for group in by_month)
         for by_month in by_location_by_month)
    by_location_by_month_by_product_id = \
        (((order, groupallby(lines, key=lambda line: line.product_id.id, only_groups=True))
          for order, lines in by_month)
         for by_month in by_location_by_month_lines)
    by_location_by_month_aggregated = \
        (((order, (aggregate_lines_for_a_single_product(lines) for lines in by_product_id))
          for order, by_product_id in by_month)
         for by_month in by_location_by_month_by_product_id)

    no = 37
    for month_0, by_month in enumerate(by_location_by_month_aggregated):
        for order, rows in by_month:
            rows = list(rows)
            if not rows:
                continue
            no += 1
            update_row_dims(worksheet, source_row=12, target_row=no)
            for col in "EF":
                clone_cell(worksheet,
                           source_cell_index="G12",
                           target_cell_index=f"{col}{no}")
            for col in "GHIJK":
                clone_cell(worksheet,
                           source_cell_index=f"{col}12",
                           target_cell_index=f"{col}{no}")
            no += 1
            update_row_dims(worksheet, source_row=13, target_row=no)
            start, end = location_start_and_end_dates[order.fsm_location_id.id]
            for col in "FGHIJK":
                copy_cell_style(worksheet,
                                source_cell_index=f"{col}13",
                                target_cell_index=f"{col}{no}")
            clone_cell(worksheet,
                       source_cell_index="E13",
                       target_cell_index=f"E{no}")
            worksheet[f"F{no}"] = order.fsm_location_id.name
            worksheet[f"I{no}"] = start
            worksheet[f"K{no}"] = end
            no += 1
            update_row_dims(worksheet, source_row=14, target_row=no)
            for col in "EFGHIJK":
                clone_cell(worksheet,
                           source_cell_index=f"{col}14",
                           target_cell_index=f"{col}{no}")
            wrote_month_col = False
            for row in rows:
                no += 1
                update_row_dims(worksheet, source_row=15, target_row=no)
                if not wrote_month_col:
                    month_row = (order.date_order.month - 1) * 2 + 15
                    clone_cell(worksheet,
                               source_cell_index=f"E{month_row}",
                               target_cell_index=f"E{no}")
                    wrote_month_col = True
                else:
                    clone_cell(worksheet,
                               source_cell_index="E16",
                               target_cell_index=f"E{no}")
                for col in "FGHIJK":
                    copy_cell_style(worksheet,
                                    source_cell_index=f"{col}15",
                                    target_cell_index=f"{col}{no}")
                worksheet[f"F{no}"] = order.date_order
                worksheet[f"G{no}"] = row.product_id.default_code
                worksheet[f"H{no}"] = row.product_id.name
                worksheet[f"I{no}"] = str(row.qty) + " " + str(row.product_id.uom_id.name)
                worksheet[f"J{no}"] = row.product_id.standard_price
                worksheet[f"K{no}"] = row.subtotal

    # Delete rows from 13 to 37 + the one extra header line
    worksheet.delete_rows(13, 37 - 13 + 2)

    return workbook


def report_sales_by_day(workbook, env, instances):
    worksheet = workbook["Hoja2"]

    sorted_instances = list(instances)
    sorted_instances.sort(key=lambda order: order.date_order)
    by_location = list(groupallby(sorted_instances, key=lambda order: order.fsm_location_id.id, only_groups=True))
    location_start_and_end_dates = {group[0].fsm_location_id.id: (group[0].date_order, group[-1].date_order) for group in by_location}

    by_location_this_month = ([order for order in orders
                               if order.date_order.date() == datetime.date.today()]
                              for orders in by_location)

    by_location_this_month_by_lines = ((orders[0], flatten(order.order_line for order in orders))
                                       for orders in by_location_this_month
                                       if orders)
    by_location_this_month_by_product_id = \
        ((order, groupallby(lines, key=lambda line: line.product_id.id, only_groups=True))
         for order, lines in by_location_this_month_by_lines)
    by_location_this_month_by_aggregated = \
        ((order, (aggregate_lines_for_a_single_product(lines) for lines in groups))
         for order, groups in by_location_this_month_by_product_id)

    no = 73
    for order, rows in by_location_this_month_by_aggregated:
        rows = list(rows)
        if not rows:
            continue
        no += 1
        update_row_dims(worksheet, source_row=10, target_row=no)
        for col in "CDE":
            clone_cell(worksheet,
                       source_cell_index="G10",
                       target_cell_index=f"{col}{no}")
        for col in "FGH":
            clone_cell(worksheet,
                       source_cell_index=f"{col}10",
                       target_cell_index=f"{col}{no}")
        no += 1
        update_row_dims(worksheet, source_row=11, target_row=no)
        clone_cell(worksheet,
                   source_cell_index="C11",
                   target_cell_index=f"C{no}")
        start, end = location_start_and_end_dates[order.fsm_location_id.id]
        worksheet[f"D{no}"] = order.fsm_location_id.name
        worksheet[f"F{no}"] = start
        worksheet[f"H{no}"] = end
        no += 1
        update_row_dims(worksheet, source_row=12, target_row=no)
        for col in "CDEFGH":
            clone_cell(worksheet,
                       source_cell_index=f"{col}12",
                       target_cell_index=f"{col}{no}")
        wrote_day_col = False
        for row in rows:
            no += 1
            update_row_dims(worksheet, source_row=13, target_row=no)
            for col in "CDEFGH":
                copy_cell_style(worksheet,
                                source_cell_index=f"{col}14",
                                target_cell_index=f"{col}{no}")
            if not wrote_day_col:
                wrote_day_col = True
                day_col = (order.date_order.day - 1) * 2 + 13
                clone_cell(worksheet,
                           source_cell_index=f"C{day_col}",
                           target_cell_index=f"C{no}")
            worksheet[f"E{no}"] = order.partner_id.name
            worksheet[f"F{no}"] = str(row.qty) + " " + str(row.product_id.uom_id.name)
            worksheet[f"H{no}"] = row.subtotal

    # Delete rows from 11 to 73 + the one extra header line
    worksheet.delete_rows(11, 73 - 11 + 2)

    return workbook


class ExcelReport(models.Model):
    _name = "portugas_reports.excel_report"

    name = fields.Char(stored=False, compute="_compute_name")

    REPORT_TYPES = [
        ("report_products_by_year", "Report Sales By Year"),
        ("report_sales_by_product", "Report Sales by Product"),
        ("report_sales_by_day", "Report Sales by Day"),
        ("report_sales_by_client", "Report Sales by Client"),
        ("report_daily_close", "Report Daily Closing"),
    ]

    report_type = fields.Selection(
        string="Report type",
        selection=REPORT_TYPES,
        required=True,
    )

    xlsx_template = fields.Many2one(
        "ir.attachment",
        ondelete="restrict",
        required=False,
        string="Template to use when creating the reports.",
    )

    @api.depends("report_type")
    def _compute_name(self):
        mapping = dict(self.REPORT_TYPES)
        for record in self:
            record.name = mapping[record.report_type]

    def _template_contents(self):
        return self.xlsx_template.raw

    def _fill_template_workbook(self, wb, extra):
        if self.report_type == "report_daily_close":
            # Defined on another odoo module
            return self._report_daily_close(wb, extra)
        else:
            sales = self.env["sale.order"].search([])
            if self.report_type == "report_products_by_year":
                return report_products_by_year(wb, self.env, sales)
            elif self.report_type == "report_sales_by_product":
                return report_sales_by_product(wb, self.env, sales)
            elif self.report_type == "report_sales_by_client":
                return report_sales_by_client(wb, self.env, sales)
            elif self.report_type == "report_sales_by_day":
                return report_sales_by_day(wb, self.env, sales)
            else:
                raise Exception("Invalid report type: " + self.report_type)

    def _generate_attachment_name(self, wb):
        return datetime.datetime.now().isoformat()

    def _create_attachment_with_report(self, extra):
        with NamedTemporaryFile(suffix=".xlsx") as tmp:
            tmp.write(self._template_contents())
            wb = xlsx.load_workbook(tmp.name)
        self._fill_template_workbook(wb, extra)
        with NamedTemporaryFile(suffix=".xlsx") as tmp:
            wb.save(tmp.name)
            tmp.seek(0)
            contents = tmp.read()
        b64_data = base64.encodebytes(contents)
        return self.env["ir.attachment"].create({
            "name": "Report {}.xlsx".format(self._generate_attachment_name(wb)),
            "datas": b64_data,
        })

    def action_create_report(self, extra=None):
        assert self.xlsx_template
        if extra is None:
            extra = {}
        attachment = self._create_attachment_with_report(extra)
        return {
            "type": "ir.actions.act_window",
            "res_model": "ir.attachment",
            "views": [[False, "form"]],
            "res_id": attachment.id,
        }
