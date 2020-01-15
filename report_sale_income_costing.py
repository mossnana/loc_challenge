# -*- coding: utf-8 -*-
import xlwt
from io import StringIO, BytesIO
import base64
import time

from odoo import models, api, fields, _
from odoo.exceptions import Warning
from dateutil import parser
from dateutil.relativedelta import relativedelta
from datetime import datetime, timedelta
from odoo.tools import DEFAULT_SERVER_DATE_FORMAT as DF
from pdb import set_trace as debug
from itertools import groupby


def font_style(position='left', bold=0, fontos=0, font_height=200, border=0, color=False):
    font = xlwt.Font()
    font.name = 'Verdana'
    font.bold = bold
    font.height = font_height
    center = xlwt.Alignment()
    center.horz = xlwt.Alignment.HORZ_CENTER
    center.vert = xlwt.Alignment.VERT_CENTER
    center.wrap = xlwt.Alignment.VERT_JUSTIFIED

    left = xlwt.Alignment()
    left.horz = xlwt.Alignment.HORZ_LEFT
    left.vert = xlwt.Alignment.VERT_CENTER
    left.wrap = xlwt.Alignment.VERT_JUSTIFIED

    right = xlwt.Alignment()
    right.horz = xlwt.Alignment.HORZ_RIGHT
    right.vert = xlwt.Alignment.VERT_CENTER
    right.wrap = xlwt.Alignment.VERT_JUSTIFIED

    borders = xlwt.Borders()
    borders.right = 1
    borders.left = 1
    borders.top = 1
    borders.bottom = 1

    orient = xlwt.Alignment()
    orient.orie = xlwt.Alignment.ORIENTATION_90_CC

    style = xlwt.XFStyle()

    if border == 1:
        style.borders = borders

    if fontos == 'red':
        font.colour_index = 2
        style.font = font
    if fontos == 'purple_ega':
        font.colour_index = 0x14
        style.font = font
    else:
        style.font = font

    if position == 'center':
        style.alignment = center
    elif position == 'right':
        style.alignment = right
    else:
        style.alignment = left
    if color == 'grey':
        badBG = xlwt.Pattern()
        badBG.pattern = badBG.SOLID_PATTERN
        badBG.pattern_fore_colour = 22
        style.pattern = badBG
    if color == 'red':
        badBG = xlwt.Pattern()
        badBG.pattern = badBG.SOLID_PATTERN
        badBG.pattern_fore_colour = 5
        style.pattern = badBG

    if color == 'yellow':
        badBG = xlwt.Pattern()
        badBG.pattern = badBG.SOLID_PATTERN
        badBG.pattern_fore_colour = 0x0D
        style.pattern = badBG

    if color == 'purple':
        badBG = xlwt.Pattern()
        badBG.pattern = badBG.SOLID_PATTERN
        badBG.pattern_fore_colour = 0x14
        style.pattern = badBG

    return style

class SaleIncomeCostingAnalysisExcel(models.TransientModel):
    _name = 'sale.income.costing.analysis.excel'

    start_date = fields.Date(
        string='From Date',
        required=True,
        default=lambda *a: (parser.parse(datetime.now().strftime(DF)))
    )
    end_date = fields.Date(
        string='To Date',
        required=True,
        default=lambda *a: (parser.parse(datetime.now().strftime(DF)))
    )
    filter_product_ids = fields.Many2many('product.product', string='Products')
    filter_product_categ_ids = fields.Many2many(
        'product.category', string='Categories')

    @api.multi
    def report_history(self):
        return {
            'name': 'History',
            'type': 'ir.actions.act_window',
            'res_model': 'sale.income.costing.analysis.excel',
            'view_id': self.env.ref('ac_naraipak.view_sale_income_costing_analysis_excel_history').id,
            'view_mode': 'tree',
        }

    @api.multi
    def print_xls_report(self):
        self.ensure_one()
        self.fetch_data()
        # New Workbook
        workbook = xlwt.Workbook()
        # Set Styles
        header_style = font_style(
            position='center', bold=1, border=1, fontos='black', font_height=400, color='grey')
        column_style_1 = font_style(
            position='center', bold=1, border=1, fontos='black', font_height=180, color='grey')
        detail_style_1 = font_style(
            position='center', fontos='purple_ega', bold=1, font_height=180)
        # Header
        header_text = 'Sales Income and Costing Analysis'
        sheet = workbook.add_sheet(header_text)
        sheet.set_panes_frozen(True)
        sheet.set_horz_split_pos(8)
        sheet.row(0).height = 256 * 3
        sheet.write_merge(0, 0, 0, 11, header_text, header_style)
        # Subheader
        sheet_start_header = 3
        sheet_start_value = 4
        column_row_start = 7
        sheet.write_merge(sheet_start_header, sheet_start_header,
                          0, 1, 'Date', column_style_1)
        sheet.write_merge(sheet_start_value, sheet_start_value, 0, 1, '{} To {}'.format(
            self.start_date, self.end_date), column_style_1)
        # Columns
        columns = [
            'Distr. Channel',
            'Product Code',
            'Product Name',
            'Quantity',
            'Revenue',
            'Plus',
            'Deduction',
            'Net Revenue',
            'COGs',
            'Margin 1',
            'Total Prod. Labor',
            'Variant Adj. Labor',
            'Variant Adj. Material Cost',
            'Variant Adj. Workcenter Cost',
            'Production Amount',
            'Production Quantity',
            'Variant Adj. Subcontract',
            'Subcontract Amount',
            'Subcontract Quantity',
            'Variant Adjust Solvent',
            'PP. Print Amount',
            'PP. Print Quantity',
            'Total Adj. Variant',
            'Total Variant',
            'Margin 2',
            'Sale Costs',
            'Administration Costs',
            'Total Sales & Admin Costs',
            'Margin 3',
        ]
        # Rows
        rows = [
            'channel',
            'product_code',
            'product_name',
            'quantity',
            'revenue',
            'plus',
            'deduction',
            'net_revenue',
            'cogs',
            'margin_1',
            'total_product_variant',
            'variant_adjust_labor',
            'variant_adjust_material_cost',
            'variant_adjust_workcenter_cost',
            'production_amount',
            'production_quantity',
            'variant_adjust_subcontract',
            'subcontract_amount',
            'subcontract_quantity',
            'variant_adjust_solvent',
            'pp_print_amount',
            'pp_print_quantity',
            'total_adjust_variant',
            'total_variant',
            'margin_2',
            'sale_costs',
            'administrator_costs',
            'total_sale_administrator_costs',
            'margin_3',
        ]
        for index, column in enumerate(columns):
            sheet.write(column_row_start, index, column, column_style_1)
            sheet.col(index).width = 256 * 30
        row = 8
        data = self.fetch_data()
        for key, value in data.items():
            for lines in value:
                for line_index, line_value in enumerate(rows):
                    sheet.row(row).height_mismatch = True
                    sheet.row(row).height = int(256 * 1.5)
                    sheet.write(row, line_index,
                                lines[line_value], detail_style_1)
                row += 1
        stream = BytesIO()
        workbook.save(stream)
        ExcelFile = self.env['sale.income.costing.analysis.excel.file']
        res_id = ExcelFile.create({
            'file': base64.encodestring(stream.getvalue()),
            'name': 'Production Variance Report.xls'
        })
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/binary/download_document?model=sale.income.costing.analysis.excel.file&field=file&id={}&filename=Production Variance Report.xls'.format(res_id.id),
            'target': 'new',
        }

    @api.multi
    def fetch_data(self):
        Product = self.env['product.product']
        SaleOrder = self.env['sale.order']
        SaleOrderLine = self.env['sale.order.line']
        CostLine = self.env['account.cost.line']
        # Excel Data Template
        vals = {
            'production': [],
            'export': [],
            'domestic': [],
            'interco': [],
            'other': [],
        }
        # Product Line Data Template
        line_vals = {
            'channel': '',
            'product_code': '',
            'product_name': '',
            'quantity': 0,
            'revenue': 0.00,
            'plus': 0.00,
            'deduction': 0.00,
            'net_revenue': 0.00,
            'cogs': 0.00,
            'margin_1': 0.00,
            'total_product_variant': 0.00,
            'variant_adjust_labor': 0.00,
            'variant_adjust_material_cost': 0.00,
            'variant_adjust_workcenter_cost': 0.00,
            'production_amount': 0.00,
            'production_quantity': 0,
            'variant_adjust_subcontract': 0.00,
            'subcontract_amount': 0.00,
            'subcontract_quantity': 0.00,
            'variant_adjust_solvent': 0.00,
            'pp_print_amount': 0.00,
            'pp_print_quantity': 0,
            'total_adjust_variant': 0.00,
            'total_variant': 0.00,
            'margin_2': 0.00,
            'sale_costs': 0.00,
            'administrator_costs': 0.00,
            'total_sale_administrator_costs': 0.00,
            'margin_3': 0.00,
        }
        # Check Category
        selected_cate = self.filter_product_categ_ids.mapped('id')
        all_child_cate = Product.read_group(
            [('categ_id', 'child_of', selected_cate)], ['categ_id'], ['categ_id'])
        all_child_cate_list = []
        for cate in all_child_cate:
            if cate.get('categ_id'):
                all_child_cate_list.append(cate.get('categ_id')[0])
        product_id_from_cate = Product.search(
            [('categ_id', 'in', all_child_cate_list)])
        # Check Product
        selected_product = self.filter_product_ids
        # Combine Cate Product Line and Product Line
        total_product = selected_product | product_id_from_cate
        # Production
        production_lines = self.fetch_data_production(total_product)
        # Cost Lines
        cost_lines = CostLine.search([
            ['date', '>=', self.start_date],
            ['date', '<=', self.end_date],
            ['picking_type_id.code', '=', 'outgoing'],
        ])
        cost_lines_origin = cost_lines.mapped('picking_origin')
        if total_product:
            order_lines = SaleOrderLine.search([
                ['order_id.name', 'in', cost_lines_origin],
                ['product_id.id', 'in', total_product.mapped('id')],
                ['order_id.state', 'in', ['sale','done']]
            ])
        else:
            order_lines = SaleOrderLine.search([
                ['order_id.name', 'in', cost_lines_origin],
                ['order_id.state', 'in', ['sale','done']],
            ])
        # Export
        export_order_lines = order_lines.filtered(lambda line: line.order_id.sale_type == 'out')
        print(export_order_lines)
        export_lines = self.fetch_data_export(export_order_lines)
        # Domestic
        domestic_order_lines = order_lines.filtered(lambda line: line.order_id.sale_type == 'in')
        domestic_lines = self.fetch_data_domestic(domestic_order_lines)
        # Excel Data
        vals = {
            'production': production_lines,
            'export': export_lines,
            'domestic': domestic_lines,
            'interco': [],
            'other': [],
        }
        return vals

    def fetch_data_production(self, products):
        if products:
            productions = self.env['mrp.production'].search([
                ('date_planned_start', '>=', self.start_date + " 00:00:00"),
                ('date_planned_start', '<=', self.end_date + " 23:59:59"),
                ('state', 'in', ['progress', 'done']),
                ('product_id', 'in', products.mapped('id')),
            ])
        else:
            productions = self.env['mrp.production'].search([
                ('date_planned_start', '>=', self.start_date + " 00:00:00"),
                ('date_planned_start', '<=', self.end_date + " 23:59:59"),
                ('state', 'in', ['progress', 'done']),
            ])
        new_line = {}
        lines_list = []
        grouped_list = []
        for production in productions:
            total_produced = 0
            for finished in production.finished_move_line_ids:
                if finished.date >= self.start_date + " 00:00:00" and finished.date <= self.end_date + " 23:59:59" and production.product_id.id == finished.product_id.id:
                    total_produced += finished.qty_done
            variance_material_cost = round(
                ((production.material_total * total_produced / production.product_qty) - production.total_actual_material_cost), 2)
            variance_labour_cost = round(
                ((production.labor_total * total_produced / production.product_qty) - production.total_actual_labour_cost), 2)
            variance_workcenter_cost = round(
                ((production.workcenter_total * total_produced / production.product_qty) - production.total_actual_workcenter_cost), 2)
            total_production_variance = round(
                variance_material_cost + variance_labour_cost + variance_workcenter_cost, 2)
            product_code = production.product_id.code
            product_name = production.product_id.name
            production_amount = round(production.final_total_actual_cost, 2)
            # Variant Adjust Solvent
            MaterialCostLine = production.direct_material_ids
            variant_adjust_solvent = sum(
                [p.total_cost * (p.product_id.categ_id.solvent_per_unit / 100) for p in MaterialCostLine])
            total_product_variant = variance_material_cost + \
                variance_labour_cost + variance_workcenter_cost
            # Total Adjust Var.
            total_adjust_variant = variance_material_cost + variance_labour_cost + \
                variance_workcenter_cost + variant_adjust_solvent
            total_product_adjust_variant = (
                variance_material_cost + variance_labour_cost + variance_workcenter_cost) - total_adjust_variant
            margin_2 = total_adjust_variant
            new_line = {
                'channel': 'Production',
                'product_code': product_code,
                'product_name': product_name,
                'quantity': 0,
                'revenue': 0.00,
                'plus': 0.00,
                'deduction': 0.00,
                'net_revenue': 0.00,
                'cogs': 0.00,
                'margin_1': 0.00,
                'total_product_variant': total_product_variant,
                'variant_adjust_labor': variance_labour_cost,
                'variant_adjust_material_cost': variance_material_cost,
                'variant_adjust_workcenter_cost': variance_workcenter_cost,
                'production_amount': production_amount,
                'production_quantity': total_produced,
                'variant_adjust_subcontract': 0.00,
                'subcontract_amount': 0.00,
                'subcontract_quantity': 0.00,
                'variant_adjust_solvent': variant_adjust_solvent,
                'pp_print_amount': 0.00,
                'pp_print_quantity': 0,
                'total_adjust_variant': total_adjust_variant,
                'total_variant': total_product_adjust_variant,
                'margin_2': margin_2,
                'sale_costs': 0.00,
                'administrator_costs': 0.00,
                'total_sale_administrator_costs': 0.00,
                'margin_3': 0.00,
            }
            lines_list.append(new_line)
        lines_list = sorted(lines_list, key=lambda x: x['product_code'])
        for model, group in groupby(lines_list, lambda x: x['product_code']):
            product_group = list(group)
            new_line = {
                'channel': 'Production',
                'product_code': product_group[0]['product_code'],
                'product_name': product_group[0]['product_name'],
                'quantity': 0,
                'revenue': 0.00,
                'plus': 0.00,
                'deduction': 0.00,
                'net_revenue': 0.00,
                'cogs': 0.00,
                'margin_1': 0.00,
                'total_product_variant': 0.00,
                'variant_adjust_labor': 0.00,
                'variant_adjust_material_cost': 0.00,
                'variant_adjust_workcenter_cost': 0.00,
                'production_amount': 0.00,
                'production_quantity': 0,
                'variant_adjust_subcontract': 0.00,
                'subcontract_amount': 0.00,
                'subcontract_quantity': 0.00,
                'variant_adjust_solvent': 0.00,
                'pp_print_amount': 0.00,
                'pp_print_quantity': 0,
                'total_adjust_variant': 0.00,
                'total_variant': 0.00,
                'margin_2': 0.00,
                'sale_costs': 0.00,
                'administrator_costs': 0.00,
                'total_sale_administrator_costs': 0.00,
                'margin_3': 0.00,
            }
            for product_line in product_group:
                new_line['total_product_variant'] += product_line['total_product_variant']
                new_line['variant_adjust_labor'] += product_line['variant_adjust_labor']
                new_line['variant_adjust_material_cost'] += product_line['variant_adjust_material_cost']
                new_line['variant_adjust_workcenter_cost'] += product_line['variant_adjust_workcenter_cost']
                new_line['production_amount'] += product_line['production_amount']
                new_line['production_quantity'] += product_line['production_quantity']
                new_line['variant_adjust_solvent'] += product_line['variant_adjust_solvent']
                new_line['total_adjust_variant'] += product_line['total_adjust_variant']
                new_line['total_variant'] += product_line['total_variant']
                new_line['margin_2'] += product_line['margin_2']
                new_line['total_product_variant'] += product_line['total_product_variant']
            grouped_list.append(new_line)
        return grouped_list

    def fetch_data_export(self, order_lines):
        CostLine = self.env['account.cost.line']
        AccountInvoice = self.env['account.invoice']
        lines_list = []
        grouped_list = []
        for order_line in order_lines:
            cost_lines = CostLine.search([['picking_origin', '=', order_line.order_id.name]])
            invoice_orders = order_line.mapped('order_id').mapped('invoice_ids')
            out_invoice_orders = invoice_orders.filtered(lambda inv: inv.type == 'out_invoice' and inv.state in ['open','paid','cancel'])
            plus_invoice_orders = AccountInvoice.search([['origin', 'in', out_invoice_orders.mapped('number')],['type','=','out_refund'],['state','in',['open','paid','cancel']]])
            deduction_invoice_orders = AccountInvoice.search([['origin', 'in', out_invoice_orders.mapped('number')],['type','=','out_charge'],['state','in',['open','paid','cancel']]])
            out_invoice_lines = out_invoice_orders.mapped('invoice_line_ids').filtered(lambda line: line.product_id.id == order_line.product_id.id)
            plus_invoice_lines = plus_invoice_orders.mapped('invoice_line_ids').filtered(lambda line: line.product_id.id == order_line.product_id.id)
            deduction_invoice_lines = deduction_invoice_orders.mapped('invoice_line_ids').filtered(lambda line: line.product_id.id == order_line.product_id.id)
            line_vals = {
                'channel': 'Export',
                'product_code': '',
                'product_name': '',
                'quantity': 0,
                'revenue': 0.00,
                'plus': 0.00,
                'deduction': 0.00,
                'net_revenue': 0.00,
                'cogs': 0.00,
                'margin_1': 0.00,
                'total_product_variant': 0.00,
                'variant_adjust_labor': 0.00,
                'variant_adjust_material_cost': 0.00,
                'variant_adjust_workcenter_cost': 0.00,
                'production_amount': 0.00,
                'production_quantity': 0,
                'variant_adjust_subcontract': 0.00,
                'subcontract_amount': 0.00,
                'subcontract_quantity': 0.00,
                'variant_adjust_solvent': 0.00,
                'pp_print_amount': 0.00,
                'pp_print_quantity': 0,
                'total_adjust_variant': 0.00,
                'total_variant': 0.00,
                'margin_2': 0.00,
                'sale_costs': 0.00,
                'administrator_costs': 0.00,
                'total_sale_administrator_costs': 0.00,
                'margin_3': 0.00,
            }
            line_vals['product_code'] = order_line.product_id.code
            line_vals['product_name'] = order_line.product_id.name
            line_vals['quantity'] = round(sum(out_invoice_lines.mapped('quantity')),2)
            for inv in out_invoice_lines:
                line_vals['revenue'] += round((inv.price_subtotal / inv.invoice_id.currency_rate),2) if inv.invoice_id.currency_rate != 0 else 0
            for plus_inv in plus_invoice_lines:
                line_vals['plus'] += round((plus_inv.price_subtotal / plus_inv.invoice_id.currency_rate),2) if plus_inv.invoice_id.currency_rate != 0 else 0
            for deduction_inv in deduction_invoice_lines:
                line_vals['deduction'] += round((deduction_inv.price_subtotal / deduction_inv.invoice_id.currency_rate),2) if deduction_inv.invoice_id.currency_rate != 0 else 0
            line_vals['net_revenue'] = line_vals['revenue'] + line_vals['plus'] - line_vals['deduction']
            for cost_line in cost_lines:
                line_vals['cogs'] += round(cost_line.amount,2)
            line_vals['margin_1'] = line_vals['net_revenue'] - line_vals['cogs']
            lines_list.append(line_vals)
        lines_list = sorted(lines_list, key=lambda x: x['product_code'])
        for model, group in groupby(lines_list, lambda x: x['product_code']):
            product_group = list(group)
            new_line = {
                'channel': 'Export',
                'product_code': product_group[0]['product_code'],
                'product_name': product_group[0]['product_name'],
                'quantity': 0,
                'revenue': 0.00,
                'plus': 0.00,
                'deduction': 0.00,
                'net_revenue': 0.00,
                'cogs': 0.00,
                'margin_1': 0.00,
                'total_product_variant': 0.00,
                'variant_adjust_labor': 0.00,
                'variant_adjust_material_cost': 0.00,
                'variant_adjust_workcenter_cost': 0.00,
                'production_amount': 0.00,
                'production_quantity': 0,
                'variant_adjust_subcontract': 0.00,
                'subcontract_amount': 0.00,
                'subcontract_quantity': 0.00,
                'variant_adjust_solvent': 0.00,
                'pp_print_amount': 0.00,
                'pp_print_quantity': 0,
                'total_adjust_variant': 0.00,
                'total_variant': 0.00,
                'margin_2': 0.00,
                'sale_costs': 0.00,
                'administrator_costs': 0.00,
                'total_sale_administrator_costs': 0.00,
                'margin_3': 0.00,
            }
            for product_line in product_group:
                new_line['quantity'] += product_line['quantity']
                new_line['revenue'] += product_line['revenue']
                new_line['plus'] += product_line['plus']
                new_line['deduction'] += product_line['deduction']
                new_line['net_revenue'] += product_line['net_revenue']
                new_line['cogs'] += product_line['cogs']
            new_line['margin_1'] = new_line['net_revenue'] - product_line['cogs']
            grouped_list.append(new_line)
        return grouped_list

    def fetch_data_domestic(self, order_lines):
        CostLine = self.env['account.cost.line']
        AccountInvoice = self.env['account.invoice']
        lines_list = []
        grouped_list = []
        for order_line in order_lines:
            cost_lines = CostLine.search([['picking_origin', '=', order_line.order_id.name]])
            invoice_orders = order_line.mapped('order_id').mapped('invoice_ids')
            out_invoice_orders = invoice_orders.filtered(lambda inv: inv.type == 'out_invoice' and inv.state in ['open','paid','cancel'])
            plus_invoice_orders = AccountInvoice.search([['origin', 'in', out_invoice_orders.mapped('number')],['type','=','out_refund'],['state','in',['open','paid','cancel']]])
            deduction_invoice_orders = AccountInvoice.search([['origin', 'in', out_invoice_orders.mapped('number')],['type','=','out_charge'],['state','in',['open','paid','cancel']]])
            out_invoice_lines = out_invoice_orders.mapped('invoice_line_ids').filtered(lambda line: line.product_id.id == order_line.product_id.id)
            plus_invoice_lines = plus_invoice_orders.mapped('invoice_line_ids').filtered(lambda line: line.product_id.id == order_line.product_id.id)
            deduction_invoice_lines = deduction_invoice_orders.mapped('invoice_line_ids').filtered(lambda line: line.product_id.id == order_line.product_id.id)
            line_vals = {
                'channel': 'Domestic',
                'product_code': '',
                'product_name': '',
                'quantity': 0,
                'revenue': 0.00,
                'plus': 0.00,
                'deduction': 0.00,
                'net_revenue': 0.00,
                'cogs': 0.00,
                'margin_1': 0.00,
                'total_product_variant': 0.00,
                'variant_adjust_labor': 0.00,
                'variant_adjust_material_cost': 0.00,
                'variant_adjust_workcenter_cost': 0.00,
                'production_amount': 0.00,
                'production_quantity': 0,
                'variant_adjust_subcontract': 0.00,
                'subcontract_amount': 0.00,
                'subcontract_quantity': 0.00,
                'variant_adjust_solvent': 0.00,
                'pp_print_amount': 0.00,
                'pp_print_quantity': 0,
                'total_adjust_variant': 0.00,
                'total_variant': 0.00,
                'margin_2': 0.00,
                'sale_costs': 0.00,
                'administrator_costs': 0.00,
                'total_sale_administrator_costs': 0.00,
                'margin_3': 0.00,
            }
            line_vals['product_code'] = order_line.product_id.code
            line_vals['product_name'] = order_line.product_id.name
            line_vals['quantity'] = round(sum(out_invoice_lines.mapped('quantity')),2)
            for inv in out_invoice_lines:
                line_vals['revenue'] += round(inv.price_subtotal,2)
            for plus_inv in plus_invoice_lines:
                line_vals['plus'] += round(plus_inv.price_subtotal,2)
            for deduction_inv in deduction_invoice_lines:
                line_vals['deduction'] += round(deduction_inv.price_subtotal,2)
            line_vals['net_revenue'] = line_vals['revenue'] + line_vals['plus'] - line_vals['deduction']
            for cost_line in cost_lines:
                line_vals['cogs'] += round(cost_line.amount,2)
            line_vals['margin_1'] = line_vals['net_revenue'] - line_vals['cogs']
            lines_list.append(line_vals)
        lines_list = sorted(lines_list, key=lambda x: x['product_code'])
        for model, group in groupby(lines_list, lambda x: x['product_code']):
            product_group = list(group)
            new_line = {
                'channel': 'Domestic',
                'product_code': product_group[0]['product_code'],
                'product_name': product_group[0]['product_name'],
                'quantity': 0,
                'revenue': 0.00,
                'plus': 0.00,
                'deduction': 0.00,
                'net_revenue': 0.00,
                'cogs': 0.00,
                'margin_1': 0.00,
                'total_product_variant': 0.00,
                'variant_adjust_labor': 0.00,
                'variant_adjust_material_cost': 0.00,
                'variant_adjust_workcenter_cost': 0.00,
                'production_amount': 0.00,
                'production_quantity': 0,
                'variant_adjust_subcontract': 0.00,
                'subcontract_amount': 0.00,
                'subcontract_quantity': 0.00,
                'variant_adjust_solvent': 0.00,
                'pp_print_amount': 0.00,
                'pp_print_quantity': 0,
                'total_adjust_variant': 0.00,
                'total_variant': 0.00,
                'margin_2': 0.00,
                'sale_costs': 0.00,
                'administrator_costs': 0.00,
                'total_sale_administrator_costs': 0.00,
                'margin_3': 0.00,
            }
            for product_line in product_group:
                new_line['quantity'] += product_line['quantity']
                new_line['revenue'] += product_line['revenue']
                new_line['plus'] += product_line['plus']
                new_line['deduction'] += product_line['deduction']
                new_line['net_revenue'] += product_line['net_revenue']
                new_line['cogs'] += product_line['cogs']
            new_line['margin_1'] = new_line['net_revenue'] - product_line['cogs']
            grouped_list.append(new_line)
        return grouped_list

class SaleIncomeCostingAnalysisExcelFile(models.TransientModel):
    _name = 'sale.income.costing.analysis.excel.file'

    file = fields.Binary('File', readonly=True)
    name = fields.Char('File Name')
