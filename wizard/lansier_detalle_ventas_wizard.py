# -*- encoding: utf-8 -*-

from odoo import models, fields, api
import xlsxwriter
import base64
import io
import logging

class LansierSalesDataillWizard(models.TransientModel):

    _name = 'lansier.sale.detail.wizard'

    date_from = fields.Date("Date from")
    date_to = fields.Date("Date to")
    name = fields.Char('File Name', size=32)
    archivo = fields.Binary('File', filters='.xls')

    def print_report(self):
        for w in self:
            dict = {}
            invoice_ids = self.env['account.move'].search([('date','>=', w.date_from),('date','<=', w.date_to)])

            f = io.BytesIO()
            workbook = xlsxwriter.Workbook(f)
            worksheet = workbook.add_worksheet('Sale detail')

            worksheet.write(0, 0, 'FECHA VENTA')
            worksheet.write(0, 1, 'FACT')
            worksheet.write(0, 2, 'NO. FACT')
            worksheet.write(0, 3, 'NOMBRE DEL CLIENTE')
            worksheet.write(0, 4, 'NIT')
            worksheet.write(0, 5, 'DIRECCION DE ENTREGA')
            worksheet.write(0, 6, 'NOMBRE COMERCIAL')
            worksheet.write(0, 7, 'CATEG.')
            worksheet.write(0, 8, 'CODIGO')
            worksheet.write(0, 9, 'PRODUCTO')
            worksheet.write(0, 10, 'LOTE')
            worksheet.write(0, 11, 'FECHA VENCIMIENTO')
            worksheet.write(0, 12, 'QTY')
            worksheet.write(0, 13, 'PRECIO  UNITARIO Q.')
            worksheet.write(0, 14, 'DSCTO')
            worksheet.write(0, 15, 'PRECIO CON DESCUENTO')
            worksheet.write(0, 16, 'TOTAL X LINEA')
            worksheet.write(0, 17, 'TOTAL X LINEA SIN IVA QTZ')
            worksheet.write(0, 18, 'TOTAL X LINEA SIN IVA USD $')
            worksheet.write(0, 19, 'N/E')
            worksheet.write(0, 20, 'CREDITO')
            worksheet.write(0, 21, 'VISITADOR MEDICO')

            row = 1
            if len(invoice_ids) > 0:
                for invoice in invoice_ids:
                    if len(invoice.invoice_line_ids) > 0:
                        for line in invoice.invoice_line_ids:
                            lot = ''
                            expiration_date = ''
                            if line.sale_line_ids:
                                move_id = self.env["stock.move"].search([("sale_line_id","=", line.sale_line_ids.id)])
                                if move_id and move_id.lot_ids:
                                    lot = move_id.lot_ids.name
                                    expiration_date = str(move_id.lot_ids.expiration_date)
                            price_total_discount = 0
                            if line.discount > 0:
                                price_total_discount = line.price_total / line.quantity
                            price_usd = linea.price_subtotal if line.currency_id.id != line.company_id.currency_id.id else 0
                            credit = linea.move_id.invoice_payment_term_id.line_ids.nb_days if linea.move_id.invoice_payment_term_id else 0
                            medic = ''
                            worksheet.write(row, 0, str(line.invoice_date))
                            worksheet.write(row, 0, 'FACT')
                            worksheet.write(row, 0, line.move_id.fel_numero)
                            worksheet.write(row, 0, line.move_id.partner_id.name)
                            worksheet.write(row, 0, line.move_id.partner_id.vat)
                            worksheet.write(row, 0, line.move_id.partner_id.x_studio_direccion_entrega)
                            worksheet.write(row, 0, line.move_id.partner_id.x_studio_nombre_comercial)
                            worksheet.write(row, 0, line.product_id.categ_id.name)
                            worksheet.write(row, 0, line.move_id.partner_id.ref)
                            worksheet.write(row, 0, line.product_id.name)
                            worksheet.write(row, 0, lot)
                            worksheet.write(row, 0, expiration_date)
                            worksheet.write(row, 0, line.quantity)
                            worksheet.write(row, 0, line.price_unit)
                            worksheet.write(row, 0, line.discount)
                            worksheet.write(row, 0, price_total_discount)
                            worksheet.write(row, 0, line.price_total)
                            worksheet.write(row, 0, line.amount_currency / 1.12)
                            worksheet.write(row, 0, price_usd)
                            worksheet.write(row, 0, 'D')
                            worksheet.write(row, 0, credit)
                            worksheet.write(row, 0, medic)
                            
            workbook.close()
            data = base64.b64encode(f.getvalue())
            self.write({'archivo':data, 'name':'sale_detail.xlsx'})
        return {
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'lansier.sale.detail.wizard',
            'res_id': self.id,
            'view_id': False,
            'type': 'ir.actions.act_window',
            'target': 'new',
        }