# -*- encoding: utf-8 -*-

from odoo import models, fields, api
import xlsxwriter
import base64
import io
import logging
import math

class LansierSalesDataillWizard(models.TransientModel):

    _name = 'lansier.sale.detail.wizard'

    date_from = fields.Date("Date from")
    date_to = fields.Date("Date to")
    name = fields.Char('File Name', size=32)
    archivo = fields.Binary('File', filters='.xls')

    def print_report(self):
        for w in self:
            dict = {}
            invoice_ids = self.env['account.move'].search([('invoice_date','>=', w.date_from),('invoice_date','<=', w.date_to),('move_type','in',['out_invoice','out_refund']),('state','=','posted')])
            logging.warning(invoice_ids)
            
            f = io.BytesIO()
            workbook = xlsxwriter.Workbook(f)
            worksheet = workbook.add_worksheet('Sale detail')
            f1=workbook.add_format()
            f1.set_bold(True)
            worksheet.write(0, 0, 'FECHA VENTA', f1)
            worksheet.write(0, 1, 'FACT', f1)
            worksheet.write(0, 2, 'NO. FACT', f1)
            worksheet.write(0, 3, 'NOMBRE DEL CLIENTE', f1)
            worksheet.write(0, 4, 'NIT', f1)
            worksheet.write(0, 5, 'DIRECCION DE ENTREGA', f1)
            worksheet.write(0, 6, 'NOMBRE COMERCIAL', f1)
            worksheet.write(0, 7, 'CATEG.', f1)
            worksheet.write(0, 8, 'CODIGO', f1)
            worksheet.write(0, 9, 'PRODUCTO', f1)
            worksheet.write(0, 10, 'LOTE', f1)
            worksheet.write(0, 11, 'FECHA VENCIMIENTO', f1)
            worksheet.write(0, 12, 'QTY', f1)
            worksheet.write(0, 13, 'PRECIO  UNITARIO Q.', f1)
            worksheet.write(0, 14, 'DSCTO', f1)
            worksheet.write(0, 15, 'PRECIO CON DESCUENTO', f1)
            worksheet.write(0, 16, 'TOTAL X LINEA', f1)
            worksheet.write(0, 17, 'TOTAL X LINEA SIN IVA QTZ', f1)
            worksheet.write(0, 18, 'TOTAL X LINEA SIN IVA USD $', f1)
            worksheet.write(0, 19, 'N/E', f1)
            worksheet.write(0, 20, 'CREDITO', f1)
            worksheet.write(0, 21, 'VISITADOR MEDICO', f1)
            worksheet.write(0, 22, 'UNIDADES BONIFICADAS', f1)

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
                                    lot = move_id.lot_ids[0].name
                                    expiration_date = str(move_id.lot_ids[0].expiration_date)
                            price_total_discount = 0
                            if line.discount > 0:
                                price_total_discount = line.price_total / line.quantity
                            #Este se utilzaba en al linea 91    
                            #price_usd = line.price_subtotal if line.currency_id.id != line.company_id.currency_id.id else 0
                            credit = line.move_id.invoice_payment_term_id.line_ids.nb_days if line.move_id.invoice_payment_term_id else 0
                            medic = line.move_id.partner_id.x_studio_vendedor_lansier if line.move_id.partner_id.x_studio_vendedor_lansier else ''
                            unidades_bonificadas = ((line.discount / 100) * line.quantity) if line.discount > 0 else 0
                            cantidad = line.quantity*-1 if line.move_id.move_type == 'out_refund' else line.quantity
                            worksheet.write(row, 0, str(line.invoice_date))
                            worksheet.write(row, 1, line.move_id.journal_id.tipo_factura)
                            worksheet.write(row, 2, line.move_id.fel_numero)
                            worksheet.write(row, 3, line.move_id.partner_id.name)
                            worksheet.write(row, 4, line.move_id.partner_id.vat)
                            worksheet.write(row, 5, line.move_id.partner_id.x_studio_direccion_de_entrega)
                            worksheet.write(row, 6, line.move_id.partner_id.x_studio_nombre_comercial)
                            worksheet.write(row, 7, line.product_id.categ_id.name)
                            worksheet.write(row, 8, line.move_id.partner_id.ref)
                            worksheet.write(row, 9, line.product_id.name)
                            worksheet.write(row, 10, lot)
                            worksheet.write(row, 11, expiration_date)
                            worksheet.write(row, 12, cantidad)
                            worksheet.write(row, 13, line.price_unit)
                            worksheet.write(row, 14, line.discount)
                            worksheet.write(row, 15, price_total_discount)
                            worksheet.write(row, 16, line.price_total)
                            worksheet.write(row, 17, (line.price_total / 1.12))
                            linea_sin_iva_usd = (((line.price_total / 1.12)*-1)/7.8 ) * -1
                            worksheet.write(row, 18, linea_sin_iva_usd)
                            worksheet.write(row, 19, 'D')
                            worksheet.write(row, 20, credit)
                            worksheet.write(row, 21, medic)
                            worksheet.write(row, 22, math.ceil(unidades_bonificadas)),
                            row += 1
                            
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


