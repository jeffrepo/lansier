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

            worksheet.write(0, 0, 'Detale venta')
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