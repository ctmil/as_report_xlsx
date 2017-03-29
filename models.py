from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx

class PartnerXlsx(ReportXlsx):
    
	def generate_xlsx_report(self, workbook, data, partners):
		for obj in partners:
			report_name = obj.name
			# One sheet by partner
			sheet = workbook.add_worksheet(report_name[:31])
			bold = workbook.add_format({'bold': True})
			sheet.write(0, 0, obj.name, bold)


PartnerXlsx('report.res.partner.xlsx','res.partner')

class StockMoveXlsx(ReportXlsx):
    
	def generate_xlsx_report(self, workbook, data, moves):
		if moves:
			report_name = moves[0].name
			sheet = workbook.add_worksheet(report_name[:31])
			row = 1
			headings = ['Fecha Prevista','Fecha','Documento Origen','Ubicacion Destino','Proveedor','Tipo Entrega','Entregar a','Requerimientos','Producto','Marca','Cantidad',\
				'Unidad de Medida','Estado']
			columna = 0
			for heading in headings:
				sheet.write(0,columna,heading)
				columna += 1
			for obj in moves:
				sheet.write(row, 0, str(obj.date_expected)[:10])
				sheet.write(row, 1, str(obj.date)[:10])
				#sheet.write(row, 2, obj.picking_id.name)
				sheet.write(row, 2, obj.origin)
				sheet.write(row, 3, obj.location_dest_id.complete_name)
				sheet.write(row, 4, obj.picking_partner_id.name)
				if obj.tipo_entrega == 'proveedor':
					sheet.write(row, 5, 'Retiramos de deposito del proveedor')
				else:
					sheet.write(row, 5, obj.tipo_entrega)
				sheet.write(row, 6, obj.picking_type_id.name)
				sheet.write(row, 7, obj.request_name)
				sheet.write(row, 8, obj.product_id.name)
				sheet.write(row, 9, obj.brand_id.name)
				sheet.write(row, 10, obj.product_uom_qty)
				sheet.write(row, 12, obj.product_uom.name)
				sheet.write(row, 12, obj.state)
				row += 1


StockMoveXlsx('report.stock.move.xlsx','stock.move')
