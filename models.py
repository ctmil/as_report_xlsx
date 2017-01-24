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
			headings = ['Fecha Prevista','Fecha','Documento Origen','Ubicacion Destino','Proveedor','Tipo Entrega','Requerimientos','Producto','Marca','Cantidad',\
				'Unidad de Medida','Estado']
			columna = 0
			for heading in headings:
				sheet.write(0,columna,heading)
				columna += 1
			for obj in moves:
				# One sheet by partner
				bold = workbook.add_format({'bold': True})
				try:
					sheet.write(row, 0, obj.name)
					row += 1
				except:
					pass


StockMoveXlsx('report.stock.move.xlsx','stock.move')
