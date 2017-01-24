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
			for obj in moves:
				# One sheet by partner
				bold = workbook.add_format({'bold': True})
			sheet.write(0, 0, obj.name, bold)


StockMoveXlsx('report.productos.recibir.xlsx','stock.move')
