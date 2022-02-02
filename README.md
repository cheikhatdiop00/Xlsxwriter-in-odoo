# Xlsxwriter-in-odoo


by Cheikh-Ahmed-T diop- 
github -cheikhatdiop00
MAILing - ctd@moore.sn

odoo methode to generate xlsx report in odoo



  \\\  by cTD



  # migration du module private_budget sur Odoo v_14
def generate_xlsx_report(self, workbook, data, line):
      
        n = 0
        l=18
        p=15
        w=25
        for lines in line:
            n += 1
           
            format_colonne = workbook.add_format({'font_size': 11, 'align': 'vcenter', 'bold': True,'bg_color': '#FFC7CE'})
            format_titre = workbook.add_format({'font_size': 13, 'align': 'vcenter', 'bold': True})
            format_data =  workbook.add_format({'font_size': 11, 'align': 'vcenter', 'bold': False })
            
            format2 = workbook.add_format({'font_size': 10, 'align': 'vcenter', })
           
            sheet = workbook.add_worksheet('Suivi Budgetaire')
       
            sheet.set_column(7, 0, l)
            sheet.set_column(7, 1, l)
            sheet.set_column(7, 2, l)
            sheet.set_column(7, 3, l)
            sheet.set_column(7, 4, l)
            sheet.set_column(7, 5, l)
            
           
            sheet.set_column(0, 0, w)
          
           
            sheet.write(0, 0, 'Suivi Budgetaire - My Company - EUR', format_titre)
            
            sheet.write(4, 1, 'Periode:', format_colonne)
            sheet.write(4, 2, ('Du: %s Au: %s') % (lines.date_from, lines.date_to),format_data )
            
            
            sheet.write(7, 0, 'Poste budgétaire', format_colonne)
            
            sheet.write(7,1,'Compte analytique',format_colonne)
            sheet.write(7,2,'Montant prevu',format_colonne)
            sheet.write(7,3,'Montant engage',format_colonne)
            sheet.write(7,4,'Montant realise',format_colonne)
            sheet.write(7,5,'Montant disponible',format_colonne)
            
            row_xlsx = 8
            col_xlsx = 0
          
            for line_budgetaire in (lines.line_ids):
            
                sheet.write(row_xlsx, col_xlsx,  ('%s') % (line_budgetaire.general_budget),
                                                                                        format_data )
                sheet.write(row_xlsx, col_xlsx+1,  ('%s') % (line_budgetaire.account_analytic),
                                                                                            format_data )
                sheet.write(row_xlsx, col_xlsx+2,  ('%s') % (line_budgetaire.planned_amount),
                                                                                            format_data )
                sheet.write(row_xlsx, col_xlsx+3,  ('%s') % (line_budgetaire.engage_amount),
                                                                                            format_data )
                sheet.write(row_xlsx, col_xlsx+4,('%s') % (line_budgetaire.practical_amount),
                                                                                            format_data )
                sheet.write(row_xlsx,col_xlsx+5, ('%s') % (line_budgetaire.available_amount),
                                                                                            format_data)
               
                row_xlsx += 1

