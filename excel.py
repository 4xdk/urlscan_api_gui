import pdb
from openpyxl import Workbook, load_workbook
import riskiq

def load_template():
    excel_file = load_workbook(filename="suspicious domain.xlsx")
    return excel_file

'''
Put all the prepared data into an excel sheet
data e.g. {'domain': 'atos.net', 'ns': ['ns1.atos.net', 'ns2.atos.net'], 'mx': ['mx1.atos.net', 'mx2.atos.net'], 'whois': 'RAW TEXT HERE', 
'ip': '127.0.0.1', 'ip_loc': 'Amsterdam', 'asn': 'AS 1234', 'image_blob': '...'}
'''
def prepare(data):
    excel_file = load_template()
    template_sheet = excel_file.active

    current_domain_sheet = Workbook.copy_worksheet(excel_file, template_sheet)
    current_domain_sheet.title = domain
    current_domain_sheet['B3'] = data['domain']
    current_domain_sheet['B4'] = "\n".join(data['ns'])
    current_domain_sheet['B6'] = "\n".join(data['mx'])
    current_domain_sheet['B7'] = "\n".join(data['whois'])

    # TODO join the 3 below to proper format
    current_domain_sheet['B5'] = "\n".join(data['ip'])
    current_domain_sheet['B5'] = "\n".join(data['ip_loc'])
    current_domain_sheet['B5'] = "\n".join(data['asn'])
    # TODO extract the image blob data to excel file
#    current_domain_sheet['C4'] = "\n".join(data['image_blob'])


    Workbook.remove(excel_file, template_sheet) # remove the template sheet
    excel_file.save('result.xlsx')
    
    return {'status': 'success', 'message': 'Saved to result.xlxs'}
