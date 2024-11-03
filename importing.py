import pandas as pd

# mandate =pd.read_excel("/home/leander/gei/faktura/abrechnung_24_q3/CC100438_abrechnung_Abr_YQ-2024-3_export(1).xlsx",sheet_name="Liste")
# df = pd.read_excel("/home/leander/gei/faktura/abrechnung_24_q3/CC100438_abrechnung_Abr_YQ-2024-3_export(1).xlsx",sheet_name="Liste")
class Data():
    def __init__(self,f_for_data_loading,f_for_template_loading):
        self.load_data = f_for_data_loading
        self.load_template = f_for_template_loading

        self.data = []
        self.template_for_export = []



def load_mandates(filepath):
    print(1)
    print(filepath)
def load_invoices(filepath):
    print(2)
    print(filepath)
def load_mandate_template(filepath):
    print(3)
    print(filepath)
def load_invoice_template(filepath):
    print(4)
    print(filepath)


mandates = Data(load_mandates,load_mandate_template)
invoices = Data(load_invoices,load_invoice_template)

