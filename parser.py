from openpyxl import *
import re
from datetime import date

#Spreadsheet.client_encoding = 'windows-1251:utf-8'
orders_folder_year = "Pedidos/2021/012020"
clients_folder = orders_folder_year + "/Generador de pedidos/Pedidos"
orders_generator = load_workbook(orders_folder_year + "/Generador de pedidos/Generador de pedidos.xlsx")

consolidated_orders = load_workbook("Pedidos/2021/012020/Pedidos consolidados 01-2021.xlsx")
client_order = load_workbook(clients_folder + "/Planilla cliente.xlsx")
prices = consolidated_orders['Precios y Menú']

# costo_de_envio = 0
# envio_gratis = 0
products = {} #para saber cual es premium, guarnicion o daily
def load_products(products):
    for row in prices.iter_rows(values_only=True):
        if (row[8] != None and "=" not in row[8]): products[row[8]] = row[7]
        # if row[2] == "Costo de envío": costo_de_envio = row[4]
        # if row[2] == "Envio Gratis": envio_gratis = row[4]
load_products(products)


messages = orders_generator["Mensajes WPP"]
tipos_de_envios = ["Programado", "24 hs", "Retiro en Local"]


def set_excel_attributes(excel_file, excel_dictionary):
    for key, value in excel_dictionary.items():



for message in messages.iter_rows(values_only= True):
    if (message[0] == 'Si' or message[0] == 'Procesado'): continue #or message[1] == None
    excel_dictionary = {}
    order = message[1]
    excel_dictionary['Fecha de pedido'] = date.today().strftime("%d/%m/%Y")
    excel_dictionary['Tipo de envío'] = 'Programada' #TODO cambiar esto, por ahora lo dejo hardcodeado
    #total = re.search("(?<=Total\: ).+",order)[0]
    excel_dictionary['Pedido'] = re.search("(?i)Pedido:\K.+",order)[0] #TODO evitar esta repeticion de logica con un diccionario y una lista de variables que tienen misma regex
    items_regex = "(?i)— .+"
    items = order.findall(items_regex)
    items_dictionary = {}
    for item in items:
        qty_regex = "\[ \d \]"
        qty_product = re.search(qty_regex,item)
        qty =  1 if qty_product == None else qty_product[0].gsub("\[|\]","").strip()
        product = item.search("(?i)(?<=—( " + qty_regex + ")?).+(?=\>)")[0].strip()
        guarniciones = item.search("Guarnici.n(es (\(.+))?:\K.+(?=>)").strip().split(",")
        items_dictionary[product] = qty
    excel_dictionary['Fecha de entrega'] = re.search("\d{2}/\d{2}",order)[0] #TODO VER SI ESTO ROMPE en algun caso
    try:
        excel_dictionary['Rango Horario'] = re.search("\d{2}hs a \d{2}hs",order)[0]
    except:
        excel_dictionary['Rango Horario'] = "17 a 20hs" #Para 24 hs
    # excel_dictionary['Zona'] = re.search("(?i)(?<=Zona de entrega: ).+",order)[0]
    #Maneras del delivery_date -> Martes 12/01 , Entrega en 24hs!, Retiro en local
    excel_dictionary['Direccion'] = re.search("(?i)(?<=Direcci.n de entrega: ).+",order)[0]
    payment_method = re.search("(?i)(?<=Forma de pago: ).+",order)[0].strip()
    excel_dictionary['Medio de pago'] = re.sub("\(.+\)", "", payment_method).strip()
    excel_dictionary['Cliente'] = re.search("(?i)(?<=Nombre y Apellido: ).+")
    message[0] = "Si"
    set_excel_attributes(client_order,excel_dictionary)
    client_order.save(excel_dictionary['Pedido'] + excel_dictionary['Cliente'])
    #Setear en Si el procesado
#orders_generator.write '26-12 Planilla de pedidos.xls'
#orders_generator.save("26-12 Planilla de pedidos.xlsx")
p items_dictionary
#def work_excel
