import datetime
import re
from datetime import date

from openpyxl import *

# Spreadsheet.client_encoding = 'windows-1251:utf-8'
from model.pedido import Plato, Combo

orders_folder_year = "Pedidos/2021/012020"
clients_folder = orders_folder_year + "/Generador de pedidos/Pedidos"
orders_generator = load_workbook(orders_folder_year + "/Generador de pedidos/Generador de pedidos.xlsx")

consolidated_orders = load_workbook("Pedidos/2021/012020/Pedidos consolidados 01-2021.xlsx", data_only=True)
prices = consolidated_orders['Precios y Menú']

# costo_de_envio = 0
# envio_gratis = 0
products = {}  # para saber cual es premium, guarnicion o daily


def load_products(products):
    for row in prices.iter_rows(values_only=True):
        if (row[8] != None and "=" not in row[8] and row[7] != "Producto"): products[row[8]] = row[7]
        # if row[2] == "Costo de envío": costo_de_envio = row[4]
        # if row[2] == "Envio Gratis": envio_gratis = row[4]


load_products(products)

messages = orders_generator["Mensajes de wapp"]
tipos_de_envios = ["Programado", "24hs", "Retiro en Local"]


class Coordinate:
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def increment_column(self):
        self.y = self.y + 1


def get_free_space_cell(cell, worksheet,order_id):
    # TODO ver tema de si ya tiene un numero que se frene en esa posicion
    current_value = cell.value
    coordinate = Coordinate(cell.row, cell.column)
    while current_value is not None and current_value != order_id:
        coordinate.increment_column()
        current_value = worksheet.cell(coordinate.x, coordinate.y).value
    return coordinate


def set_value_in_empty_space(excel_file, name, value,order_id):
    try:
        worksheet = excel_file['Pedido']
    except:
        worksheet = excel_file['Pedido Pency']
    for row in worksheet.iter_rows():
        for cell in row:
            if (cell.value != None and isinstance(cell.value, str)): # No sea Null y sea string
                if cell.value.upper() == "PEDIDO":
                    free_coordinate = get_free_space_cell(cell, worksheet,order_id)
                if cell.value.upper() == name.upper().strip():
                    coordinate = Coordinate(cell.row,free_coordinate.y)
                    # coordinate = get_free_space_cell(cell, worksheet)
                    # chr(ord('a') + 1) # para conseguir proximo caracter
                    # cell.row -> te devuelve la fila
                    # cell.column -> te devuelve la columna
                    # Si es None es que no hay nada, habria que conseguir la proxima letra del abecedario de esa columna que esta vacia
                    actual_value = worksheet.cell(coordinate.x, coordinate.y).value or 0
                    if actual_value != 0:
                        worksheet.cell(coordinate.x, coordinate.y).value = int(value) + int(actual_value)
                    else:
                        worksheet.cell(coordinate.x, coordinate.y).value = value
                    return



def set_excel_attributes(excel_file, excel_dictionary,order_id):
    for key, value in excel_dictionary.items():
        set_value_in_empty_space(excel_file, key, value,order_id)


def set_products_to_excel(excel_file, products,order_id):
    for product in products:
        set_value_in_empty_space(excel_file, product.name, product.get_qty(),order_id)


message_row = 0


def add_atributes_to_excels(excels, excel_dictionary, order_products):
    order_id = excel_dictionary['Pedido']
    for excel in excels:
        set_excel_attributes(excel, excel_dictionary,order_id)
        set_products_to_excel(excel, order_products,order_id)


def get_tipo_de_envio(tipos_de_envios, texto):
    for tipo_de_envio in tipos_de_envios:
        if tipo_de_envio.upper() in texto.upper():
            return tipo_de_envio
    return 'Programado'  # Es el valor default


def work_products(product,order_products,order_qty):
    if "Combo" in product:
        combo_name = re.search("(?i).+(?=plato)",product)[0].strip()
        combo_name = re.sub("\d\.|\(.+\)","",combo_name).strip()
        product = re.search("(?i)(?<=:).+(?=-)",product)[0]
        products = product.split(",")
        for product in products:
            nombre = re.sub("X\d","",product)
            qty = re.search("X\d", product)
            final_qty = 1 if qty == None else int(re.sub("X", "", qty[0]).strip())
            order_products.append(Plato(final_qty, nombre))
        combo = Combo(order_products,order_qty,combo_name)
        return combo.get_platos()
    else:
        return order_products.append(Plato(order_qty, product))




for message in messages.iter_rows(values_only=True):
    message_row += 1
    if (message[0] == 'Si' or message[0] == 'Procesado'): continue  # or message[1] == None
    client_order = load_workbook(clients_folder + "/Planilla Cliente.xlsx", data_only=True)
    excel_dictionary = {}
    order = message[1]
    excel_dictionary['Fecha de pedido'] = date.today().strftime("%d/%m/%Y")
    excel_dictionary['Tipo de envío'] = get_tipo_de_envio(tipos_de_envios,
                                                          re.search("(?i)(?<=Fecha de entrega: ).+", order)[0])
    # total = re.search("(?<=Total\: ).+",order)[0]
    excel_dictionary['Pedido'] = re.search("(?i)(?<=Pedido: ).+", order)[0]
    excel_dictionary['Rango Horario'] = "16 a 20hs"  # Para 24 hs
    direccion = re.search("(?i)(?<=Direcci.n de entrega: ).+", order)[0]
    excel_dictionary['Dirección'] = re.sub("\*","",direccion).strip()
    payment_method = re.search("(?i)(?<=Forma de pago: ).+", order)[0].strip()
    excel_dictionary['Medio de pago'] = re.sub("\(.+\)", "", payment_method).strip()
    excel_dictionary['Cliente'] = re.search("(?i)(?<=Nombre y Apellido: ).+", order)[0]
    try:
        excel_dictionary['Fecha de entrega'] = re.search("\d{2}/\d{2}", order)[0]
    except:
        excel_dictionary['Fecha de entrega'] = (datetime.date.today() + datetime.timedelta(days=1)).strftime("%d/%m")
    items_regex = "(?i)— .+"
    items = re.findall(items_regex, order)
    order_products = []
    for item in items:
        qty_regex = "\[ \d \]"
        qty_product = re.search(qty_regex, item)
        qty = 1 if qty_product == None else int(re.sub("\[|\]", "", qty_product[0]).strip())
        guarnicion_regex = "(?i)Guarnici.n(es)?"
        product = re.sub(r"" + qty_regex + "|" + guarnicion_regex + ".+", "",
                         re.search("(?i)(?<=—).+(?=\>)", item)[0].strip())
        #order_products.append(Plato(qty, product))
        guarniciones = re.sub(guarnicion_regex + "?:", "",
                              re.search(guarnicion_regex + "?(\(.+)?: .+(?=>)", item)[0]).strip()
        guarniciones = re.findall('[A-W][^A-W]*', guarniciones)
        orders_guarnicion = []
        for guarnicion in guarniciones:
            nombre = re.sub(guarnicion_regex + "?: |X\d|,", "", guarnicion).strip()
            guarnicion_qty = re.search("X\d", guarnicion)
            final_qty = 1 if guarnicion_qty == None else int(re.sub("X", "", guarnicion_qty[0]).strip())
            order_products.append(Plato(final_qty, nombre))
        final_products = work_products(product,order_products,qty)
    # excel_dictionary['Zona'] = re.search("(?i)(?<=Zona de entrega: ).+",order)[0]
    messages.cell(message_row, 1).value = "Si"
    add_atributes_to_excels([client_order, consolidated_orders], excel_dictionary, final_products)
    client_order.save(excel_dictionary['Pedido'] + excel_dictionary['Cliente'] + ".xlsx")
    client_order.close()
#orders_generator.save(orders_folder_year + "/Generador de pedidos/Generador de pedidos.xlsx")
#consolidated_orders.save("Pedidos/2021/012020/Pedidos consolidados 01-2021.xlsx")
consolidated_orders.close()
orders_generator.close()
# TODO terminar de ver que se hagan los combos