import re

from openpyxl import *

from model.commons import load_products, add_atributes_to_excels
from model.pedido import Plato, get_basic_information, work_products

excels = []
# TODO abrir dos archivos, uno con data_only y otro con todas las formulas y escribir en el que tiene todas las formulas
orders_folder_year = "Pedidos/2021/012020"
clients_folder = orders_folder_year + "/Generador de pedidos/Pedidos/"
orders_generator = load_workbook(orders_folder_year + "/Generador de pedidos/Generador de pedidos.xlsx")
consolidated_orders_with_formulas = load_workbook("Pedidos/2021/012020/Pedidos consolidados 01-2021.xlsx")
consolidated_orders_data_only = load_workbook("Pedidos/2021/012020/Pedidos consolidados 01-2021.xlsx", data_only=True)
prices = consolidated_orders_data_only['Precios y Menú']
excels.extend([consolidated_orders_with_formulas, consolidated_orders_data_only, orders_generator])
products_dictionary = {}  # para saber cual es premium, guarnicion o daily

load_products(products_dictionary, prices)

messages = orders_generator["Mensajes de wapp"]


def process_orders(messages):
    message_row = 0
    for message in messages.iter_rows(values_only=True):
        message_row += 1
        if message[0] == 'Si' or message[0] == 'Procesado': continue
        client_order_data_only = load_workbook(clients_folder + "/Planilla Cliente.xlsx", data_only=True)
        client_order_with_formula = load_workbook(clients_folder + "/Planilla Cliente.xlsx")
        excel_dictionary = {}
        order = message[1]
        get_basic_information(message, excel_dictionary, order)
        items_regex = "(?i)— .+"
        items = re.findall(items_regex, order)
        order_products = []
        guarnicion_regex = "(?i)Guarnici.n(es)?"
        for item in items:
            qty_regex = "\[ \d \]"
            qty_product = re.search(qty_regex, item)
            qty = 1 if qty_product == None else int(re.sub("\[|\]", "", qty_product[0]).strip())
            product = re.sub(r"" + qty_regex + "|" + guarnicion_regex + ".+", "",
                             re.search("(?i)(?<=—).+(?=\>)", item)[0].strip())
            extra_brackets_info = "\s?(\(.+)?:"
            cleaner_item = re.search(guarnicion_regex + extra_brackets_info + " .+(?=>)", item)[0]
            guarniciones = re.sub(guarnicion_regex + extra_brackets_info, "", cleaner_item).strip()
            guarniciones = re.findall('[A-W][^A-W]*', guarniciones)
            guarniciones_list = []
            for guarnicion in guarniciones:
                nombre = re.sub(guarnicion_regex + "?: |X\d|,", "", guarnicion).strip()
                guarnicion_qty = re.search("X\d", guarnicion)
                final_qty = 1 if guarnicion_qty == None else int(re.sub("X", "", guarnicion_qty[0]).strip())
                guarniciones_list.append(Plato(final_qty, nombre))
            work_products(product, order_products, qty, guarniciones_list,products_dictionary)
        # excel_dictionary['Zona'] = re.search("(?i)(?<=Zona de entrega: ).+",order)[0]
        messages.cell(message_row, 1).value = "Si"
        add_atributes_to_excels([[client_order_data_only, client_order_with_formula],
                                 [consolidated_orders_data_only, consolidated_orders_with_formulas]], excel_dictionary,
                                order_products)
        client_order_with_formula.save(clients_folder + excel_dictionary['Pedido'] + excel_dictionary['Cliente'] + ".xlsx")
        client_order_with_formula.close()
        client_order_data_only.close()
    orders_generator.save(orders_folder_year + "/Generador de pedidos/Generador de pedidos.xlsx")
    consolidated_orders_with_formulas.save("Pedidos/2021/012020/Pedidos consolidados 01-2021.xlsx")
    for excel in excels:
        excel.close()


process_orders(messages)
# TODO terminar de ver que se hagan los combos
# TODO ARMAR TESTS URGENTEEE
