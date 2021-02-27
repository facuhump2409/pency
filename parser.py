from openpyxl import *

from files import clients_folder, consolidated_orders_data_only, consolidated_orders_with_formulas, orders_generator, \
    orders_folder_year, excels, year, orders_messages
from model.commons import add_atributes_to_excels
from model.pedido import get_basic_information, extract_items_from_order


def process_orders(messages):
    message_row = 0
    for message in messages.iter_rows(values_only=True):
        # try:
        message_row += 1
        if message[0] == 'Si' or message[0] == 'Procesado': continue
        client_order_data_only = load_workbook(clients_folder + "/Planilla Cliente.xlsx", data_only=True)
        client_order_with_formula = load_workbook(clients_folder + "/Planilla Cliente.xlsx")
        excel_dictionary = {}
        order = message[1]
        get_basic_information(excel_dictionary, order)
        order_products = extract_items_from_order(order)
        # excel_dictionary['Zona'] = re.search("(?i)(?<=Zona de entrega: ).+",order)[0]
        messages.cell(message_row, 1).value = "Si"
        add_atributes_to_excels([[client_order_data_only, client_order_with_formula],
                                 [consolidated_orders_data_only, consolidated_orders_with_formulas]],
                                excel_dictionary,
                                order_products)
        client_order_with_formula.save(
            clients_folder + excel_dictionary['Pedido'] + "_" + excel_dictionary['Cliente'] + ".xlsx")
        client_order_with_formula.close()
        client_order_data_only.close()
    # except:
    # messages.cell(message_row, 1).value = "ERROR"
    orders_generator.save(orders_folder_year + "/Generador de pedidos/Generador de pedidos.xlsx")
    consolidated_orders_with_formulas.save(orders_folder_year + "Pedidos consolidados 01-" + year + ".xlsx")
    for excel in excels:
        excel.close()


process_orders(orders_messages)
# TODO ver convensiones de nombres de hojas de excel
