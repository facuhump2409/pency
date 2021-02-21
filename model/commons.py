import re

from model.coordinates import set_value_in_empty_space


def load_products(products, prices):
    for row in prices.iter_rows(values_only=True):
        if (row[8] != None and "=" not in row[8] and row[7] != "Producto"): products[row[8].upper()] = row[7]
        # if row[2] == "Costo de envío": costo_de_envio = row[4]
        # if row[2] == "Envio Gratis": envio_gratis = row[4]


def set_excel_attributes(excel_file_with_data, excel_file_with_formula, excel_dictionary, order_id):
    for key, value in excel_dictionary.items():
        set_value_in_empty_space(excel_file_with_data, excel_file_with_formula, key, value, order_id)


def set_products_to_excel(excel_file_with_data, excel_file_with_formula, products, order_id):
    for product in products:
        set_value_in_empty_space(excel_file_with_data, excel_file_with_formula, product.name, product.get_qty(),
                                 order_id)

def add_atributes_to_excels(excels, excel_dictionary, order_products):
    # First excel has data, and second excel has formulas
    order_id = excel_dictionary['Pedido']
    for excel in excels:
        set_excel_attributes(excel[0], excel[1], excel_dictionary, order_id)
        set_products_to_excel(excel[0], excel[1], order_products, order_id)

def get_rid_between_brackets(string):
    return re.sub("\(.+\)","",string)