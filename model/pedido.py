import datetime
import re
from datetime import date

from model.commons import get_rid_between_brackets
from model.coordinates import find_cell_with_value

tipos_de_envios = ["Programado", "24hs", "Retiro en Local"]


class Plato:
    def __init__(self, qty, name, multiplier=None):
        self.name = name
        self.qty = qty
        self.multiplier = 1 if multiplier is None else multiplier

    def get_qty(self):
        return self.qty * self.multiplier

    def set_multiplier(self, multiplier):
        self.multiplier = multiplier
        return self


class Combo:
    def __init__(self, platos, qty, name):
        self.name = name
        self.platos = [plato.set_multiplier(qty) for plato in platos]
        self.qty = qty

    def get_platos(self):
        return self.platos + [Plato(self.qty, self.name)]


def get_tipo_de_envio(tipos_de_envios, texto):
    for tipo_de_envio in tipos_de_envios:
        if tipo_de_envio.upper() in texto.upper():
            return tipo_de_envio
    return 'Programado'  # Es el valor default


def add_type_of_order(order_qty, product_name, order_products, consolidated_orders_data_only):
    worksheet = consolidated_orders_data_only["Precios y Menú"]
    cell = find_cell_with_value(product_name, worksheet)
    product_type = worksheet.cell(cell.row, cell.column - 1).value
    order_products.append(Plato(order_qty, product_type))


def divide_surtido_products(products):
    new_products = []
    for product in products:
        if "-" in product:
            without_brackets = re.sub("(?i)(Daily|Premium|:)", "", get_rid_between_brackets(product))
            separated_products = without_brackets.split("-")
            new_products.extend(separated_products)
        else:
            new_products.append(product)
    return new_products


def work_products(product, order_products, order_qty, guarniciones, consolidated_orders_data_only):
    if "Combo" in product:
        # combo_name = re.search("(?i).+(?=plato)", product)[0].strip()
        combo_name = re.search("(?i)Combo (Surtido|Premium|Daily)", product)[0].strip()
        # combo_name = re.sub("\d\.|\(.+\)", "", combo_name).strip()
        # product_without_between_brackets = get_rid_between_brackets(product)
        product = re.search("(?i)(?<=:).+(?=-)", product)[0]
        products = product.split(",")
        products_list = []
        divided_products = divide_surtido_products(products)
        for product in divided_products:
            nombre = re.sub("(?i)(X\d|(Daily|Premium)|:)", "", product).strip()
            qty = re.search("X\d", product)
            final_qty = 1 if qty is None else int(re.sub("X", "", qty[0]).strip())
            products_list.append(Plato(final_qty, nombre))
        products_list.extend(guarniciones)
        combo = Combo(products_list, order_qty, combo_name)
        order_products.extend(combo.get_platos())
    else:
        guarniciones = list(map(lambda x: x.set_multiplier(order_qty), guarniciones))
        add_type_of_order(order_qty, product, order_products, consolidated_orders_data_only)
        order_products.append(Plato(order_qty, product.strip()))
        order_products.extend(guarniciones)


def get_basic_information(message, excel_dictionary, order):
    excel_dictionary['Fecha de pedido'] = date.today().strftime("%d/%m/%Y")
    excel_dictionary['Tipo de envío'] = get_tipo_de_envio(tipos_de_envios,
                                                          re.search("(?i)(?<=Fecha de entrega: ).+", order)[0])
    # total = re.search("(?<=Total\: ).+",order)[0]
    excel_dictionary['Pedido'] = re.search("(?i)(?<=Pedido: ).+", order)[0]
    excel_dictionary['Rango Horario'] = "16 a 20hs"  # Para 24 hs
    direccion = re.search("(?i)(?<=Direcci.n de entrega: ).+", order)[0]
    excel_dictionary['Dirección'] = re.sub("\*", "", direccion).strip()
    payment_method = re.search("(?i)(?<=Forma de pago: ).+", order)[0].strip()
    excel_dictionary['Medio de pago'] = re.sub("\(.+\)", "", payment_method).strip()
    excel_dictionary['Cliente'] = re.search("(?i)(?<=Nombre y Apellido: ).+", order)[0]
    try:
        excel_dictionary['Fecha de entrega'] = re.search("\d{2}/\d{2}", order)[0]
    except:
        excel_dictionary['Fecha de entrega'] = (datetime.date.today() + datetime.timedelta(days=1)).strftime(
            "%d/%m")
