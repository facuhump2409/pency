import datetime
import re
from datetime import date

from files import products_dictionary
from model.commons import get_rid_between_brackets
from model.coordinates import find_cell_with_value

tipos_de_envios = ["Programado", "24hs", "Retiro en Local"]
guarnicion_regex = "(?i)Guarnici.n(es)?"


class Plato:
    def __init__(self, qty, name, multiplier=None):
        self.name = name.strip()
        self.qty = qty
        self.multiplier = 1 if multiplier is None else multiplier

    def get_qty(self):
        return self.qty * self.multiplier

    def set_multiplier(self, multiplier):
        self.multiplier = multiplier
        return self

    def __eq__(self, other):
        if isinstance(other, Plato):
            return (self.name == other.name and self.get_qty() == other.get_qty())
        return False


class Combo:
    def __init__(self, platos, qty, name):
        self.name = name
        self.platos = [plato.set_multiplier(qty) for plato in platos]
        self.qty = qty

    def get_platos(self):
        return self.platos + [Plato(self.qty, self.name)]


def get_tipo_de_envio(texto):
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


def work_products(product, order_products, order_qty, guarniciones):
    if "Combo" in product:
        combo_name = re.search("(?i)Combo (Surtido|Premium|Daily)", product)[0].strip()
        product = re.search("(?i)(?<=:).+(?=-)", product)[0]
        products = product.split(",")
        products_list = []
        divided_products = divide_surtido_products(products)
        for product in divided_products:
            nombre = re.sub("(?i)(X\d|(Daily|Premium)|:)", "", product.strip()).strip()
            qty = re.search("X\d", product)
            final_qty = 1 if qty is None else int(re.sub("X", "", qty[0]).strip())
            products_list.append(Plato(final_qty, nombre))
        products_list.extend(guarniciones)
        combo = Combo(products_list, order_qty, combo_name)
        order_products.extend(combo.get_platos())
    else:
        guarniciones = list(map(lambda x: x.set_multiplier(order_qty), guarniciones))
        # add_type_of_order(order_qty, product, order_products, consolidated_orders_data_only)
        product_type = products_dictionary[product.strip().upper()]
        order_products.append(Plato(order_qty, product_type))
        order_products.append(Plato(order_qty, product.strip()))
        order_products.extend(guarniciones)


def get_basic_information(excel_dictionary, order):
    excel_dictionary['Fecha de pedido'] = date.today().strftime("%d/%m/%Y")
    excel_dictionary['Tipo de envío'] = get_tipo_de_envio(re.search("(?i)(?<=Fecha de entrega: ).+", order)[0])
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


def extract_guarniciones(item):
    extra_brackets_info = "\s?(\(.+)?:"
    cleaner_item = re.search(guarnicion_regex + extra_brackets_info + " .+(?=>)", item)[0]
    guarniciones = re.sub(guarnicion_regex + extra_brackets_info, "", cleaner_item).strip()
    upper_case_guarniciones = guarniciones.upper()
    found_guarniciones = []
    for product in products_dictionary.keys():
        if product.upper() in upper_case_guarniciones:
            found_guarniciones.append(product)
    guarniciones_list = []
    for guarnicion in found_guarniciones:
        nombre = guarnicion.strip()  # re.sub(guarnicion_regex + "?: |X\d|,", "", guarnicion).strip()
        guarnicion_qty = re.search("(?i)(?<=" + guarnicion + ") X\d((?=,)|(?=$))",
                                   guarniciones)  # re.search("X\d", guarnicion)
        final_qty = 1 if guarnicion_qty == None else int(re.sub("X", "", guarnicion_qty[0]).strip())
        guarniciones_list.append(Plato(final_qty, nombre))
    return guarniciones_list


def extract_items_from_order(order):
    items_regex = "(?i)— .+"
    items = re.findall(items_regex, order)
    order_products = []
    for item in items:
        qty_regex = "\[ \d \]"
        qty_product = re.search(qty_regex, item)
        qty = 1 if qty_product == None else int(re.sub("\[|\]", "", qty_product[0]).strip())
        product = re.sub(r"" + qty_regex + "|" + guarnicion_regex + ".+", "",
                         re.search("(?i)(?<=—).+(?=\>)", item)[0].strip())
        extra_brackets_info = "\s?(\(.+)?:"
        cleaner_item = re.search(guarnicion_regex + extra_brackets_info + " .+(?=>)", item)[0]
        dirty_guarniciones = re.sub(guarnicion_regex + extra_brackets_info, "", cleaner_item).strip()
        # guarniciones = re.findall('[A-W][^A-W]*', guarniciones)
        work_products(product, order_products, qty, extract_guarniciones(item))
    return order_products
