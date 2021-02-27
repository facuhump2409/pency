import datetime
import unittest
from datetime import date

from model.pedido import get_basic_information, extract_guarniciones, Plato, extract_items_from_order


class MyTestCase(unittest.TestCase):
    simple_order = """PEDIDO: TIE-G0DXY
                    — [ 2 ] BBQ Ribs Guarnición: Puré de papas > $ 780,00
                    
                    Total: $ 780,00

                    Nombre y apellido: Facundo
                    Fecha de entrega: Entrega en 24hs! - Desde $250
                    Zona de entrega: Capital Federal
                    Dirección de entrega: Av. Santa Fé 3000
                    Forma de pago: Transferencia bancaria (Alias: SMARTFOODSOLUTIONS)
                    """
    guarnicion = "Guarnición: Puré de papas"
    order = "— [ 2 ] BBQ Ribs Guarnición: Puré de papas > $ 780,00"

    def test_getsBasicInformationCorrectly(self):
        result_dictionary = {}
        get_basic_information(result_dictionary, self.simple_order)
        expected_dictionary = {}
        expected_dictionary['Pedido'] = "TIE-G0DXY"
        expected_dictionary['Fecha de pedido'] = date.today().strftime("%d/%m/%Y")
        expected_dictionary['Tipo de envío'] = "24hs"
        expected_dictionary['Dirección'] = "Av. Santa Fé 3000"
        expected_dictionary['Medio de pago'] = "Transferencia bancaria"
        expected_dictionary['Cliente'] = "Facundo"
        expected_dictionary['Fecha de entrega'] = (datetime.date.today() + datetime.timedelta(days=1)).strftime("%d/%m")
        expected_dictionary['Rango Horario'] = '16 a 20hs'  # Esto esta seteado para todos
        self.assertEqual(expected_dictionary, result_dictionary)

    def test_extracts_guarniciones_correctly(self):
        expected_result = [Plato(1, 'Puré de papas'.upper())]
        result_guarniciones = extract_guarniciones(self.order)
        self.assertEqual(expected_result[0].name, result_guarniciones[0].name)
        self.assertEqual(expected_result[0].get_qty(), result_guarniciones[0].get_qty())

    def test_main_dish_correctly(self):
        expected_result = sorted([Plato(2, 'BBQ Ribs'), Plato(1, 'Puré de papas'.upper(), 2), Plato(2, 'Premium')],
                                 key=lambda x: x.name, reverse=True)
        result_order = sorted(extract_items_from_order(self.order), key=lambda x: x.name, reverse=True)
        self.assertTrue(expected_result[0] == result_order[0])


if __name__ == '__main__':
    unittest.main()
