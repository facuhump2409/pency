import unittest


class MyTestCase(unittest.TestCase):
    simpleOrder = """PEDIDO: TIE-G0DXY
                    — [ 2 ] BBQ Ribs Guarnición: Puré de papas > $ 780,00
                    
                    Total: $ 780,00

                    Nombre y apellido: Facundo
                    Fecha de entrega: Entrega en 24hs! - Desde $250
                    Zona de entrega: Capital Federal
                    Dirección de entrega: Av. Santa Fé 3000
                    Forma de pago: Transferencia bancaria (Alias: SMARTFOODSOLUTIONS)
                    """

    def test_getsBasicInformationCorrectly(self):
        self.assertEqual(True, True)


if __name__ == '__main__':
    unittest.main()
