class Plato:
  def __init__(self, cantidad, nombre,multiplicador= None):
    self.nombre = nombre
    self.cantidad = cantidad
    self.multiplicador = 1 if multiplicador is None else multiplicador
  def get_cantidad(self):
    return self.cantidad * self.multiplicador
  def set_multiplicador(self,multiplicador):
    self.multiplicador = multiplicador
    return self

class Combo:
  def __init__(self, platos, cantidad, nombre):
    self.nombreCombo = nombre
    self.platos = [plato.set_multiplicador(cantidad) for plato in platos]
    self.cantidad = cantidad
