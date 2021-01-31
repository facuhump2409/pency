class Plato:
  def __init__(self, qty, name,multiplier= None):
    self.name = name
    self.qty = qty
    self.multiplier = 1 if multiplier is None else multiplier
  def get_qty(self):
    return self.qty * self.multiplier
  def set_multiplier(self,multiplier):
    self.multiplier = multiplier
    return self

class Combo:
  def __init__(self, platos, qty, name):
    self.name = name
    self.platos = [plato.set_multiplier(qty) for plato in platos]
    self.qty = qty
