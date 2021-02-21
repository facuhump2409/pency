class Coordinate:
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def increment_column(self):
        self.y = self.y + 1


def get_free_space_cell(cell, worksheet, order_id):
    # TODO ver tema de si ya tiene un numero que se frene en esa posicion
    current_value = cell.value
    coordinate = Coordinate(cell.row, cell.column)
    while current_value is not None and current_value != order_id:
        coordinate.increment_column()
        current_value = worksheet.cell(coordinate.x, coordinate.y).value
    return coordinate

def set_value_in_empty_space(excel_file_with_data, excel_file_with_formula, name, value, order_id):
    try:
        worksheet = excel_file_with_data['Pedido']
        worksheet_with_formulas = excel_file_with_formula['Pedido']
    except:
        worksheet = excel_file_with_data['Pedido Pency']
        worksheet_with_formulas = excel_file_with_formula['Pedido Pency']
    for row in worksheet.iter_rows():
        for cell in row:
            if (cell.value != None and isinstance(cell.value, str)):  # No sea Null y sea string
                if cell.value.upper() == "PEDIDO":
                    free_coordinate = get_free_space_cell(cell, worksheet, order_id)
                if cell.value.upper() == name.upper().strip():
                    coordinate = Coordinate(cell.row, free_coordinate.y)
                    # coordinate = get_free_space_cell(cell, worksheet)
                    # chr(ord('a') + 1) # para conseguir proximo caracter
                    # cell.row -> te devuelve la fila
                    # cell.column -> te devuelve la columna
                    # Si es None es que no hay nada, habria que conseguir la proxima letra del abecedario de esa columna que esta vacia
                    actual_value = worksheet_with_formulas.cell(coordinate.x, coordinate.y).value or 0
                    if actual_value != 0 and isinstance(actual_value,int):
                        worksheet_with_formulas.cell(coordinate.x, coordinate.y).value = int(value) + int(actual_value)
                    else:
                        worksheet_with_formulas.cell(coordinate.x, coordinate.y).value = value
                    return

def find_cell_with_value(value,worksheet):
    for row in worksheet.iter_rows(min_col=9):
        for cell in row:
            if (cell.value != None and isinstance(cell.value, str) and cell.value.upper() == value.upper().strip()):
                return cell

