from openpyxl import *
import re

#Spreadsheet.client_encoding = 'windows-1251:utf-8'
book = load_workbook("26-12 Planilla de pedidos.xlsx")
prices = book['Precios y Menú']
costo_de_envio = 0
envio_gratis = 0
for row in prices.iter_rows(values_only= True):
    if row[2] == "Costo de envío": costo_de_envio = row[4]
    if row[2] == "Envio Gratis": envio_gratis = row[4]

messages = book["Mensajes WPP"]
items_dictionary = {} 
for row in prices.iter_rows(values_only= True):
    if (message[0] == 'Si' or message[0] == 'Procesado' or message[1] == None): continue
    order = message[1]
    total = re.search("(?<=Total\: ).+",order)[0]
    id_pedido = re.search("(?i)Pedido:\K.+",order)[0] #TODO evitar esta repeticion de logica con un diccionario y una lista de variables que tienen misma regex
    items_regex = "(?i)— .+"
    items = order.findall(items_regex)
    items.each do |item|
        qty_regex = /\[ \d \]/
        qty_product = item.match(qty_regex)
        qty = qty_product.nil? ? 1 : qty_product[0].gsub(/\[|\]/,"").strip
        product = item.match(/(?i)—( #{qty_regex})?\K.+(?=\>)/)[0].strip
        items_dictionary[product] = qty
    end
    delivery_date = re.search("(?i)entrega: \K\w+",order)[0]
    address = re.search("(?i)Direcci.n.+:\K.+",order)[0]
    payment_method = re.search(/(?i)Forma de pago: \K\w+/)
    customers_name = re.search(/(?i)Nombre y Apellido: \K\w+/)
    message[0] = "Si"
    weekly_orders << id_pedido
    #Setear en Si el procesado
end
book.write '26-12 Planilla de pedidos.xls'
p items_dictionary
#def work_excel
