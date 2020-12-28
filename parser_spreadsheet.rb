require 'spreadsheet'

#Spreadsheet.client_encoding = 'windows-1251:utf-8'
weekly_orders = []
book = Spreadsheet.open "26-12 Planilla de pedidos.xls"
precios = book.worksheet 'Precios y Menu'
precios.each do |precio|
    costo_de_envio = precio[4] if precio[2] == "Costo de envío"
    envio_gratis = precio[4] if precio[2] == "Envio Gratis"
end
messages = book.worksheet "Mensajes WPP"
items_dictionary = {} 
messages.each do |message|
    next if message[0] == 'Si' || message[0] == 'Procesado' || message[1].nil?
    order = message[1]
    total = order.match(/(?i)Total:\K.+/)[0]
    id_pedido = order.match(/(?i)Pedido:\K.+/)[0] #TODO evitar esta repeticion de logica con un diccionario y una lista de variables que tienen misma regex
    next if weekly_orders.include? id_pedido #TODO ver como manejamos para que no procese dos veces el mismo pedido -> Posible solucion: escribir los ids en un archivo
    items_regex = /(?i)— .+/
    items = order.scan(items_regex)
    items.each do |item|
        qty_regex = /\[ \d \]/
        qty_product = item.match(qty_regex)
        qty = qty_product.nil? ? 1 : qty_product[0].gsub(/\[|\]/,"").strip
        product = item.match(/(?i)—( #{qty_regex})?\K.+(?=\>)/)[0].strip
        items_dictionary[product] = qty
    end
    delivery_date = order.match(/(?i)entrega: \K\w+/)[0]
    address = order.match(/(?i)Direcci.n.+:\K.+/)[0]
    payment_method = order.match(/(?i)Forma de pago: \K\w+/)
    customers_name = order.match(/(?i)Nombre y Apellido: \K\w+/)
    message[0] = "Si"
    weekly_orders << id_pedido
    #Setear en Si el procesado
end
book.write '26-12 Planilla de pedidos.xls'
p items_dictionary
#def work_excel