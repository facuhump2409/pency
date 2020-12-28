require 'rubyXL'

#Spreadsheet.client_encoding = 'windows-1251:utf-8'
book = RubyXL::Parser.parse("26-12 Planilla de pedidos.xlsx")
precios = book['Precios y Menú']
costo_de_envio = 0
envio_gratis = 0
precios.each do |precio|
    precio.cells.each {|cell| p cell.value}
    costo_de_envio = precio[4] if precio[2] == "Costo de envío"
    envio_gratis = precio[4] if precio[2] == "Envio Gratis"
end
messages = book.worksheet["Mensajes WPP"]
items_dictionary = {} 
messages.each do |message|
    next if message[0] == 'Si' || message[0] == 'Procesado' || message[1].nil?
    order = message[1]
    total = order.match(/(?i)Total:\K.+/)[0]
    id_pedido = order.match(/(?i)Pedido:\K.+/)[0] #TODO evitar esta repeticion de logica con un diccionario y una lista de variables que tienen misma regex
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
    message.push "Si"
    p message[0]
    #Setear en Si el procesado
end
p items_dictionary
#def work_excel