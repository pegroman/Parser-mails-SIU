#gem install roo-xls
require 'roo-xls'

archivo = Roo::Spreadsheet.open('./file.xls')

contactos = archivo.sheet(0).column(6)

salida = File.new("mails.txt", "w+")

i = 1
while contactos[i] != nil do
    salida.write (contactos[i].split(" ")[2] + ",\n")
    i = i + 1
end

salida.close()