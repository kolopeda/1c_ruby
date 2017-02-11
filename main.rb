require './ole_1c.rb'

server = OLE_1C.new
element = server.Invoke('Справочники', 'Номенклатура', ['НайтиПоКоду','РТ-00000057'])
#puts server.Ref_UUID(element) # UUID элемента
#puts server.MetadataName(element) # Наименование метаданных
#puts server.MetadataFullName(element) # Наименование методанных справочника
#puts element.Code # Код номенклатуры
#puts element.Description.encode('UTF-8') # Наименование номенклатуры
puts server.get_query.to_s
