require 'win32ole'

class OLE_1C

  attr_reader :connect

  def initialize
    @connect = WIN32OLE.new 'V83.Application'
    @connect.Connect("File=\"D:\InfoBase\"; Usr=\"AAAAA\"; Pwd=\"0000\"")
    @connect.Visible = false
  end

  # Преобразует объект в строку UTF-8
  def String(obj)
    @connect.String(obj).encode('UTF-8')
  end

  # Получить UUID ссылки
  def Ref_UUID(obj)
    String(obj.Ref.UUID)
  end

  # Возвращает имя метаданных ссылки
  def MetadataName(obj)
    String(obj.Metadata.Name)
  end

  # Возвращает полное имя метаданных ссылки
  def MetadataFullName(obj)
    String(obj.Metadata.FullName)
  end

  # Выполняет метод 1с предприятия
  # element = server.Invoke('Справочники', 'Контрагенты', ['НайтиПоКоду','00000002'])
  def Invoke(*obj)
    obj.inject(@connect) do |result, element|
      result.invoke *Array(element)
    end
  end

  # Получить ссылку по UUID
  # element = server.UUID_Ref('Справочники', 'Контрагенты', '6f579662-7453-11e3-b5e2-00269e72fb28')
  def UUID_Ref(type, name, uuid)
    Invoke(type, name).GetRef(@connect.NewObject('UUID',uuid))
  end

  def get_query
    query = @connect.NewObject('Query')
    query.Text = 'ВЫБРАТЬ Продажи.СуммаОстаток
    ИЗ РегистрНакопления.Продажи.Остатки(КОНЕЦПЕРИОДА(&ДатаОтчета, ДЕНЬ), ) КАК Продажи'
    cur_date = @connect.CurrentDate()
    query.SetParameter('ДатаОтчета', cur_date)
    result = query.Execute.Unload;
    record = result.Get(0)
    record.Get(0)
    return
  
    #return result
  end
end
