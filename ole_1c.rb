require 'win32ole'

class OLE_1C

  attr_reader :connect

  def initialize
    @connect = WIN32OLE.new 'V83.Application'
    @connect.Connect("File=\"D:\InfoBase\"; Usr=\"Андрей\"; Pwd=\"8512481430\"")
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
    query.Text = 'ВЫБРАТЬ Продажи.Период, Продажи.Стоимость ИЗ РегистрНакопления.Продажи КАК Продажи'
    #nom=@connect.ТекущаяУниверсальнаяДата()
    query.SetParameter('Продажи.Период','11.02.2017')
    result=query.Execute.Unload

    sers=(0..result.Count).collect do |i|
      record=result.Get(i)
    #  record.Get(0)
      puts record
    end
    #return result
  end
end
