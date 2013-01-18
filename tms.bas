Attribute VB_Name = "tms"
Public Const idTaskInProgressProperty = "EntryIdForTaskInProgress"
Public Const globalPropertySchema = "http://schemas.microsoft.com/mapi/string/{FFF40745-9999-4C11-9E14-92701F001EB3}/"
Public Const idAppointmentInProgressProperty = "appointmentEntryID"
Public Const endWorkTime = "17:00:00"
Public Const endWorkTimeMax = "23:00:00"
Public Const idTaskInAppointment = "taskEntryID"
' Необходимо создать заметку, найти её EntryId и заменить ниже.
' Найти с помощью отладки EntryId Application.GetNamespace("MAPI").GetDefaultFolder(olFolderContacts).Items
Public Const noteEntryID = "00000000AB47E344241E8B4FBF655870D2432CF907007114BF7C467D8D49880B8531D4C95A5A0000000002760000F40C8A0FF820264AAE328DE7A10CF74600004D1047060000"


' Проверяет может ли быть объект задачей
Public Function canBeTask(obj As Object) As Boolean

    canBeTask = False

    If obj Is Nothing Then
        MsgBox "Объект не передан"
        Exit Function
    End If

    ' Задачами могут быть только объекты определенных классов
    If obj.Class = olMail Or _
       obj.Class = olTask Then
        
        canBeTask = True
        
    End If

End Function

' Устанавливает статус выполнения объекту
Public Sub setStatus(obj As Object, status As OlTaskStatus)

    Dim PropertyAccessor As PropertyAccessor
    Dim statusPropertySchemaName As String

    If obj Is Nothing Then
        MsgBox "Объект не передан"
        Exit Sub
    End If

    If obj.Class = olTask Then
        obj.status = status
        obj.Save
        Exit Sub
    End If
    
    ' Для других объектов, помеченных как задача
    ' Это свойство "Статус"
    statusPropertySchemaName = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003"
    Set PropertyAccessor = obj.PropertyAccessor
    Call PropertyAccessor.SetProperty(statusPropertySchemaName, status)
    obj.Save
    
End Sub

' Получает статус выполнения объекта
Public Function getStatus(obj As Object) As OlTaskStatus

    Dim PropertyAccessor As PropertyAccessor
    Dim statusPropertySchemaName As String

    getStatus = olTaskNotStarted

    If obj Is Nothing Then
        MsgBox "Объект не передан"
        Exit Function
    End If

    If obj.Class = olTask Then
        getStatus = obj.status
        Exit Function
    End If
    
    ' Для других объектов, помеченных как задача
    ' Это свойство "Статус"
    statusPropertySchemaName = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003"
    Set PropertyAccessor = obj.PropertyAccessor
    getStatus = PropertyAccessor.GetProperty(statusPropertySchemaName)

End Function

' Возвращает объект по идентификатору
Public Function getItemFromID(entryIdItem As String) As Object

    If entryIdItem = "" Then
        Exit Function
    End If

    On Error GoTo errorItemNotFound
    Set getItemFromID = Application.GetNamespace("MAPI").getItemFromID(entryIdItem)

    Exit Function

errorItemNotFound:

        Exit Function

End Function


' Ищет задачу в статусе "Выполняется" и приостанавливает её
Public Sub stopTaskInProgress()
    
    Dim entryIdForTaskInProgress As String
    Dim obj As Object
        
    entryIdForTaskInProgress = tms.getEntryIdForTaskInProgress()
    
    Set obj = tms.getItemFromID(entryIdForTaskInProgress)
    
    If Not obj Is Nothing Then
        objStatus = tms.getStatus(obj)
            
        If objStatus = olTaskInProgress Then
            Call tms.stopTask(obj)
        End If
    End If
    
End Sub

' Завершает задачу
Public Sub completeTask(obj As Object)

    Dim entryIdForTaskInProgress As String
    Dim appointmentEntryID As String
    
    If obj Is Nothing Then
        MsgBox "Объект не передан"
        Exit Sub
    End If
    
    ' Убираем галку Выполняется
    Call tms.setUserProperty(obj, "Выполняется", False)
    
    ' проверим текущая выполняемая задача это или нет
    ' получим идентификатор выполняемой задачи
    entryIdForTaskInProgress = tms.getEntryIdForTaskInProgress()
    
    If entryIdForTaskInProgress = obj.entryId Then
        ' если текущая остановим её
        ' очистим идентификатор выполняемой задачи
        Call tms.saveEntryIdForTaskInProgress("")
        
        ' получаем идентификатор встречи
        appointmentEntryID = tms.getUserProperty(obj, tms.idAppointmentInProgressProperty)
        
        ' делаем время окончания встречи на сейчас
        Call tms.stopAppointmentByEntryId(appointmentEntryID)
        
    End If

End Sub
' Приостанавливает работу задачи
Public Sub stopTask(obj As Object)
    
    Dim appointmentEntryID As String
    
    If obj Is Nothing Then
        MsgBox "Объект не передан"
        Exit Sub
    End If
    
    ' Убираем галку Выполняется
    Call tms.setUserProperty(obj, "Выполняется", False)
    
    ' Меняем статус на Отложена
    Call tms.setStatus(obj, olTaskDeferred)
    
    ' очистим идентификатор выполняемой задачи
    Call tms.saveEntryIdForTaskInProgress("")
    
    ' получаем идентификатор встречи
    appointmentEntryID = tms.getUserProperty(obj, tms.idAppointmentInProgressProperty)
    
    ' делаем время окончания встречи на сейчас
    Call tms.stopAppointmentByEntryId(appointmentEntryID)
    
    ' чистим в задаче идентификатор встречи, т.к. работа над задачей завершена
    ' убрал пока
    ' Call tms.setUserProperty(obj, tms.idAppointmentInProgressProperty, "")
    
End Sub

' сохраняет EntryId выполняемой задачи в глобальном хранилище
Public Sub saveEntryIdForTaskInProgress(entryId As String)
    
    Call tms.saveGlobalProperty(tms.idTaskInProgressProperty, entryId)
    
End Sub

' получает EntryId выполняемой задачи из глобального хранилища
Public Function getEntryIdForTaskInProgress() As String

    getEntryIdForTaskInProgress = tms.getGlobalProperty(tms.idTaskInProgressProperty)
    
End Function

' Сохраняет глобальную переменную (в свойствах заметки временно)
Public Sub saveGlobalProperty(propName As String, propValue As String)
    
    Dim noteEntryID As String
    Dim propNameFull As String
    Dim objNote As NoteItem

    Set objNote = tms.getItemFromID(tms.noteEntryID)
  
    If objNote Is Nothing Then
        MsgBox "Заметка для глобальных переменных не создана"
        Exit Sub
    End If
 
    ' полное имя переменной
    propNameFull = tms.globalPropertySchema & propName

    Call objNote.PropertyAccessor.SetProperty(propNameFull, propValue)
 
    objNote.Save
    
End Sub

' метод реализует получение глобальных переменных из свойств заметки
Public Function getGlobalProperty(propName As String)

    Dim noteEntryID As String
    Dim propNameFull As String
    Dim objNote As NoteItem

    Set objNote = tms.getItemFromID(tms.noteEntryID)
  
    If objNote Is Nothing Then
        MsgBox "Заметка для глобальных переменных не создана"
        Exit Function
    End If
 
    ' полное имя переменной
    propNameFull = tms.globalPropertySchema & propName

    On Error GoTo ErrPropertyNotExists
    
    getGlobalProperty = objNote.PropertyAccessor.GetProperty(propNameFull)
    
    Exit Function
    
ErrPropertyNotExists:
    ' создаем свойство c пустым значением
    Call tms.saveGlobalProperty(propName, "")
    
End Function


' Получает пользовательское поле из объекта
Public Function getUserProperty(obj As Object, propertyName As String) As String
    
    Dim userProperty As userProperty
    
    getUserProperty = ""
    
    ' ищем свойство
    Set userProperty = obj.UserProperties.Find(propertyName)
    
    If userProperty Is Nothing Then
        Exit Function
    End If
    
    getUserProperty = userProperty.Value

End Function

' Сохраняет пользовательское поле в объект
Public Sub setUserProperty(obj As Object, propertyName As String, propertyValue As String)
     
    Dim userProperty As userProperty
    Dim propertyType As UserDefinedProperty
 
    ' ищем свойство
     Set userProperty = obj.UserProperties.Find(propertyName)
    ' если нет то добавим
    If userProperty Is Nothing Then
    
        Set propertyType = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderToDo).UserDefinedProperties.Find(propertyName)
     
        If propertyType Is Nothing Then
           MsgBox "Свойство не определено на уровне папки"
           Exit Sub
        End If
    
        Set userProperty = obj.UserProperties.Add(propertyName, propertyType.Type)
        
    End If

    userProperty.Value = propertyValue

    obj.Save
    
End Sub

' Завершает встречу текущим временем
Public Sub stopAppointmentByEntryId(entryId As String)
    
    Dim objAppointment As AppointmentItem
    Dim curDate As Date
    Dim endTime

    Set objAppointment = tms.getItemFromID(entryId)
    If objAppointment Is Nothing Then
        Exit Sub
    End If
    
    curDate = Date
    endTime = Time
    
    endDateStr = curDate & " " & endTime

    objAppointment.End = endDateStr
    objAppointment.Save
    
End Sub


' создание встречи из задачи или письма
Public Function createAppointment(obj As Object) As String

    Dim objAppointment As AppointmentItem
    Dim curDate As Date
    Dim curTime
    Dim endTime

    curDate = Date
    curTime = Time
    endTime = tms.endWorkTime
    
    If curTime > endTime Then
        endTime = tms.endWorkTimeMax
    End If
    
    curDateStr = curDate & " " & curTime
    endDateStr = curDate & " " & endTime

    Set objAppointment = Application.CreateItem(olAppointmentItem)
   
    With objAppointment
        .Subject = obj.Subject
        .Body = obj.Body
        .Location = ""
        .AllDayEvent = False
        .Categories = obj.Categories
        .Start = curDateStr
        .End = endDateStr
        .ReminderSet = False
        .ReminderMinutesBeforeStart = 0
        .ReminderPlaySound = False
        .Importance = obj.Importance
        .BusyStatus = olBusy
        .Sensitivity = olPrivate
    End With
    
    ' сохраним идентификатор задачи во встрече
    Call tms.setUserProperty(objAppointment, tms.idTaskInAppointment, obj.entryId)

    objAppointment.Save
    
    createAppointment = objAppointment.entryId
    
    Set objAppointment = Nothing

End Function

' Начинает выполнение задачи
Public Sub startTask(obj As Object)

    Dim appointmentEntryID As String

    ' Убираем галку Выполняется
    Call tms.setUserProperty(obj, "Выполняется", True)

    ' ставим статус в работе
    Call tms.setStatus(obj, olTaskInProgress)
    
    ' сохраним идентификатор выполняемой задачи для дальнейшего поиска
    Call tms.saveEntryIdForTaskInProgress(obj.entryId)
    
    ' создаем встречу с текущего времени до конца рабочего дня
    appointmentEntryID = tms.createAppointment(obj)
    
    ' записываем в задачу идентификатор встречи
    Call tms.setUserProperty(obj, tms.idAppointmentInProgressProperty, appointmentEntryID)
    
End Sub

' Создает заметку для хранения глобальных данных
Sub createNoteForGlobalProperties(ForNothing)

    Dim noteEntryID As String
    Dim noteText As String
    Dim objNote As NoteItem
    Dim noteFolder As Outlook.MAPIFolder
    
    noteEntryID = "00000000AB57E344241E8B4FBF655870D2432CF907007114BF7C467D8D49880B8531D4C95A5A0000000002760000F40C8A0FF820264AAE328DE7A10CF74600004D1047060000"
    noteText = "7777777Не удалять!!! Здесь хранится идентификатор текущей выполняемой задачи!!!"
    
    Set noteFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderNotes)
    Set objNote = noteFolder.Items.Add(olNoteItem)
    
    objNote.entryId = noteEntryID
    objNote.Body = noteText
    
    objNote.Save
    
End Sub

' Ищет незавершенные задачи
Public Function getNotCompletedItems() As Items

    Dim filterNotCompletedTasks As String
    ' Ищем задачи у которых не стоит флаг Завершена и Пустая дата завершения
    filterNotCompletedTasks = "@SQL=""http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/810f0040"" Is Null And ""http://schemas.microsoft.com/mapi/proptag/0x10910040"" Is Null"
    Set getNotCompletedItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderToDo).Items.Restrict(filterNotCompletedTasks)

End Function


Public Sub onItemChange(Item As Object)

    Dim flagStatus
    Dim taskStatus
    Dim taskComplete
    'Dim propertyChema As String
    
    'propertyChema = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/810f0040"
    
    ' статус флага
    flagStatus = tms.getUserProperty(Item, "Выполняется")
    ' статус задачи
    taskStatus = tms.getStatus(Item)
    ' Дата выполнения
    ' taskComplete = Item.PropertyAccessor.GetProperty(propertyChema)
    
    ' сравним статус задачи и значение флага, если не соответствуют - значит нажали на флаг
    ' если менять статус вручную то будет некорректно работать - точнее он не даст поменять, а может будет - смотреть
    
    ' может быть завершили задачу или чтото поменяли в завершенной задаче
    If taskStatus = olTaskComplete Then
        ' Завершили задачу
        Call tms.completeTask(Item)
        
        MsgBox ("Ура! Задача выполнена! :)")
        Exit Sub
        
    End If
    
    ' задача выполняется - ничего не делаем
    If flagStatus = True And taskStatus = olTaskInProgress Then
        Exit Sub
    End If
    
    ' задача не выполняется - ничего не делаем
    If flagStatus = False And taskStatus <> olTaskInProgress Then
        Exit Sub
    End If
    
    ' или запустили путем нажатия на галку или вручную поменяли статус
    ' галка важнее - запускаем задачу
    If flagStatus = True And taskStatus <> olTaskInProgress Then
        ' Останавливаем выполняемую задачу
        Call tms.stopTaskInProgress
        ' запустили задачу
        Call tms.startTask(Item)
        MsgBox ("Задача выполняется")
        Exit Sub
    End If
    
    ' или остановили путем нажатия на галку или вручную поменяли статус
    ' галка важнее - останавливаем задачу
    If flagStatus = False And taskStatus = olTaskInProgress Then
        ' остановили задачу
        Call tms.stopTask(Item)
        MsgBox ("Задача остановлена")
        Exit Sub
    End If

End Sub

Public Sub onPanelModuleSwitch(module As NavigationModule)

    Dim objPane As NavigationPane
    Set objPane = Application.ActiveExplorer.NavigationPane
    ' текущий модуль отправляем в конец
    objPane.CurrentModule.Position = objPane.Modules.Count
    
    If module.NavigationModuleType = olModuleMail Then
        objPane.IsCollapsed = False
    Else
        objPane.IsCollapsed = True
    End If
    
    'For Each objModule In objPane.Modules
    '    objModule.Visible = True
    'Next
     
End Sub
