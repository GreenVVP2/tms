Attribute VB_Name = "tms"
Public Const idTaskInProgressProperty = "EntryIdForTaskInProgress"
Public Const globalPropertySchema = "http://schemas.microsoft.com/mapi/string/{FFF40745-9999-4C11-9E14-92701F001EB3}/"
Public Const idAppointmentInProgressProperty = "appointmentEntryID"
Public Const endWorkTime = "17:00:00"
Public Const endWorkTimeMax = "23:00:00"
Public Const idTaskInAppointment = "taskEntryID"


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
        getItemFromID = Nothing
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

' Приостанавливает работу задачи
Public Sub stopTask(obj As Object)
    
    Dim appointmentEntryID As String
    
    If obj Is Nothing Then
        MsgBox "Объект не передан"
        Exit Sub
    End If
    
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

    noteEntryID = "00000000AB47E344241E8B4FBF655870D2432CF907007114BF7C467D8D49880B8531D4C95A5A0000000002760000F40C8A0FF820264AAE328DE7A10CF74600004D1047060000"
    
    Set objNote = tms.getItemFromID(noteEntryID)
  
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

    noteEntryID = "00000000AB47E344241E8B4FBF655870D2432CF907007114BF7C467D8D49880B8531D4C95A5A0000000002760000F40C8A0FF820264AAE328DE7A10CF74600004D1047060000"
    
    Set objNote = tms.getItemFromID(noteEntryID)
  
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
     
     ' ищем свойство
     Set userProperty = obj.UserProperties.Find(propertyName)
    ' если нет то добавим
    If userProperty Is Nothing Then
        Set userProperty = obj.UserProperties.Add(propertyName, olText)
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
