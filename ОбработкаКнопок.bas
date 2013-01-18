Attribute VB_Name = "ОбработкаКнопок"
Sub СтартЗадачи()

    Dim CurrentItem As Object
    Dim curTaskStatus As OlTaskStatus
    
    If TypeName(Application.ActiveWindow) <> "Inspector" Then
        MsgBox ("Старт задачи вызван не в окне задачи")
        Exit Sub
    End If
    
    Set CurrentItem = Application.ActiveWindow.CurrentItem

    If Not tms.canBeTask(CurrentItem) Then
        MsgBox ("Этот объект не может быть задачей")
        Exit Sub
    End If
    
    ' проверяем статус задачи
    curTaskStatus = tms.getStatus(CurrentItem)
    
    If curTaskStatus = olTaskInProgress Then
        MsgBox ("Задача уже выполняется")
        Exit Sub
    End If
    
    ' Останавливаем выполняемую задачу
    Call tms.stopTaskInProgress
    
    ' начинаем работать над задачей
    Call tms.startTask(CurrentItem)
    
    MsgBox ("Задача выполняется")

End Sub

Sub СтопЗадачи()

    Dim CurrentItem As Object
    
    If TypeName(Application.ActiveWindow) <> "Inspector" Then
        MsgBox ("Стоп задачи вызван не в окне задачи")
        Exit Sub
    End If
    
    Set CurrentItem = Application.ActiveWindow.CurrentItem

    If Not tms.canBeTask(CurrentItem) Then
        MsgBox ("Этот объект не является задачей")
        Exit Sub
    End If
    
    ' Останавливаем задачу
    Call tms.stopTask(CurrentItem)
    
    ' Добавить сколько часов отработал и сколько всего сделал - миниотчет
    MsgBox ("Задача остановлена")
    
End Sub

' Создает встречу для задачи начиная с текущего времени до конца рабочего дня
' выполнение текущей задачи не приостанавливает
Sub СоздатьВстречу()

    Dim CurrentItem As Object
    Dim appointmentEntryID As String
    Dim objAppointment As AppointmentItem
    
    If TypeName(Application.ActiveWindow) <> "Inspector" Then
        MsgBox ("Создание встречи для задачи вызвано не в окне задачи")
        Exit Sub
    End If
    
    Set CurrentItem = Application.ActiveWindow.CurrentItem

    If Not tms.canBeTask(CurrentItem) Then
        MsgBox ("Этот объект не является задачей")
        Exit Sub
    End If
    
    ' Создаем встречу
    appointmentEntryID = tms.createAppointment(CurrentItem)
    
    ' Встреча создана - получим её
    Set objAppointment = tms.getItemFromID(appointmentEntryID)
    
    ' откроем встречу
    objAppointment.Display
    
End Sub

Sub УшелДомой()

End Sub



