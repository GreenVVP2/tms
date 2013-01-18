VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Module1 - Подпрограммы, для вызова

Public WithEvents myOlItems As Outlook.Items
Attribute myOlItems.VB_VarHelpID = -1

Public Sub Initialize_handler()
    Set myOlItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderToDo).Items
    Stop
End Sub

Private Sub Application_Startup()
    Initialize_handler
Stop
End Sub


Private Sub myOlItems_ItemChange(ByVal Task As Object)
Stop

Dim flagStatus
Dim taskStatus

' временно для подсказок
Dim Item As TaskItem
Set Item = Task

' статус флага
flagStatus = Item.UserProperties("Выполняется")
' статус задачи
taskStatus = tms.getStatus(Item)

' сравним статус задачи и значение флага, если не соответствуют - значит нажали на флаг
If flagStatus = True And taskStatus = olTaskInProgress Then
    Exit Sub
End If

If flagStatus = False And taskStatus <> olTaskInProgress Then
    Exit Sub
End If

If flagStatus = True And taskStatus <> olTaskInProgress Then
' запустили задачу
Call tms.startTask(Item)
Exit Sub
End If

If flagStatus = False And taskStatus = olTaskInProgress Then
' остановили задачу
Call tms.stopTask(Item)
Exit Sub
End If

End Sub

Public Sub testTest(obj As Object)

End Sub

Private Function getObjStatus(obj As Object)

    If obj.Class = olTask Then
        getObjStatus = obj.status
        Exit Function
    End If
    
    If obj.Class = olMail Then
        statusProperty = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003"
        Set PropertyAccessor = obj.PropertyAccessor
        getObjStatus = PropertyAccessor.GetProperty(statusProperty)
        Exit Function
    End If

End Function

