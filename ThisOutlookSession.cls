VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public WithEvents myOlItems As Outlook.Items
Attribute myOlItems.VB_VarHelpID = -1
Public WithEvents mNavigationPane As NavigationPane
Attribute mNavigationPane.VB_VarHelpID = -1

Public Sub Initialize_handler()
    Set myOlItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderToDo).Items
    Set mNavigationPane = Application.ActiveExplorer.NavigationPane
End Sub

Private Sub Application_Startup()
    Initialize_handler
End Sub

Private Sub myOlItems_ItemChange(ByVal task As Object)
    Call tms.onItemChange(task)
End Sub

Private Sub mNavigationPane_ModuleSwitch(ByVal CurrentModule As NavigationModule)
    Call tms.onPanelModuleSwitch(CurrentModule)
End Sub





