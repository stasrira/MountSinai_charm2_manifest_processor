VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next 'in case the menu item has already been deleted
    Application.CommandBars(cWsMenuBar).Controls(cCustomMenuName).Delete 'delete the menu item
End Sub

Private Sub Workbook_Open()
    'create custom menus
    LoadCustomMenus
End Sub

