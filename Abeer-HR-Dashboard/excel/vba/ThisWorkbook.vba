Private Sub Workbook_Open()
    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").Activate
    ActiveWindow.DisplayGridlines = False
    Application.StatusBar = "Abeer HR Command Center ready - use the buttons at the top of the Dashboard."
End Sub
