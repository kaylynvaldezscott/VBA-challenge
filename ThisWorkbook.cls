VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    If MsgBox("Press OK to begin process, or CANCEL to exit.", vbQuestion + vbOKCancel, "Kaylyn Valdez-Scott - Solution 2") = vbOK Then
        Call GatherData
        MsgBox "Calculations complete", vbInformation, "Kaylyn Valdez-Scott - Solution 2"
    End If
    

End Sub

