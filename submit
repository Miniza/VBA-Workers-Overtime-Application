Sub Reset()

Dim iRow As Long

iRow = [Counta(Database!A:A)]

With frmForm
.cmbName.Value = ""
.cmbAssistant.Value = ""
.cmbVehicle.Value = ""
.cmbReason.Value = ""
.cmbSite.Value = ""
.cmbEstimated.Value = ""
.cmbDate.Value = ""
.cmbMonth.Value = ""
.cmbAssService.Value = ""
.cmbTrailer.Value = ""
.cmbNoLoads.Value = ""
.cmbAddManager.Value = ""
.cmbYear.Value = ""

 .txtComments.Visible = False
 .Label9.Visible = False
 .txtSpecify.Visible = False
 .Label16.Visible = False
 
.txtRowNumber.Value = ""
.txtComments.Value = ""
.txtReason.Value = ""

.lstDatabase.ColumnCount = 19
.lstDatabase.ColumnHeads = True

.lstDatabase.ColumnWidths = "60,90,90,90,90,90,90,120,90,90,90,90,90,90,90,90,90,90,90"

If iRow > 1 Then
.lstDatabase.RowSource = "Database!A2:S" & iRow
Else
.lstDatabase.RowSource = "Database!A2:S2"
End If
End With

End Sub

Sub submit()

Dim sh As Worksheet
Dim iRow As Long

Set sh = ThisWorkbook.Sheets("Database")
Set sb = ThisWorkbook.Sheets("Monthly")
If frmForm.txtRowNumber.Value = "" Then
iRow = [Counta(Database!A:A)] + 1
Else
iRow = frmForm.txtRowNumber.Value
End If


With sh
.Cells(iRow, 1) = iRow - 1
.Cells(iRow + 1, 1) = iRow - 1
.Cells(iRow, 2) = frmForm.cmbName.Value
.Cells(iRow + 1, 2) = frmForm.cmbAssistant.Value
.Cells(iRow, 3) = frmForm.cmbService.Value
.Cells(iRow, 4) = frmForm.cmbAssistant.Value
.Cells(iRow + 1, 4) = "I am The Assistant"
.Cells(iRow, 5) = frmForm.cmbAssService.Value
.Cells(iRow + 1, 5) = "I am The Assistant"
.Cells(iRow + 1, 3) = frmForm.cmbAssService.Value
.Cells(iRow, 6) = frmForm.cmbVehicle.Value
.Cells(iRow + 1, 6) = frmForm.cmbVehicle.Value
.Cells(iRow, 7) = frmForm.cmbTrailer.Value
.Cells(iRow + 1, 7) = frmForm.cmbTrailer.Value
.Cells(iRow, 8) = frmForm.cmbNoLoads.Value
.Cells(iRow + 1, 8) = frmForm.cmbNoLoads.Value
.Cells(iRow, 9) = frmForm.cmbReason.Value
.Cells(iRow + 1, 9) = frmForm.cmbReason.Value
.Cells(iRow, 10) = frmForm.txtSpecify.Value
.Cells(iRow + 1, 10) = frmForm.txtSpecify.Value
.Cells(iRow, 11) = frmForm.cmbSite.Value
.Cells(iRow + 1, 11) = frmForm.cmbSite.Value
.Cells(iRow, 12) = frmForm.cmbEstimated.Value
.Cells(iRow + 1, 12) = frmForm.cmbEstimated.Value
.Cells(iRow, 13) = frmForm.txtComments.Value
.Cells(iRow + 1, 13) = frmForm.txtComments.Value
.Cells(iRow, 14) = Application.UserName
.Cells(iRow + 1, 14) = Application.UserName
.Cells(iRow, 15) = frmForm.cmbDate.Value
.Cells(iRow + 1, 15) = frmForm.cmbDate.Value
.Cells(iRow, 16) = frmForm.cmbMonth.Value
.Cells(iRow + 1, 16) = frmForm.cmbMonth.Value
.Cells(iRow, 17) = frmForm.cmbYear.Value
.Cells(iRow + 1, 17) = frmForm.cmbYear.Value
.Cells(iRow, 18) = [Text(now(), "DD-MM-YY HH:MM:SS")]
.Cells(iRow + 1, 18) = [Text(now(), "DD-MM-YY HH:MM:SS")]
.Cells(iRow, 19) = frmForm.txtReason.Value
.Cells(iRow + 1, 19) = frmForm.txtReason.Value
End With

With sb
.Cells(iRow, 1) = iRow - 1
.Cells(iRow + 1, 1) = iRow - 1
.Cells(iRow, 2) = frmForm.cmbName.Value
.Cells(iRow + 1, 2) = frmForm.cmbAssistant.Value
.Cells(iRow, 3) = frmForm.cmbService.Value
.Cells(iRow, 4) = frmForm.cmbAssistant.Value
.Cells(iRow + 1, 4) = "I am The Assistant"
.Cells(iRow, 5) = frmForm.cmbAssService.Value
.Cells(iRow + 1, 5) = "I am The Assistant"
.Cells(iRow + 1, 3) = frmForm.cmbAssService.Value
.Cells(iRow, 6) = frmForm.cmbVehicle.Value
.Cells(iRow + 1, 6) = frmForm.cmbVehicle.Value
.Cells(iRow, 7) = frmForm.cmbTrailer.Value
.Cells(iRow + 1, 7) = frmForm.cmbTrailer.Value
.Cells(iRow, 8) = frmForm.cmbNoLoads.Value
.Cells(iRow + 1, 8) = frmForm.cmbNoLoads.Value
.Cells(iRow, 9) = frmForm.cmbReason.Value
.Cells(iRow + 1, 9) = frmForm.cmbReason.Value
.Cells(iRow, 10) = frmForm.txtSpecify.Value
.Cells(iRow + 1, 10) = frmForm.txtSpecify.Value
.Cells(iRow, 11) = frmForm.cmbSite.Value
.Cells(iRow + 1, 11) = frmForm.cmbSite.Value
.Cells(iRow, 12) = frmForm.cmbEstimated.Value
.Cells(iRow + 1, 12) = frmForm.cmbEstimated.Value
.Cells(iRow, 13) = frmForm.txtComments.Value
.Cells(iRow + 1, 13) = frmForm.txtComments.Value
.Cells(iRow, 14) = Application.UserName
.Cells(iRow + 1, 14) = Application.UserName
.Cells(iRow, 15) = frmForm.cmbDate.Value
.Cells(iRow + 1, 15) = frmForm.cmbDate.Value
.Cells(iRow, 16) = frmForm.cmbMonth.Value
.Cells(iRow + 1, 16) = frmForm.cmbMonth.Value
.Cells(iRow, 17) = frmForm.cmbYear.Value
.Cells(iRow + 1, 17) = frmForm.cmbYear.Value
.Cells(iRow, 18) = [Text(now(), "DD-MM-YY HH:MM:SS")]
.Cells(iRow + 1, 18) = [Text(now(), "DD-MM-YY HH:MM:SS")]
.Cells(iRow, 19) = frmForm.txtReason.Value
.Cells(iRow + 1, 19) = frmForm.txtReason.Value
End With

End Sub
Sub show_Form()

frmForm.Show

End Sub

Function Selected_List() As Long
Dim i As Long
Selected_List = 0
For i = 0 To frmForm.lstDatabase.ListCount - 1
If frmForm.lstDatabase.Selected(i) = True Then
Selected_List = i + 1
Exit For
End If


Next i

End Function
