Const PLACEHOLDER As String = "Enter Main Duties that will be performed"

Private Sub cckCopy_Click()
If cckCopy.Value = True Then
Label19.Visible = True
cmbAddManager.Visible = True
End If
End Sub

Private Sub cmbAssistant_Change()
Dim index As Integer
index = cmbAssistant.ListIndex

cmbAssService.Clear
Select Case index
Case Is = 0
With cmbAssService
.AddItem "2154234"
End With
Case Is = 1
With cmbAssService
.AddItem "9652302"
End With
Case Is = 2
With cmbAssService
.AddItem "6664806"
End With
Case Is = 3
With cmbAssService
.AddItem "9652205"
End With
Case Is = 4
With cmbAssService
.AddItem "9652803"
End With
Case Is = 5
With cmbAssService
.AddItem "9652701"
End With
Case Is = 6
With cmbAssService
.AddItem "9652507"
End With
Case Is = 7
With cmbAssService
.AddItem "2150580"
End With
Case Is = 8
With cmbAssService
.AddItem "6738001"
End With
Case Is = 9
With cmbAssService
.AddItem "9797904"
End With
Case Is = 10
With cmbAssService
.AddItem "6912206"
End With
Case Is = 11
With cmbAssService
.AddItem "7551703"
End With
Case Is = 12
With cmbAssService
.AddItem "9798706"
End With
Case Is = 13
With cmbAssService
.AddItem "9799303"
End With
Case Is = 14
With cmbAssService
.AddItem "9114904"
End With
Case Is = 15
With cmbAssService
.AddItem "9798200"
End With
Case Is = 16
With cmbAssService
.AddItem "7552203"
End With
Case Is = 17
With cmbAssService
.AddItem "9798005"
End With
Case Is = 18
With cmbAssService
.AddItem "9798609"
End With
Case Is = 19
With cmbAssService
.AddItem "7906102"
End With

Case Is = 20
With cmbAssService
.AddItem "2162836"
End With

Case Is = 21
With cmbAssService
.AddItem "9059406"
End With

Case Is = 22
With cmbAssService
.AddItem "9114807"
End With

Case Is = 23
With cmbAssService
.AddItem "2162832"
End With

Case Is = 24
With cmbAssService
.AddItem "9553304"
End With

Case Is = 25
With cmbAssService
.AddItem "9553401"
End With

Case Is = 26
With cmbAssService
.AddItem "9554009"
End With

Case Is = 27
With cmbAssService
.AddItem "9554300"
End With

Case Is = 28
With cmbAssService
.AddItem "9553207"
End With

Case Is = 29
With cmbAssService
.AddItem "9552901"
End With

Case Is = 30
With cmbAssService
.AddItem "9552707"
End With

Case Is = 31
With cmbAssService
.AddItem "9552600"
End With

Case Is = 32
With cmbAssService
.AddItem "9552405"
End With

Case Is = 33
With cmbAssService
.AddItem "9552308"
End With

Case Is = 34
With cmbAssService
.AddItem "9552103"
End With

Case Is = 35
With cmbAssService
.AddItem "8882401"
End With

Case Is = 36
With cmbAssService
.AddItem "5643902"
End With

Case Is = 37
With cmbAssService
.AddItem "9928405"
End With

Case Is = 38
With cmbAssService
.AddItem "2162414"
End With

Case Is = 39
With cmbAssService
.AddItem "9930108"
End With

Case Is = 40
With cmbAssService
.AddItem "9927905"
End With

Case Is = 41
With cmbAssService
.AddItem "9929703"
End With

Case Is = 42
With cmbAssService
.AddItem "9928600"
End With

Case Is = 43
With cmbAssService
.AddItem "2162477"
End With

Case Is = 44
With cmbAssService
.AddItem "2154329"
End With

Case Is = 45
With cmbAssService
.AddItem "2162488"
End With

Case Is = 46
With cmbAssService
.AddItem "2162466"
End With

Case Is = 47
With cmbAssService
.AddItem "9295500"
End With

Case Is = 48
With cmbAssService
.AddItem "2162589"
End With

Case Is = 49
With cmbAssService
.AddItem "2162470"
End With

Case Is = 50
With cmbAssService
.AddItem "2162462"
End With

Case Is = 51
With cmbAssService
.AddItem "2162593"
End With

Case Is = 52
With cmbAssService
.AddItem "2169263"
End With

Case Is = 53
With cmbAssService
.AddItem "2169262"
End With
End Select
End Sub



Private Sub cmbEstimated_Change()
cmbEstimated.Value = Format(cmbEstimated, "hh:mm")
End Sub

Private Sub cmbName_Change()
Dim index As Integer
index = cmbName.ListIndex

cmbService.Clear
Select Case index
Case Is = 0
With cmbService
.AddItem "9928405"
End With
Case Is = 1
With cmbService
.AddItem "9553207"
End With
Case Is = 2
With cmbService
.AddItem "9114904"
End With
Case Is = 3
With cmbService
.AddItem "9930108"
End With
Case Is = 4
With cmbService
.AddItem "2162477"
End With
Case Is = 5
With cmbService
.AddItem "2162832"
End With
Case Is = 6
With cmbService
.AddItem "6912206"
End With
Case Is = 7
With cmbService
.AddItem "2162836"
End With
Case Is = 8
With cmbService
.AddItem "2162462"
End With
Case Is = 9
With cmbService
.AddItem "9553401"
End With
Case Is = 10
With cmbService
.AddItem "2162466"
End With
Case Is = 11
With cmbService
.AddItem "2169263"
End With
Case Is = 12
With cmbService
.AddItem "9553304"
End With
Case Is = 13
With cmbService
.AddItem "8882401"
End With
Case Is = 14
With cmbService
.AddItem "9798200"
End With
Case Is = 15
With cmbService
.AddItem "9114807"
End With
Case Is = 16
With cmbService
.AddItem "9797904"
End With
Case Is = 17
With cmbService
.AddItem "9798200"
End With
Case Is = 18
With cmbService
.AddItem "2162488"
End With

Case Is = 19
With cmbService
.AddItem "7552203"
End With

Case Is = 20
With cmbService
.AddItem "9059407"
End With

Case Is = 21
With cmbService
.AddItem "2162414"
End With

Case Is = 22
With cmbService
.AddItem "9552600"
End With

Case Is = 23
With cmbService
.AddItem "9927905"
End With

Case Is = 24
With cmbService
.AddItem "9554009"
End With

Case Is = 25
With cmbService
.AddItem "9652507"
End With

Case Is = 26
With cmbService
.AddItem "5643902"
End With

Case Is = 27
With cmbService
.AddItem "2162589"
End With

Case Is = 28
With cmbService
.AddItem "9799303"
End With

Case Is = 29
With cmbService
.AddItem "9552707"
End With

Case Is = 30
With cmbService
.AddItem "9552308"
End With

Case Is = 31
With cmbService
.AddItem "7551703"
End With

Case Is = 32
With cmbService
.AddItem "9652803"
End With

Case Is = 33
With cmbService
.AddItem "6738001"
End With

Case Is = 34
With cmbService
.AddItem "2167200"
End With

Case Is = 35
With cmbService
.AddItem "9798005"
End With

Case Is = 36
With cmbService
.AddItem "9652205"
End With

Case Is = 37
With cmbService
.AddItem "2154329"
End With

Case Is = 38
With cmbService
.AddItem "9798706"
End With

Case Is = 39
With cmbService
.AddItem "9552103"
End With

Case Is = 40
With cmbService
.AddItem "9929703"
End With

Case Is = 41
With cmbService
.AddItem "9552901"
End With

Case Is = 42
With cmbService
.AddItem "2150580"
End With

Case Is = 43
With cmbService
.AddItem "9928600"
End With

Case Is = 44
With cmbService
.AddItem "2169262"
End With

Case Is = 45
With cmbService
.AddItem "9552405"
End With

Case Is = 46
With cmbService
.AddItem "2162470"
End With

Case Is = 47
With cmbService
.AddItem "9554300"
End With

Case Is = 48
With cmbService
.AddItem "9928804"
End With

Case Is = 49
With cmbService
.AddItem "2162593"
End With

Case Is = 50
With cmbService
.AddItem "3017605"
End With

Case Is = 51
With cmbService
.AddItem "2154234"
End With

Case Is = 52
With cmbService
.AddItem "9652302"
End With

Case Is = 53
With cmbService
.AddItem "2162832"
End With

Case Is = 54
With cmbService
.AddItem "9798609"
End With

Case Is = 55
With cmbService
.AddItem "7906102"
End With

Case Is = 56
With cmbService
.AddItem "9652701"
End With
End Select
End Sub


Private Sub cmbReason_Change()
Dim index As Integer
index = cmbReason.ListIndex

'txtSpecify.Clear
Select Case index
Case Is = 4
With txtSpecify
Label16.Visible = True
txtSpecify.Visible = True
txtSpecify.SetFocus
End With
End Select
End Sub

Private Sub cmdEdit_Click()
If Selected_List = 0 Then
MsgBox "No Row is selected.", vbOKOnly + vbInformation, "Edit"
Exit Sub
End If

Me.txtRowNumber.Value = Selected_List + 1
Me.cmbName.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 1)
Me.cmbService.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 2)
Me.cmbAssistant.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 3)
Me.cmbAssService.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 4)
Me.cmbVehicle.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 5)
Me.cmbTrailer.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 6)
Me.cmbNoLoads.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 7)
Me.cmbReason.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 8)
Me.cmbSite.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 10)
Me.cmbEstimated.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 11)
Me.cmbDate.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 11)
Me.cmbMonth.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 12)
Me.txtComments.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 9)
Me.txtSpecify.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 9)
Me.txtReason.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 18)
Me.cmbDate.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 14)
Me.cmbMonth.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 15)
Me.cmbYear.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 16)
If txtComments.Value = "" Then
Label9.Visible = False
txtComments.Visible = False
Else
Label9.Visible = True
txtComments.Visible = True
End If

If txtSpecify.Value = "" Then
Label16.Visible = False
txtSpecify.Visible = False
Else
Label16.Visible = True
txtSpecify.Visible = True
End If

MsgBox "Please Make the required changes and click on 'Add' to re add to database.", vbOKOnly + vbInformation, "Edit"
End Sub

Private Sub cmdReset_Click()
Dim msgValue As VbMsgBoxResult
msgValue = MsgBox("Do you want to Reset form?", vbYesNo + vbInformation, "Confirmation")
If msgValue = vbNo Then Exit Sub
Call Reset
End Sub

Private Sub cmdSave_Click()

Dim i As Integer

If cmbName.ListIndex = -1 Then
Cancel = 1
MsgBox ("Employee Name field is required")
cmbName.SetFocus
Exit Sub
End If

If cmbService.ListIndex = -1 Then
Cancel = 1
MsgBox ("Service Number field is required")
cmbService.SetFocus
Exit Sub
End If

If cmbAssistant.ListIndex = -1 Then
Cancel = 1
MsgBox ("Assistant field is required")
cmbAssistant.SetFocus
Exit Sub
End If

If cmbVehicle.ListIndex = -1 Then
Cancel = 1
MsgBox ("Vehicle field is required")
cmbVehicle.SetFocus
Exit Sub
End If

If cmbDate.ListIndex = -1 Then
Cancel = 1
MsgBox ("Date field is required")
cmbDate.SetFocus
Exit Sub
End If

If cmbMonth.ListIndex = -1 Then
Cancel = 1
MsgBox ("Month of overtime required")
cmbVehicle.SetFocus
Exit Sub
End If

If cmbReason.ListIndex = -1 Then
Cancel = 1
MsgBox ("Reason for overtime field is required")
cmbReason.SetFocus
Exit Sub
End If


If cmbSite.ListIndex = -1 Then
Cancel = 1
MsgBox ("Site to service field is required")
cmbSite.SetFocus
Exit Sub
End If

If cmbEstimated.ListIndex = -1 Then
Cancel = 1
MsgBox ("Estimated time field is required")
cmbEstimated.SetFocus
Exit Sub
End If

If txtReason.Value = "" Then
MsgBox ("The comment for overtime is required")
txtReason.SetFocus
Exit Sub
End If

If txtComments.Value = "" Then
If cmbDate.Value <> [Text(now(), "DD")] Then
MsgBox ("Please provide a reason why the date applied for is different from todays date")
txtComments.Visible = True
Label9.Visible = True
txtComments.SetFocus
Exit Sub
End If
End If

Dim msgValue As VbMsgBoxResult
msgValue = MsgBox("Please re check Data and press yes to save?", vbYesNo + vbInformation, "Confirmation")

If msgValue = vbNo Then Exit Sub
Call submit
Call Reset
End Sub


Private Sub Label6_Click()

End Sub

Private Sub managersub_Click()

Dim OutlookApp As Object
Dim OutlookMail As Object
Set OutlookApp = CreateObject("Outlook.Application")
Set OutlookMail = OutlookApp.CreateItem(0)
On Error Resume Next
With OutlookMail
.To = "ryan.papanicolaou@durban.gov.za"
.HTMLBody = "Hi Manager" & vbNewLine & _
"Please find the attached Transfer stations overtime requests. Please review and make a decision (Approve or reject) on the decision column of the excel database attached and revert back to the applicant(s)"
If cckCopy.Value = True Then
If cmbAddManager.Value = "Loganathan Moodley" Then
.CC = "muhammed.osman@durban.gov.za"
.HTMLBody = "Hi Logan Moodley" & vbNewLine & _
"With Ryan being on leave, please find the attached Transfer stations overtime requests. Please review and make a decision (Approve or reject) on the decision column of the excel database attached and revert back to the applicant(s)"
ElseIf cmbAddManager.Value = "Rhandihr Sivapersad" Then
.CC = "minenhledlamini37@gmail.com"
.HTMLBody = "Hi Rhandihr Sivapersard" & vbNewLine & _
"With Ryan being on leave, please find the attached Transfer stations overtime requests. Please review and make a decision (Approve or reject) on the decision column of the excel database attached and revert back to the applicant(s)"
ElseIf cmbAddManager.Value = "Trevan Govender" Then
.CC = "sivendren.govender@durban.gov.za"
.HTMLBody = "Hi Trevan Govender" & vbNewLine & _
"With Ryan being on leave, please find the attached Transfer stations overtime requests. Please review and make a decision (Approve or reject) on the decision column of the excel database attached and revert back to the applicant(s)"
End If

End If
.BCC = ""
.Subject = "Overtime Request"
 ThisWorkbook.Sheets("Home").Visible = False
 ActiveWorkbook.Save
.Attachments.Add Application.ActiveWorkbook.FullName
'.HTMLBody = "Hi Manager" & vbNewLine & _
"please find the attached Transfer stations overtime requests please review and make a decision (Approve or reject) on the decision column of the excel database attached and revert back to the applicant(s)"
.Send
ThisWorkbook.Sheets("Home").Visible = True
ActiveWorkbook.Save
End With
Set OutlookMail = Nothing
Set OutlookApp = Nothing

MsgBox ("Request Sent")
    Sheets("Database").Select
    Range("A2:S1612").Select
    Selection.EntireRow.Delete
    Sheets("Home").Select
ActiveWorkbook.Save
End Sub


Private Sub txtReason_Change()
 If txtReason.Value = "" Then
        txtReason.Value = PLACEHOLDER
        txtReason.ForeColor = &H8000000C
        txtReason.SelStart = 0
    ElseIf StrComp(txtReason.Value, PLACEHOLDER, vbTextCompare) Then
        Application.EnableEvents = False
        txtReason.Value = Replace(txtReason.Value, PLACEHOLDER, "", , , vbTextCompare)
        txtReason.ForeColor = &H80000008
        Application.EnableEvents = True
    End If
End Sub

Private Sub UserForm_Initialize()
txtReason.SelStart = 0
txtComments.SelStart = 0

Dim lRow As Long, lLar As Long, sTest As String
With Worksheets("Data")
 lLar = .Cells(Rows.Count, "A").End(xlUp).Row
 If lLar > 2 Then
 For lRow = 2 To .Cells(Rows.Count, "A").End(xlUp).Row
 sTest = .Cells(lRow, "A")
 Me.cmbName.AddItem sTest
 Next lRow
 End If
 
 lLar = .Cells(Rows.Count, "A").End(xlUp).Row
 If lLar > 2 Then
 For lRow = 2 To .Cells(Rows.Count, "A").End(xlUp).Row
 sTest = .Cells(lRow, "A")
 Me.cmbAssistant.AddItem sTest
 Next lRow
 End If
 End With



With cmbReason
.AddItem "Overflow of refuse"
.AddItem "Truck Breakdown"
.AddItem "Plant Electrical failure"
.AddItem "Plant Mechanical failure"
.AddItem "Other"
End With

With cmbAddManager
.AddItem "Loganathan Moodley"
.AddItem "Rhandihr Sivapersad"
.AddItem "Trevan Govender"
End With

With cmbSite
.AddItem "Clermont TS"
.AddItem "MT.Edgecombe TS"
.AddItem "Electron Road TS"
.AddItem "Flower Road TS"
.AddItem "Umlazi TS"
.AddItem "Hammarsdale"
.AddItem "Amanzimtoti TS"
.AddItem "Chatsworth TS"
End With

With cmbVehicle
.AddItem "Not Applicable"
.AddItem "NDM 11949"
.AddItem "NDM 11951"
.AddItem "NDM 11953"
.AddItem "NDM 11954"
.AddItem "NDM 11955"
.AddItem "NDM 11958"
.AddItem "NDM 11959"
.AddItem "NDM 11736"
.AddItem "NDM 11737"
.AddItem "NDM 17825"
.AddItem "NDM 17826"
.AddItem "NDM 17828"
.AddItem "NDM 11804"
.AddItem "NDM 4504"
.AddItem "NDM 4118"
.AddItem "NDM 4119"
.AddItem "NDM 4214"
End With

With cmbDate
.AddItem "01"
.AddItem "02"
.AddItem "03"
.AddItem "04"
.AddItem "05"
.AddItem "06"
.AddItem "07"
.AddItem "08"
.AddItem "09"
.AddItem "10"
.AddItem "11"
.AddItem "12"
.AddItem "13"
.AddItem "14"
.AddItem "15"
.AddItem "16"
.AddItem "17"
.AddItem "18"
.AddItem "19"
.AddItem "20"
.AddItem "21"
.AddItem "22"
.AddItem "23"
.AddItem "24"
.AddItem "25"
.AddItem "26"
.AddItem "27"
.AddItem "28"
.AddItem "29"
.AddItem "30"
.AddItem "31"
End With

With cmbMonth
.AddItem "Jan"
.AddItem "Feb"
.AddItem "Mar"
.AddItem "Apr"
.AddItem "May"
.AddItem "Jun"
.AddItem "July"
.AddItem "Aug"
.AddItem "Sep"
.AddItem "Oct"
.AddItem "Nov"
.AddItem "Dec"
End With

With cmbEstimated
.AddItem "14:30"
.AddItem "14:45"
.AddItem "15:00"
.AddItem "15:30"
.AddItem "15:45"
.AddItem "16:00"
.AddItem "16:30"
.AddItem "16:45"
.AddItem "17:00"
.AddItem "17:30"
.AddItem "17:45"
.AddItem "18:00"
.AddItem "18:30"
.AddItem "18:45"
.AddItem "19:00"
.AddItem "19:15"
.AddItem "19:30"
.AddItem "19:45"
.AddItem "20:00"
.AddItem "20:00"
.AddItem "20:15"
.AddItem "20:30"
End With

With cmbTrailer
.AddItem "Not Applicable"
.AddItem "NDM 11949"
.AddItem "NDM 11804"
.AddItem "NDM 11959"
.AddItem "NDM 17825"
.AddItem "NDM 17826"
.AddItem "NDM 17737"
.AddItem "NDM 19328"
.AddItem "NDM 11951"
End With

With cmbNoLoads
.AddItem "Not Applicable"
.AddItem "01"
.AddItem "02"
.AddItem "03"
.AddItem "04"
End With

With cmbYear
.AddItem "2021"
.AddItem "2022"
.AddItem "2023"
.AddItem "2024"
.AddItem "2025"
.AddItem "2026"
.AddItem "2027"
.AddItem "2028"
End With


Call Reset

End Sub
