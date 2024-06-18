VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewStaffFormSave 
   Caption         =   "Edit Staff Member"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5955
   OleObjectBlob   =   "NewStaffFormSave.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewStaffFormSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CloseCommandButton_Click()
Unload Me

End Sub

Private Sub SaveCommandButton_Click()
If MsgBox("Do you wish to save this record?", vbYesNo, "Save Record") = vbYes Then
    'Write the data to the worksheet from controls
    Call WriteDataToSheet
    'Empty the textboxes
    Call emptyTextBoxes
    'Create a new Id
    Call CreateNewId
End If

End Sub
Private Sub WriteDataToSheet()
    Dim newRow As Long
    With Sheet1
        newRow = .Cells(.Rows.Count, 1).End(xlUp).row + 1
        
        .Cells(newRow, 1).Value = IDTextBox.Value
        .Cells(newRow, 2).Value = FirstNameTextBox.Value
        .Cells(newRow, 3).Value = LastNameTextBox.Value
        .Cells(newRow, 4).Value = CountryComboBox.Value
        .Cells(newRow, 5).Value = IIf(FulltimeOptionButton.Value = True, "Full-Time", "Part-Time")
        .Cells(newRow, 6).Value = DepartmentComboBox.Value
    End With

End Sub
Private Sub emptyTextBoxes()
Dim c As Control
For Each c In Me.Controls
    If TypeName(c) = "TextBox" Then
        c.Value = " "
    End If
Next c

End Sub

Private Sub UserForm_Activate()
'Create a new id
    Call CreateNewId
'Initialize the controls
    Call InitializeControls
End Sub

Private Sub CreateNewId()
Me.IDTextBox.Value = GetNewID()

End Sub

Private Sub InitializeControls()
Me.CountryComboBox.List = GetCountries()
Me.CountryComboBox.ListIndex = 0
Me.DepartmentComboBox.List = GetDepartments()
Me.DepartmentComboBox.ListIndex = 0

Me.FulltimeOptionButton.Value = True

End Sub


