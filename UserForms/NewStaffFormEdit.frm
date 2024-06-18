VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewStaffFormEdit 
   Caption         =   "Edit Staff Member"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5955
   OleObjectBlob   =   "NewStaffFormEdit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewStaffFormEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mCurrentRow As Long
Public Property Let CurrentRow(ByVal newCurrentRow As Long)
mCurrentRow = newCurrentRow
End Property

Private Sub CloseCommandButton_Click()
Unload Me

End Sub

Private Sub UpdateCommandButton_Click()
Call WriteDataToSheet
End Sub

Private Sub UserForm_Activate()
 'Fill the comboboxes and load data
 Call FillComboBoxes
 Call LoadData
 
End Sub

Public Sub FillComboBoxes()

Me.CountryComboBox.List = GetCountries()
Me.DepartmentComboBox.List = GetDepartments()
End Sub

Public Sub LoadData()
    With Sheet1.Range("A2").Offset(mCurrentRow)
          IDTextBox.Value = .Cells(1, 1).Value
          FirstNameTextBox = .Cells(1, 2).Value
          LastNameTextBox.Value = .Cells(1, 3).Value
          CountryComboBox.Value = .Cells(1, 4).Value
          FulltimeOptionButton.Value = IIf(.Cells(1, 5).Value = "Full-time", True, False)
          PartTimeOptionButton.Value = IIf(.Cells(1, 5).Value = "Part-time", True, False)
          DepartmentComboBox.Value = .Cells(1, 6).Value
     
    End With
    
End Sub
Public Sub WriteDataToSheet()
If MsgBox("Are you sure you wist to update this record", vbYesNo, "Update record") = vbYes Then
    With Sheet1.Range("A2").Offset(mCurrentRow)
        .Cells(1, 1).Value = IDTextBox.Value
        .Cells(1, 2).Value = FirstNameTextBox
        .Cells(1, 3).Value = LastNameTextBox.Value
        .Cells(1, 4).Value = CountryComboBox.Value
        .Cells(1, 5).Value = IIf(FulltimeOptionButton.Value = True, "Full-time", "Part-time")
        .Cells(1, 6) = DepartmentComboBox.Value
    End With
End If
End Sub
