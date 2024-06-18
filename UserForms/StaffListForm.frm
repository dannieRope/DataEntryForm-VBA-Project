VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StaffListForm 
   Caption         =   "Staff List"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12660
   OleObjectBlob   =   "StaffListForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StaffListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CloseCommandButton_Click()
Unload Me

End Sub

Private Sub DeleteCommandButton_Click()
DeleteRow (staffListBox.ListIndex)
End Sub

Private Sub EditCommandButton_Click()
Call EditRow
End Sub

Private Sub NewCommandButton_Click()
 Dim frm As New NewStaffFormSave
 frm.Show
 
 Call AddDataToListBox

 
End Sub

Private Sub UserForm_Activate()
Call AddDataToListBox
End Sub

Private Sub AddDataToListBox()
'firstly, we get the range(Data)
Dim rng As Range
Set rng = GetData()
'Then link the range(Data) to the ListBox
       With staffListBox
         .RowSource = rng.Address(External:=True) 'assign the address of the rng range
         .ColumnCount = rng.Columns.Count 'Set the number of columns = to the number of columns of the data
         .ColumnHeads = True 'Show headers
         .ColumnWidths = "30;85;85;150;140;100" 'adjust column width
         .ListIndex = 0 'set the initially selected item in the listbox As the first item
       End With
       
End Sub

Private Sub EditRow()
Dim frm As New NewStaffFormEdit
frm.CurrentRow = staffListBox.ListIndex
frm.Show vbModal

End Sub
