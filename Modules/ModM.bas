Attribute VB_Name = "ModM"
Option Explicit

Public Sub main()
Attribute main.VB_ProcData.VB_Invoke_Func = "q\n14"
' Declare and instantiate a new instance of the form
    Dim frm As New StaffListForm
' Show the form
    frm.Show
' Once the form is closed, release the form object and clean up memory
    Set frm = Nothing
'Go to macro and set short cut (ctrl + q) for the main subroutine

End Sub
