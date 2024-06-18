Attribute VB_Name = "ModData"
Option Explicit
Public Function GetData() As Range
'Define the initial range
Set GetData = Sheet1.Range("A1").CurrentRegion

'Now exclude the first row since it contains the header
Set GetData = GetData.Offset(1).Resize(GetData.Rows.Count - 1)

End Function

Public Function DeleteRow(ByVal row As Long)
Sheet1.Range("A2").Offset(row).EntireRow.Delete
End Function

Public Function GetNewID() As Long

GetNewID = 1 + WorksheetFunction.Max(Sheet1.Range("A2").CurrentRegion.Columns(1))

End Function

Public Function GetCountries() As Variant
GetCountries = Sheets("Lookup").ListObjects("Table1").DataBodyRange.Value
End Function


Public Function GetDepartments() As Variant
GetDepartments = Sheets("Lookup").ListObjects("Table2").DataBodyRange.Value
End Function

