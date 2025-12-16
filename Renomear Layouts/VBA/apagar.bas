Attribute VB_Name = "apagar"
Sub ap()
Attribute ap.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ap Macro
'

'
    Columns("A:C").Select
    Range("B1").Activate
    Selection.ClearContents
    Range("A1").Select
End Sub
