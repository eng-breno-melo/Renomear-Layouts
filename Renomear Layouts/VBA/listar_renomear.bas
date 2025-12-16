Attribute VB_Name = "listar_renomear"
Option Explicit

Dim AcadApp As Object
Dim ThisDrawing As Object

Sub ConnectCAD()
    On Error Resume Next
    ' Primeiro tenta GstarCAD
    Set AcadApp = GetObject(, "GStarCAD.Application")
    ' Se não encontrar, tenta AutoCAD
    If AcadApp Is Nothing Then
        Set AcadApp = GetObject(, "AutoCAD.Application")
    End If
    ' Se ainda nada, mostra mensagem e sai
    If AcadApp Is Nothing Then
        MsgBox "Nenhuma instância do GstarCAD ou AutoCAD encontrada. Abra o CAD primeiro.", vbCritical
        Exit Sub
    End If
    ' Conecta ao desenho ativo
    Set ThisDrawing = AcadApp.ActiveDocument
    On Error GoTo 0
End Sub

' Listar layouts no Excel
Sub ListarLayouts()
    Call ConnectCAD
    If ThisDrawing Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Dim i As Long
    Dim lay As Object
    
    Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.Clear
    ws.Range("A1") = "Layout Atual"
    ws.Range("B1") = "Novo Nome"

    i = 2
    For Each lay In ThisDrawing.Layouts
        If lay.Name <> "Model" Then
            ws.Cells(i, 1).Value = lay.Name
            ws.Cells(i, 2).Value = lay.Name ' copia o mesmo nome p/ edição
            i = i + 1
        End If
    Next
    MsgBox "Layouts listados na planilha.", vbInformation
End Sub

' Renomear layouts com base no Excel
Sub RenomearLayouts()
    Call ConnectCAD
    If ThisDrawing Is Nothing Then Exit Sub
    
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim oldName As String, newName As String
    
    Set ws = ThisWorkbook.Sheets(1)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        oldName = Trim(ws.Cells(i, 1).Value)
        newName = Trim(ws.Cells(i, 2).Value)
        
        If oldName <> "" And newName <> "" Then
            If oldName <> newName Then
                On Error Resume Next
                ThisDrawing.Layouts(oldName).Name = newName
                If Err.Number <> 0 Then
                    ws.Cells(i, 3).Value = "Erro: " & Err.Description
                    Err.Clear
                Else
                    ws.Cells(i, 3).Value = "Renomeado"
                End If
                On Error GoTo 0
            Else
                ws.Cells(i, 3).Value = "Mantido"
            End If
        End If
    Next
    
    MsgBox "Processo concluído.", vbInformation
End Sub

