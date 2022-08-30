Attribute VB_Name = "Módulo1"
Sub ImprimeEtiquetasBin()
Attribute ImprimeEtiquetasBin.VB_Description = "Imprime as etiquetas de bin"
Attribute ImprimeEtiquetasBin.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' ImprimeEtiquetasBin Macro
' Imprime as etiquetas de bin
'
' Atalho do teclado: Ctrl+Shift+P
'

Dim first_page_data As Range
Dim second_page_data As Range
Dim valida_second_page As Boolean
Dim copy_first_page_data As Range

Set first_page_data = Range("B5:I24")
Set second_page_data = Range("B25:I45")
Set copy_first_page_data = Range("BB5:BI24")

    Sheets("ETIQ. BIN").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
        Sheets("PREENCHER").Select
    Range("B3").Select
    
For Each C In second_page_data
    If C.Value <> "" Then
        valida_second_page = True
        Exit For
    Else: valida_second_page = False
    End If
Next

If valida_second_page Then
    
    first_page_data.Select
    Selection.Copy
    copy_first_page_data.Select
    ActiveSheet.Paste
    
    
    second_page_data.Select
    Selection.Copy
    first_page_data.Select
    ActiveSheet.Paste
    
    Sheets("ETIQ. BIN").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Sheets("PREENCHER").Select
    
    
    copy_first_page_data.Select
    Selection.Copy
    first_page_data.Select
    ActiveSheet.Paste
    
    Range("B3").Select
End If
End Sub
