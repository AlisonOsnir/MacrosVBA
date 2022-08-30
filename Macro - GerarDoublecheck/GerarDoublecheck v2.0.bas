Attribute VB_Name = "GerarDoublecheck"
Sub ImprimeDoubleCheck()
Attribute ImprimeDoubleCheck.VB_Description = "Imprime DoubleCheck"
Attribute ImprimeDoubleCheck.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' ImprimeDoubleCheck Macro
' Imprime DoubleCheck
'
' Atalho do teclado: Ctrl+Shift+D
'

Dim first_page_data As Range
Dim second_page_data As Range
Dim valida_second_page As Boolean


Set first_page_data = Range("B8:G51")
Set second_page_data = Range("B52:G95")

    Sheets("DOUBLECHECK").Select
'    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Sheets("PREENCHER").Select
    
For Each c In second_page_data
    If c.Value <> "" Then
        valida_second_page = True
        Exit For
    Else: valida_second_page = False
    End If
Next

If valida_second_page Then
    second_page_data.Select
    Selection.Copy
    first_page_data.Select
    ActiveSheet.Paste
    
    Sheets("DOUBLECHECK").Select
'   ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Sheets("PREENCHER").Select
End If
    
End Sub
Sub ImprimeMascara()
Attribute ImprimeMascara.VB_Description = "Imprime Mascara em A4"
Attribute ImprimeMascara.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' ImprimeMascara Macro
' Imprime Mascara em A4
'
' Atalho do teclado: Ctrl+Shift+M
'
    Sheets("MASCARA").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Sheets("PREENCHER").Select
End Sub
