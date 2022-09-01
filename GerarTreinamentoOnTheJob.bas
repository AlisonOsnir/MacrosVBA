Attribute VB_Name = "GeradorForm"
Sub GerarFormulários()
Attribute GerarFormulários.VB_Description = "Essa macro gera formulários de treinamento substituindo valores na Tab Aux por cada linha da lista."
Attribute GerarFormulários.VB_ProcData.VB_Invoke_Func = " \n14"
'
' GerarFormulários Macro
' Essa macro gera formulários de treinamento substituindo valores na Tab Aux por cada linha da lista.
'

    Dim i As Integer
    i = 0

    Application.ScreenUpdating = False
    
    'Formata lista de inclusão
    Sheets("Lista").Select
    Range("A3:h22").Borders.LineStyle = 1
    
    'Desprotege Tab Aux
    Sheets("Tab Aux").Select
    ActiveSheet.Unprotect
    
    'Manipula filtro
    Sheets("Treinamento On the job").Select
    Range("A2:H2").Select
    Selection.AutoFilter
    Selection.AutoFilter
      
    'Itera sobre linhas de inclusão
    Do While i < 20 'Numero de linhas
        Sheets("Lista").Select
        
        If i = 0 Then
            Range("A3:h3").Select
        Else
            ActiveCell.Offset(1, 0).Range("A1:h1").Select
        End If
               
        If Not WorksheetFunction.CountA(Range(Selection.Address)) < 7 Then
                 
            'Copia para refencia do form
            Selection.Copy
            Sheets("Tab Aux").Select
            Range("F4:M4").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                 :=False, Transpose:=False
            Selection.Copy
            
            'Inclui no controle
            Sheets("Treinamento On the job").Select
            Range("A3:h3").Select
            Selection.Insert Shift:=xlDown
        
            'Imprime formulario
            Sheets("FM 346").Select
            Application.CutCopyMode = False
            ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                IgnorePrintAreas:=False
        End If
        i = i + 1
    Loop
    
    'Protege Tab Aux
    Sheets("Tab Aux").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Sheets("Treinamento On the job").Select
End Sub
