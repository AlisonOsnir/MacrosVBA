Attribute VB_Name = "Gerar_relatorios_Treetech"

Sub gerar_relat�rio_treetech()
Attribute gerar_relat�rio_treetech.VB_Description = "Ao selecionar o primeiro serial e executar gera relat�rios em PDF e exportados para uma pasta especifica.\n    "
Attribute gerar_relat�rio_treetech.VB_ProcData.VB_Invoke_Func = "r\n14"
    '
    ' Atalho do teclado: Ctrl+R
    '
    ' Ao selecionar o primeiro serial no banco de dados e executar a macro, utiliza a celula principal do relat�rio e os exporta em PDF para uma pasta especifica.
    ' Desde que o relat�rio esteja com as procuras(PROCV) corretas dever� funcionar com qualquer relat�rio.
    
    Dim celula_procv As String
    Dim arq_destino As String
    Dim arq_nome As String
    Dim qtd_relatorios As Integer
    Dim i As Integer
    Dim linhas_selecionadas As Integer
    Dim nome_aba_relatorios As String
    Dim nome_aba_banco_dados As String

    Application.ScreenUpdating = False
    
  
    '*********************************************************
    '****************_ MANUTEN��O DE CODIGOS _****************
    
    nome_aba_relatorios = "Relat�rio de Ensa�os"
    nome_aba_banco_dados = "PADR�O ABSOLUT"
    
    celula_procv = "E9"
    arq_destino = "\\adserver\Publico\Alison\BACKUPS ~ N�O MEXER\EXCEL Macro Relat�rios\teste relatorios treetech\"
    
    '*********************************************************
    '*********************************************************
    
    
    'input quantidades
    linhas_selecionadas = Selection.Rows.Count
    qtd_relatorios = Application.InputBox(Title:="Gerar Relat�rios", Type:=1, Prompt:="Quantidade de relat�rios:", Default:=linhas_selecionadas)
    Selection.Rows(1).Select

    Sheets(nome_aba_banco_dados).Select
    If qtd_relatorios <> 0 Then
        i = 0
        Do While i < qtd_relatorios
            If Selection.Value = "" Then
                ActiveCell.Offset(1, 0).Range("A1").Select
            Else
                Selection.Copy
                Sheets(nome_aba_relatorios).Select
                Range(celula_procv).Select
                ActiveSheet.Paste
                Application.CutCopyMode = False
                
                arq_nome = Range(celula_procv).Value 'serial como nome de arquivo
                
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                arq_destino & arq_nome & ".pdf", _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, ignorePrintAreas _
                :=False, OpenAfterPublish:=False
                
                Sheets(nome_aba_banco_dados).Select
                ActiveCell.Offset(1, 0).Range("A1").Select
                
                i = i + 1
            End If
        Loop
        
        If i = qtd_relatorios Then
            If qtd_relatorios = 1 Then
                MsgBox (qtd_relatorios & " Relat�rio foi gerado!")
            Else
                MsgBox (qtd_relatorios & " Relat�rios foram gerados!")
            End If
        End If
    End If
End Sub
