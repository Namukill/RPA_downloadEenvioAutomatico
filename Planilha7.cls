VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub download()
    Dim Ret_Val
    Ret_Val = Shell("python """"", vbNormalFocus) ' Coloque aqui o caminho para o código "baixarExcel.py"
End Sub

Sub email()
    Dim Ret_Val
    Ret_Val = Shell("python """"", vbNormalFocus) ' Coloque aqui o caminho para o código "EnvioSemanalPgestores.py"
End Sub

Sub organizar()
' Definição das variáveis de workbook, worksheets e variáveis auxiliares
Dim ws As Worksheet
Dim wb As Workbook
Dim total As Integer
Dim linha As Integer
Dim caminho As String
Dim WB_base As Workbook
Dim WS_base As Worksheet
Dim ID As Integer
Dim ultimalinha_aberto As Integer
Dim ultimalinha_base As Integer
Dim linhasVisiveis As Integer
Dim rng As Range
Dim cel As Range

caminho = ActiveWorkbook.FullName
Application.ScreenUpdating = True
Application.DisplayAlerts = False
MsgBox "Selecione o último download do forms para ser alimentado!" ' Mensagem para o usuário selecionar o arquivo do Forms
SelecaoForms = Application.GetOpenFilename("Arquivos excel (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm") ' Abre um explorador de arquivos para o usuário selecionar o arquivo do Forms
Set wb = Workbooks.Open(SelecaoForms) ' Abre o arquivo selecionado, no caso, o download do Forms
Set ws = wb.Sheets(1) ' Como o formulário por padrão só tem uma aba, selecionamos a primeira aba

Set WB_base = Workbooks.Open("", UpdateLinks:=1) ' Coloque aqui o caminho para o arquivo base onde os dados serão salvos
Set WS_base = WB_base.Sheets(1) ' Seleciona a primeira aba do arquivo base

ultimalinha_base = WS_base.Cells(WS_base.Rows.Count, "A").End(xlUp).Row ' Conta a quantidade de linhas preenchidas na coluna A da aba base
ID = WS_base.Cells(ultimalinha_base, "A").Value ' Pega o ID da última linha preenchida na coluna A da aba base

ultimalinha_aberto = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Conta a quantidade de linhas preenchidas na coluna A do download
ws.Range("A1:A" & ultimalinha_aberto).AutoFilter field:=1, Criteria1:=">" & ID ' Filtra a coluna A para mostrar apenas os dados com ID maior que o último ID da aba base
ws.Activate

Set rng = ws.Range("A2:A" & ws.Cells(Rows.Count, "A").End(xlUp).Row).SpecialCells(xlCellTypeVisible) ' Define o intervalo filtrado da coluna A, começando da linha 2 até a última linha preenchida na coluna A
linhasVisiveis = 0
For Each cel In rng
    linhasVisiveis = linhasVisiveis + 1
Next cel ' Conta quantas linhas estão visíveis após o filtro para evitar copiar linhas ocultas

Application.CutCopyMode = False ' Limpa a área de transferência
ws.Range("A" & ultimalinha_aberto - linhasVisiveis + 1 & ":V" & ultimalinha_aberto).Copy ' Copia os dados filtrados da aba do download
WS_base.Cells(ultimalinha_base + 1, 1).PasteSpecial Paste:=xlPasteValues ' Cola os dados copiados
wb.Close SaveChanges:=False ' Fecha o arquivo do download sem salvar
Application.ScreenUpdating = True ' Ativa novamente a atualização de tela
Application.DisplayAlerts = True ' Ativa novamente os alertas
MsgBox "fim"
End Sub