'COMANDOS VBA EXCEL

Range("H1048576").End(xlUp).Offset(2, 1).Select
' Range("H1048576"): Seleciona a linha 1048576 da coluna H (Última linha)
' End(xlUp): Pressiona CTRL + UP para subir pra última linha preenchida na coluna H
' Offset(2, 1): Adiciona 2 linhas e 1 coluna na seleção
' Select: Seleciona o intervalo

Cells.Select
' Cells: Seleciona todas as células

'==========================================================================================================================

Sub VerificandoSeAbaExiste()
Dim AbaExiste As Boolean

For i = 1 To Worksheets.Count ' faz um loop até o número de abas
    If Worksheets(i).Name = "TestandoSom" Then ' verifica se a aba existe
        AbaExiste = True ' se existir, AbaExiste recebe True
    End If
Next i

If Not AbaExiste Then
    ' se a aba não existir, cria uma nova aba
    Worksheets.Add.Name = "TestandoSom"

    ' se a aba não existir, exiba uma mensagem
    MsgBox "Aba TestandoSom criada com sucesso!"
End If
End Sub

'==========================================================================================================================
Sub Start()
    Call ProcurarColunaPorNome(3, "GuStAvO")
    MsgBox "Terminou"
End Sub

'==========================================================================================================================
Sub ProcurarColunaPorNome(NumeroDaLinha As Integer, NomeDaColuna As String)
    var_Coluna = 1 ' começa na coluna A
    NomeDaColuna = UCase(Trim(NomeDaColuna)) ' converte o nome da coluna para maiúsculo e remove espaços em branco
    'var_EnderecoDaCelula = ""
    var_ColunaEncontrada = False

    While Cells(NumeroDaLinha, var_Coluna).Value <> "" And Not var_ColunaEncontrada
        'var_EnderecoDaCelula = Cells(NumeroDaLinha, var_Coluna).Address  ' Retorna o endereço da célula
        If UCase(Trim(Cells(NumeroDaLinha, var_Coluna).Value)) = NomeDaColuna Then
            'Cells(NumeroDaLinha, var_Coluna).Select ' Seleciona a célula
            'MsgBox var_EnderecoDaCelula
            'MsgBox f_ObterLetraDaColuna(var_Coluna)
            Range(f_ObterLetraDaColuna(var_Coluna) & "1048576").End(xlUp).Offset(1, 0).Select
            var_ColunaEncontrada = True
        'Else
            'Cells(NumeroDaLinha, var_Coluna).Value = "NADA"
        End If

        var_Coluna = var_Coluna + 1
    Wend

    If Not var_ColunaEncontrada Then
        MsgBox "Não encontrou a coluna " & NomeDaColuna
    End If
End Sub

'==========================================================================================================================

Function f_ObterLetraDaColuna(NumeroDaColuna As Variant) As String
    Dim var_Array
    var_Array = Split(Cells(1, NumeroDaColuna).Address(True, False), "$")
    f_ObterLetraDaColuna = var_Array(0)
End Function

'==========================================================================================================================
Sub VerificaQuantidadeCelulasPorCriterio()

Dim var_Intervalo As Range ' Cria uma variável do tipo Range
Dim var_Criterio As String ' Cria uma variável do tipo String
Dim var_Resultado As Integer ' Cria uma variável do tipo Integer

Set var_Intervalo = Range("A1:G1") ' Define o intervalo e atribui a variável

var_Criterio = "*Gustavo*" ' Contenha o critério
'var_Criterio = "Gustavo" ' Seja igual ao critério
'var_Criterio = ">=1" ' Seja maior ou igual ao critério
'var_Criterio = "<=1" ' Seja menor ou igual ao critério
'var_Criterio = "=1" ' Seja igual ao critério
'var_Criterio = "<>" ' Não seja em branco


var_Resultado = WorksheetFunction.CountIf(var_Intervalo, var_Criterio) ' Countif retorna a quantidade de células que atendem ao critério
'var_Resultado = Evaluate("Sum(COUNTIF(A1:G1,{""Gustavo"",""Jefferson""}))") ' Multiplos critérios




MsgBox var_Resultado ' Exibe o resultado

End Sub

' ===========================================================================================================================

Sub VerificaQuantidadeCelulasPorCriterios()

Dim var_Intervalo As Range ' Cria uma variável do tipo Range
Dim var_Criterio1 As String ' Cria uma variável do tipo String
Dim var_Criterio2 As String ' Cria uma variável do tipo String
Dim var_Resultado As Integer ' Cria uma variável do tipo Integer

var_Criterio1 = ">0" ' Seja maior que 0
var_Criterio2 = "<=10" ' Seja menor ou igual a 10

Set var_Intervalo = Range("A1:G1") ' Define o intervalo e atribui a variável

var_Resultado = WorksheetFunction.CountIfs(var_Intervalo, var_Criterio1, var_Intervalo, var_Criterio2) ' Countifs retorna a quantidade de células que atendem aos critérios
' Contifs utiliza o AND entre os critérios

If var_Resultado = 0 Then
    MsgBox "Não encontrou nenhuma célula"
Else
    MsgBox var_Resultado
End If

End Sub

' ===========================================================================================================================

Function f_LarguraDaColuna(Largura As Long)
    Cells.ColumnWidth = Largura
    Cells.Rows.AutoFit
End Function

' ===========================================================================================================================


' ===========================================================================================================================
Sub Copiar_Colar_Valor_AllSheets()
    Application.ScreenUpdating = False
    Dim var_Worksheet As Worksheet
    Dim var_Range As Range
    For Each var_Worksheet In ActiveWorkbook.Worksheets
        var_Worksheet.Activate
        Call f_Copiar_Colar_Valor_AllCells()
        Range("A1").Select
    Next
    Worksheets(1).Activate 
    Application.ScreenUpdating = True
    MsgBox "Processo finalizado"
End Sub

Sub Remover_Formatacao_Condicional_AllSheets()
    Application.ScreenUpdating = False
    Dim var_Worksheet As Worksheet
    Dim var_Range As Range
    For Each var_Worksheet In ActiveWorkbook.Worksheets
        var_Worksheet.Activate
        Call f_Remover_Formatacao_Condicional_AllCells()
        Range("A1").Select
    Next
    Worksheets(1).Activate 
    Application.ScreenUpdating = True
    MsgBox "Processo finalizado"
End Sub

Sub ENVIAR_MexBeneficios()
    Application.ScreenUpdating = False
    
    Dim var_Worksheet As Worksheet
    Dim var_Range As Range
    For Each var_Worksheet In ActiveWorkbook.Worksheets
        var_Worksheet.Activate
        Call f_Remover_Formatacao_Condicional_AllCells()
        Call f_Copiar_Colar_Valor_AllCells()

        If UCase(Trim(var_Worksheet.Name)) = "DADOS" Then
            Set var_Range = Range("A1:AZ1")
            Call f_Procurar_Deletar_Coluna(var_Range, "SERVE = 1")
            Call f_Procurar_Deletar_Coluna(var_Range, "ORG")
            Call f_Procurar_Deletar_Coluna(var_Range, "VVTNOVO")
            Call f_Procurar_Deletar_Coluna(var_Range, "TVTNOVO")
            Call f_Procurar_Deletar_Coluna(var_Range, "RECPEND")
            Call f_Procurar_Deletar_Coluna(var_Range, "SALDO1")
            Call f_Procurar_Deletar_Coluna(var_Range, "TIPO1")
            Call f_Procurar_Deletar_Coluna(var_Range, "CNPJ + CPF + OPERADORA")
            Call f_Procurar_Deletar_Coluna(var_Range, "CNPJ + CPF + TIPO1")
            Call f_Procurar_Deletar_Coluna(var_Range, "BUSCADOR")
            Call f_Procurar_Deletar_Coluna(var_Range, "ORDEM")
            Call f_Procurar_Deletar_Coluna(var_Range, "CF -R$10")
        End If
        Range("A2").Select
    Next
    Worksheets(1).Activate 
    Application.ScreenUpdating = True
    MsgBox "Processo finalizado"
End Sub

Function f_Copiar_Colar_Valor_AllCells()
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
End Function

Function f_Remover_Formatacao_Condicional_AllCells()
    Cells.FormatConditions.Delete
End Function


Function f_Procurar_Deletar_Coluna(Intervalo As Range, NomeDaColuna As String)
    Dim var_Acabou As Boolean
    var_Acabou = False

    While var_Acabou = False
        For Each Cell In Intervalo
            If Trim(UCase(Cell.Value)) = NomeDaColuna Then
                Cell.EntireColumn.Delete
                var_Acabou = False
                Exit For
            Else
            var_Acabou = True
            End If
        Next
    Wend
End Function



Sub Teste4()
    Dim var_Worksheet As Worksheet
    Dim var_CUnid As Range

    Set var_Worksheet = Sheets("Compra")
    Call f_Procurar_Mover_Coluna(var_Worksheet, "c.unid", 1)
    Call f_Procurar_Mover_Coluna(var_Worksheet, "operadora", 1)
    Call f_Procurar_Mover_Coluna(var_Worksheet, "uf", 1)
    Call f_Procurar_Mover_Coluna(var_Worksheet, "empresa", 1)
End Sub

Function f_Procurar_Mover_Coluna(var_Worksheet As Worksheet, var_NomeColuna As String, var_PosicaoColuna As Integer)
Dim var_Range As Range

Set var_Range = var_Worksheet.Range("1:1").Find(What:=var_NomeColuna, MatchCase:=False)
    If Not var_Range Is Nothing Then
        Columns(var_Range.Cells.Column).Cut
        Columns(var_PosicaoColuna).Insert Shift:=xlToRight
    End If

End Function





Sub Criar_Compra_BETA()
    Dim var_Worksheet As Worksheet

    Dim var_RangeUF As Range
    Dim var_RangeOperadora As Range
    Dim var_RangeEmpresa As Range
    Dim var_RangeCUnid As Range

    Dim var_ColunaUFExiste As Boolean
    Dim var_ColunaOperadoraExiste As Boolean
    Dim var_ColunaEmpresaExiste As Boolean
    Dim var_ColunaCUnidExiste As Boolean

    Set var_Worksheet = Sheets("Compra")
    var_Worksheet.Activate

    Set var_RangeUF = var_Worksheet.Range("1:1").Find(What:="uf", MatchCase:=False)
    Set var_RangeOperadora = var_Worksheet.Range("1:1").Find(What:="operadora", MatchCase:=False)
    Set var_RangeEmpresa = var_Worksheet.Range("1:1").Find(What:="empresa", MatchCase:=False)
    Set var_RangeCUnid = var_Worksheet.Range("1:1").Find(What:="c.unid", MatchCase:=False)

    'Set var_RangeOrg = var_Worksheet.Range("1:1").Find(What:="org", MatchCase:=False)

    If Not var_RangeCUnid Is Nothing Then
        Columns(var_RangeCUnid.Cells.Column).Cut
        Columns(1).Insert Shift:=xlToRight
        var_ColunaCUnidExiste = True
    Else
        MsgBox "A coluna [C.UNID] não encontrado"
        Exit Sub
    End If

    If Not var_RangeEmpresa Is Nothing Then
        Columns(var_RangeEmpresa.Cells.Column).Cut
        Columns(1).Insert Shift:=xlToRight
        var_ColunaEmpresaExiste = True
    Else
        MsgBox "A coluna [EMPRESA] não encontrado"
        Exit Sub
    End If

    If Not var_RangeOperadora Is Nothing Then
        Columns(var_RangeOperadora.Cells.Column).Cut
        Columns(1).Insert Shift:=xlToRight
        var_ColunaOperadoraExiste = True
    Else
        MsgBox "A coluna [OPERADORA] não encontrado"
        Exit Sub
    End If

    If Not var_RangeUF Is Nothing Then
        Columns(var_RangeUF.Cells.Column).Cut
        Columns(1).Insert Shift:=xlToRight
        var_ColunaUFExiste = True
    Else
        MsgBox "A coluna [UF] não encontrado"
        Exit Sub
    End If

End Sub