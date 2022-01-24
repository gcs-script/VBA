'COMANDOS VBA EXCEL

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

    Set var_RangeOrg = var_Worksheet.Range("1:1").Find(What:="org1", MatchCase:=False)
    Set var_RangeCompraFinal = var_Worksheet.Range("1:1").Find(What:="comprafinal", MatchCase:=False)

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

    ' DELETAR COLUNAS

End Sub