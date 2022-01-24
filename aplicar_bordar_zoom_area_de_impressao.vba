Sub AplicarBordarZoomAreaDeImpressao()
'
' AplicarBordarZoomAreaDeImpressao Macro
'
' Atalho do teclado: Ctrl+Shift+F
'
    Dim var_PlanilhaAtiva As Worksheet ' cria uma variável para a planilha ativa
    Dim var_IntervaloSelecionado As Range ' cria uma variável para o intervalo selecionado

    Set var_PlanilhaAtiva = ActiveSheet ' atribui a planilha ativa a variável var_PlanilhaAtiva
    Set var_IntervaloSelecionado = Selection ' seleciona o intervalo
    
    ActiveWindow.View = xlPageBreakPreview ' exibe a visualização de quebra de página
    ActiveWindow.Zoom = 100 ' ajusta o zoom para 100%
    
    var_PlanilhaAtiva.PageSetup.PrintArea = var_IntervaloSelecionado.Address ' define o intervalo de impressão
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone ' remove a borda diagonal
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone ' remove a borda diagonal
    With Selection.Borders(xlEdgeLeft) ' define a borda esquerda
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop) ' define a borda superior
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom) ' define a borda inferior
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight) ' define a borda direita
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical) ' define a borda vertical interna
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal) ' define a borda horizontal interna
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub