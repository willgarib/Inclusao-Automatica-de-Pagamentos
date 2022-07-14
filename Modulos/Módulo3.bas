Attribute VB_Name = "Módulo3"
' Cabeçalhos e rodapés

Sub cabecaArquivo()
    ''' Carrega o cabeçalho do Arquivo '''

' Zera as variaveis de controle [célula H4 da 'Lote']
With Sheets("Lote")
    .Range("H4").Value = 1 '[N° do lote]
    .Range("H7:I7").ClearContents
    .Range("J7") = 0 '[Qtd registros cumulado]
End With

With Sheets("Saída")
    'Limpa a saída
    .Cells(3, 2).CurrentRegion.Clear
End With

' Carrega
With Sheets("Arquivo")
    headArq = "34100000      0802" & .Range("F4").Value & "                    " & .Range("F5").Value & " " & .Range("F6").Value & " " & .Range("F7").Value & .Range("F8").Value & "ITAÚ UNIBANCO S.A                       1" & .Range("F9").Value & .Range("F10").Value & "00000000000000                                                                     "
End With

Sheets("Saída").Cells(3, 2).Value = headArq

End Sub

Sub rodapeArq()
    ''' Carrega o rodapé do Arquivo '''

Dim ult_lin As Long
ult_lin = Sheets("Saída").Cells(100000, 2).End(xlUp).Row + 1

tailArq = "34199999" & String(9, " ") & CompletaEsquerda(6, Sheets("Lote").Range("H4").Value - 1) & CompletaEsquerda(6, Sheets("Lote").Range("J7").Value + 2) & String(211, " ")

Sheets("Saída").Cells(ult_lin, 2).Value = tailArq
End Sub

Sub cabecaLote()
    ''' Carrega o cabeçalho do Lote '''
    
' Dimensionamento das Variáveis
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Dim ult_lin As Long
ult_lin = Sheets("Saída").Cells(100000, 2).End(xlUp).Row + 1
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-

With Sheets("Lote")
    headLote = "341" & .Range("I4").Value & "1C" & .Range("F4").Value & .Range("F5").Value & "040 2" & .Range("F6").Value & "                    " & .Range("F7").Value & " " & .Range("F8").Value & " " & .Range("F9").Value & .Range("F10").Value & String(118, " ") & "SP" & String(18, " ")
End With

Sheets("Saída").Cells(ult_lin, 2).Value = headLote

End Sub

Sub rodapeLote()
    ''' Carrega o rodapé do Lote '''

' Dimensionamento das Variáveis
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Dim ult_lin As Long
ult_lin = Sheets("Saída").Cells(100000, 2).End(xlUp).Row + 1
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-

With Sheets("Lote")
    tailLote = "341" & .Range("I4").Value & "5000000000" & .Range("H7").Value & .Range("I7").Value & String(18, "0") & String(171, " ") & String(10, " ")
End With

Sheets("Saída").Cells(ult_lin, 2).Value = "'" & tailLote

End Sub
