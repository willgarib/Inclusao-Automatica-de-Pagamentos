Attribute VB_Name = "Módulo2"
' Módulo destinado a escrita de Funções de Apoio

Function NumeroCPF(text As String)
''' Retorna o CPF sem traços e pontos

Dim saida As String
saida = ""

For i = 1 To Len(text)
    caracter = Mid$(text, i, 1)
    If IsNumeric(caracter) Then
        saida = saida + caracter
    End If
Next
NumeroCPF = saida
End Function

Function TextoCPF(text As String)
''' Retorna o CNPJ com números e traços

TextoCPF = Mid$(text, 1, 3) & "." & Mid$(text, 4, 3) & "." & Mid$(text, 7, 3) & "-" & Mid$(text, 10, 2)

End Function

Function QtdDepoisVirg(n As String)

If InStr(1, n, ",") = 0 Then
    QtdDepoisVirg = 0
    Exit Function
End If

QtdDepoisVirg = Len(n) - InStr(1, n, ",")

End Function

Function CorrigeDin(valor As String)

Select Case QtdDepoisVirg(valor)
    Case Is = 2
        CorrigeDin = Replace(valor, ",", "") ' Tira a vírgula
    Case Is = 1
        CorrigeDin = Replace(valor, ",", "") & "0"  ' Tira a vírgula e concatena com UM zero
    Case Is = 0
        CorrigeDin = Replace(valor, ",", "") & "00" ' Tira a vírgula e concatena com DOIS zeros
End Select

End Function

Function CompletaDireita(total, texto)
' Completa um texto com brancos a direira

If Len(texto) = total Then
    CompletaDireita = texto
    Exit Function
End If

CompletaDireita = texto & String(total - Len(texto), " ")

End Function

Function CompletaEsquerda(total, texto)
' Completa um texto com zeros a esquerda

If Len(texto) = total Then
    CompletaEsquerda = texto
    Exit Function
End If

CompletaEsquerda = String(total - Len(texto), "0") & texto

End Function
