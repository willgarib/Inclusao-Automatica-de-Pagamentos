Attribute VB_Name = "M�dulo4"
' Detalhes do Lote Transfer�ncia

Sub CompilaLote(planilha As String)
    ''' Compila e carrega todos os registros do lote que est�o na planilha "Lote Detalhe" '''
    
' Dimensionamento das Vari�veis
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Dim ult_lin As Long
ult_lin = Sheets(planilha).Cells(100000, 2).End(xlUp).Row

Dim linha As Long
Dim Din As String
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-

Call cabecaLote ' Carrega o cabe�alho do lote

' Executa a forma��o do registro e carrega para "Sa�da" cada linha do lote cadastrado em "Lote Detalhe"
For linha = 5 To ult_lin

    Call geraRegistroDetalhe(linha, planilha)
    
Next

' Salva Informa��es Sobre o Lote na Planilha "Lote" para Composi��o do Rodap� do Lote
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Sheets("Lote").Range("H7").Value = "'" & CompletaEsquerda(6, ult_lin + 2 - 4) ' Qtd Registros Detalhe desse Lote [H7 da Lote]

Sheets("Lote").Range("I7").Value = Application.Sum(Sheets(planilha).Range("I5:I" & ult_lin))  ' Salva a soma do lote [I7 da Lote] ($)
Din = Format(Sheets("Lote").Range("I7").Value, "@") '
Din = CorrigeDin(Din)                               ' Din tem o valor (str) correto s� falta corrigir
Sheets("Lote").Range("I7").Value = "'" & CompletaEsquerda(18, Din) ' Coloca na c�lula da planilha o valor completado
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-

Call rodapeLote ' Carrega o Rodap� do Lote

' Adiciona 1 ao C�digo do Lote para o Proximo Lote (Sequ�ncial por Arquivo)
Sheets("Lote").Range("H4").Value = Sheets("Lote").Range("H4").Value + 1
    
' Salva Informa��es Sobre o Lote na Planilha "Lote" para Composi��o do Rodap� do Arquivo
Sheets("Lote").Range("J7").Value = Sheets("Lote").Range("J7").Value + ult_lin + 2 - 4 ' Soma ao Acumulado (Qtd de Registros)

End Sub

Sub geraRegistroDetalhe(linha As Long, planilha As String)
    ''' Gera sequ�ncia de caracteres que � um registro de detalhe de lote '''

' Dimensionamento das Vari�veis
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Dim ult_lin_saida As Long
ult_lin_saida = Sheets("Sa�da").Cells(100000, 2).End(xlUp).Row + 1

Dim registro As String
registro = ""

Dim AC As String
AC = ""

Dim textdim As String
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-

With Sheets(planilha)

    registro = registro & "341" ' Banco ITA�
    
    registro = registro & Sheets("Lote").Range("I4").Value ' C�digo do Lote
    
    registro = registro & "3" ' Campo Fixo
    
    registro = registro & CompletaEsquerda(5, linha - 4) ' Sequencial Detalhe
    
    registro = registro & "A" ' Tipo do Registro
    
    registro = registro & "000" ' Tipo de Movimenta��o [000 para inclus�o de pagamento]
    
    registro = registro & "000" ' C�mara
    
    registro = registro & CompletaEsquerda(3, .Cells(linha, 2).Value) ' Banco Favorecido [Coluna B]
    
    ' Definindo a Ag�ncia e Conta (AC)
    If .Cells(linha, 2).Value = "341" Or .Cells(linha, 2).Value = "409" Then ' ITA� e UNIBANCO
        AC = AC & "0" ' Campo Fixo
        AC = AC & CompletaEsquerda(4, .Cells(linha, 3).Value) 'Ag�ncia [Coluna C]
        AC = AC & " " ' Campo Fixo
        AC = AC & String(6, "0") ' Campo Fixo
        AC = AC & CompletaEsquerda(6, .Cells(linha, 5).Value) 'Conta & DAC [Coluna E]
        AC = AC & " " ' Campo Fixo
        AC = AC & CompletaEsquerda(1, .Cells(linha, 6).Value) ' DAC Ag�ncia [Coluna F]
    Else:                                                                    ' Demais Bancos
        AC = AC & CompletaEsquerda(5, .Cells(linha, 3).Value & .Cells(linha, 4).Value) 'Ag�ncia [Coluna C]
        AC = AC & " " ' Campo Fixo
        AC = AC & CompletaEsquerda(12, .Cells(linha, 5).Value & .Cells(linha, 6).Value) 'Conta e DAC da Conta [Coluna E e F]
        AC = AC & " " ' Campo Fixo
        AC = AC & CompletaDireita(1, .Cells(linha, 4).Value) 'DAC Ag�ncia [Coluna D]
    End If
    
    registro = registro & AC ' Concatena AC ao registro
    
    registro = registro & CompletaDireita(30, Left(.Cells(linha, 7).Value, 30)) ' Nome do Favorecido [Coluna G]
    
    registro = registro & CompletaDireita(20, Sheets("Lote").Range("J4").Value) ' Seu N�mero
    Sheets("Lote").Range("J4").Value = Sheets("Lote").Range("J4").Value + 1     ' Adiciona 1 ao "seu n�mero"
    
    ' Data Pagamento (Previs�o) [Coluna H]
    ' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    textdata = Format(.Cells(linha, 8).Value, "@") ' Formata como texto
    textdata = Replace(textdata, "/", "")                                ' Tira a barra
    registro = registro & textdata                                       ' Concatena com registro
    ' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    
    registro = registro & "009" ' Tipo Moeda (009 para REAL)
    
    registro = registro & String(8, " ") ' C�mara 8*B
    
    registro = registro & String(7, "0") ' Campo Fixo
    
    ' Valor Previsto [Coluna I]
    ' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    textdim = Format(.Cells(linha, 9).Value, "@") ' Formata com Texto
    textdim = CorrigeDin(textdim)                                       ' Tira a v�rgula e Corrige
    registro = registro & CompletaEsquerda(15, textdim)                 ' Concatena com registro
    ' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    
    registro = registro & String(15, " ") ' Nosso N� (15*B)
    
    registro = registro & String(5, " ")  ' Campo Fixo
    
    registro = registro & String(8, "0")  ' Data Pagamento (Efetivo) [Arquivo Retorno]
    
    registro = registro & String(15, "0") ' Valor Efetivo [Arquivo Retorno]
    
    registro = registro & String(20, " ") ' Finalidade do Registro/Detalhe (20*B)
    
    registro = registro & String(6, "0")  ' N� do Documento [Informado no Arquivo Retorno]
    
    ' CPF/CNPJ do Favorecido
    ' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    textdata = Format(.Cells(linha, 10).Value, "@") ' Formata como texto [Coluna J]
    registro = registro & CompletaEsquerda(14, textdata) ' Concatena com registro
    ' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    
    registro = registro & String(2, " ")  ' Finalidade DOC (2*B)
    
    registro = registro & String(5, " ")  ' Finalidade TED
    
    registro = registro & String(5, " ")  ' Campo Fixo (5*B)
    
    registro = registro & "0"             ' Aviso (0 para n�o emitir aviso ao favorecido)
    
    registro = registro & String(10, " ") ' Ocorrencia no retorno (10*B)

End With

' Registra o Detalhe
Sheets("Sa�da").Cells(ult_lin_saida, 2).Value = registro

End Sub
