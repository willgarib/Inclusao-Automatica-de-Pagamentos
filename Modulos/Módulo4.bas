Attribute VB_Name = "Módulo4"
' Detalhes do Lote Transferência

Sub CompilaLote(planilha As String)
    ''' Compila e carrega todos os registros do lote que estão na planilha "Lote Detalhe" '''
    
' Dimensionamento das Variáveis
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Dim ult_lin As Long
ult_lin = Sheets(planilha).Cells(100000, 2).End(xlUp).Row

Dim linha As Long
Dim Din As String
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-

Call cabecaLote ' Carrega o cabeçalho do lote

' Executa a formação do registro e carrega para "Saída" cada linha do lote cadastrado em "Lote Detalhe"
For linha = 5 To ult_lin

    Call geraRegistroDetalhe(linha, planilha)
    
Next

' Salva Informações Sobre o Lote na Planilha "Lote" para Composição do Rodapé do Lote
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Sheets("Lote").Range("H7").Value = "'" & CompletaEsquerda(6, ult_lin + 2 - 4) ' Qtd Registros Detalhe desse Lote [H7 da Lote]

Sheets("Lote").Range("I7").Value = Application.Sum(Sheets(planilha).Range("I5:I" & ult_lin))  ' Salva a soma do lote [I7 da Lote] ($)
Din = Format(Sheets("Lote").Range("I7").Value, "@") '
Din = CorrigeDin(Din)                               ' Din tem o valor (str) correto só falta corrigir
Sheets("Lote").Range("I7").Value = "'" & CompletaEsquerda(18, Din) ' Coloca na célula da planilha o valor completado
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-

Call rodapeLote ' Carrega o Rodapé do Lote

' Adiciona 1 ao Código do Lote para o Proximo Lote (Sequêncial por Arquivo)
Sheets("Lote").Range("H4").Value = Sheets("Lote").Range("H4").Value + 1
    
' Salva Informações Sobre o Lote na Planilha "Lote" para Composição do Rodapé do Arquivo
Sheets("Lote").Range("J7").Value = Sheets("Lote").Range("J7").Value + ult_lin + 2 - 4 ' Soma ao Acumulado (Qtd de Registros)

End Sub

Sub geraRegistroDetalhe(linha As Long, planilha As String)
    ''' Gera sequência de caracteres que é um registro de detalhe de lote '''

' Dimensionamento das Variáveis
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Dim ult_lin_saida As Long
ult_lin_saida = Sheets("Saída").Cells(100000, 2).End(xlUp).Row + 1

Dim registro As String
registro = ""

Dim AC As String
AC = ""

Dim textdim As String
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-

With Sheets(planilha)

    registro = registro & "341" ' Banco ITAÚ
    
    registro = registro & Sheets("Lote").Range("I4").Value ' Código do Lote
    
    registro = registro & "3" ' Campo Fixo
    
    registro = registro & CompletaEsquerda(5, linha - 4) ' Sequencial Detalhe
    
    registro = registro & "A" ' Tipo do Registro
    
    registro = registro & "000" ' Tipo de Movimentação [000 para inclusão de pagamento]
    
    registro = registro & "000" ' Câmara
    
    registro = registro & CompletaEsquerda(3, .Cells(linha, 2).Value) ' Banco Favorecido [Coluna B]
    
    ' Definindo a Agência e Conta (AC)
    If .Cells(linha, 2).Value = "341" Or .Cells(linha, 2).Value = "409" Then ' ITAÚ e UNIBANCO
        AC = AC & "0" ' Campo Fixo
        AC = AC & CompletaEsquerda(4, .Cells(linha, 3).Value) 'Agência [Coluna C]
        AC = AC & " " ' Campo Fixo
        AC = AC & String(6, "0") ' Campo Fixo
        AC = AC & CompletaEsquerda(6, .Cells(linha, 5).Value) 'Conta & DAC [Coluna E]
        AC = AC & " " ' Campo Fixo
        AC = AC & CompletaEsquerda(1, .Cells(linha, 6).Value) ' DAC Agência [Coluna F]
    Else:                                                                    ' Demais Bancos
        AC = AC & CompletaEsquerda(5, .Cells(linha, 3).Value & .Cells(linha, 4).Value) 'Agência [Coluna C]
        AC = AC & " " ' Campo Fixo
        AC = AC & CompletaEsquerda(12, .Cells(linha, 5).Value & .Cells(linha, 6).Value) 'Conta e DAC da Conta [Coluna E e F]
        AC = AC & " " ' Campo Fixo
        AC = AC & CompletaDireita(1, .Cells(linha, 4).Value) 'DAC Agência [Coluna D]
    End If
    
    registro = registro & AC ' Concatena AC ao registro
    
    registro = registro & CompletaDireita(30, Left(.Cells(linha, 7).Value, 30)) ' Nome do Favorecido [Coluna G]
    
    registro = registro & CompletaDireita(20, Sheets("Lote").Range("J4").Value) ' Seu Número
    Sheets("Lote").Range("J4").Value = Sheets("Lote").Range("J4").Value + 1     ' Adiciona 1 ao "seu número"
    
    ' Data Pagamento (Previsão) [Coluna H]
    ' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    textdata = Format(.Cells(linha, 8).Value, "@") ' Formata como texto
    textdata = Replace(textdata, "/", "")                                ' Tira a barra
    registro = registro & textdata                                       ' Concatena com registro
    ' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    
    registro = registro & "009" ' Tipo Moeda (009 para REAL)
    
    registro = registro & String(8, " ") ' Câmara 8*B
    
    registro = registro & String(7, "0") ' Campo Fixo
    
    ' Valor Previsto [Coluna I]
    ' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    textdim = Format(.Cells(linha, 9).Value, "@") ' Formata com Texto
    textdim = CorrigeDin(textdim)                                       ' Tira a vírgula e Corrige
    registro = registro & CompletaEsquerda(15, textdim)                 ' Concatena com registro
    ' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    
    registro = registro & String(15, " ") ' Nosso N° (15*B)
    
    registro = registro & String(5, " ")  ' Campo Fixo
    
    registro = registro & String(8, "0")  ' Data Pagamento (Efetivo) [Arquivo Retorno]
    
    registro = registro & String(15, "0") ' Valor Efetivo [Arquivo Retorno]
    
    registro = registro & String(20, " ") ' Finalidade do Registro/Detalhe (20*B)
    
    registro = registro & String(6, "0")  ' N° do Documento [Informado no Arquivo Retorno]
    
    ' CPF/CNPJ do Favorecido
    ' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    textdata = Format(.Cells(linha, 10).Value, "@") ' Formata como texto [Coluna J]
    registro = registro & CompletaEsquerda(14, textdata) ' Concatena com registro
    ' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    
    registro = registro & String(2, " ")  ' Finalidade DOC (2*B)
    
    registro = registro & String(5, " ")  ' Finalidade TED
    
    registro = registro & String(5, " ")  ' Campo Fixo (5*B)
    
    registro = registro & "0"             ' Aviso (0 para não emitir aviso ao favorecido)
    
    registro = registro & String(10, " ") ' Ocorrencia no retorno (10*B)

End With

' Registra o Detalhe
Sheets("Saída").Cells(ult_lin_saida, 2).Value = registro

End Sub
