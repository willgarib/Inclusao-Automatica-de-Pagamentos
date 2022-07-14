Attribute VB_Name = "Módulo5"
' Montando o arquivo

Sub geraArquivoDefinido()
''' Gera o Arquivo com dois lotes definidos e Salva '''

' Cabeçalho do Arquivo
Call cabecaArquivo

' Coloca o 1° lote inteiro
If Sheets("Lote Datalhe").Range("B5").Value <> "" Then Call CompilaLote("Lote Datalhe")

' Altera forma de pagamento
Sheets("Lote").Range("C5").Value = "TED – OUTRO TITULAR"

' Recalcula as fórmulas da planilha
Sheets("Lote").Calculate

' Coloca o 2° lote inteiro
If Sheets("Lote Datalhe (2)").Range("B5").Value Then Call CompilaLote("Lote Datalhe (2)")

' Rodapé do Arquivo
Call rodapeArq

' Salva o arquivo
Call salvaArquivio

End Sub
