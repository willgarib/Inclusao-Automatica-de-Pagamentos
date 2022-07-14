Attribute VB_Name = "M�dulo1"
' Salvando fora do Excel

Sub salvaArquivio()
    ''' Carrega o rodap� do Arquivo '''

' Dimensionamento das Vari�veis
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Dim ult_lin As Long
ult_lin = Sheets("Sa�da").Cells(100000, 2).End(xlUp).Row + 1

Dim caminho As Variant

Dim linha As Integer
linha = 3
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-

' Verifica se Selecionou Arquivo
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-
caminho = Application.GetOpenFilename("TXT, *.txt", 1, "Salvar em") ' Define o Nome do Arquivo
If Not IsArray(arquivos) Then
    If caminho = False Then ' Se n�o selecionou nada cancela
        MsgBox "Processo Interrompido"
        Exit Sub
    End If
End If
caminho = Replace(caminho, "\" & Dir(caminho), "")
caminho = caminho & "\" & Sheets("Lote").Range("J4").Value & NumeroCPF(Format(Sheets("Arquivo").Range("C9").Value, "dd/mm/yy")) & ".txt"
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*-

' Cria e Abre o Arquivo .txt que ser� a sa�da
Open caminho For Output As #1

Do Until Sheets("Sa�da").Cells(linha, 2).Value = ""

    ' Coloca as Linhas da Planilha "Sa�da" no Arquivo e Salva
    Print #1, Sheets("Sa�da").Cells(linha, 2).Value
    linha = linha + 1

Loop

Close #1

' Mensagem ao usu�rio
MsgBox "Arquivo Salvo em: " & caminho

End Sub
