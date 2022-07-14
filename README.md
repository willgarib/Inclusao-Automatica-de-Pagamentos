# Inclusão Automática de Pagamentos
Intercâmbio Eletrônico de Arquivos (SISPAG - ITAÚ)

Rotinas criadas em *Excel-VBA*, a partir do preenchimento de tabelas, criam um arquivo Layout que pode ser carregado no *app* do banco **ITAÚ**. Os pagamentos serão processados e ficarão aguardando autorização.

## Como Utilizar?
- No arquivo SISPAG-ITAÚ (*.xlsm*) vá até a planilha *Listas* e preencha a tabela com as informações da sua empresa.
- Depois, vá até a planilha *Arquivo* e selecione na célula ```C4``` o **CNPJ** que fará os pagamentos, note que a agência e conta são selecionadas automaticamente.
- Na planilha *Lote* selecione o **Tipo de Pagamento**, **Forma de Pagamento** e, novamente, o CNPJ.
> Nota: Verifique a compatibilidade entre os **Tipos** e **Formas** de pagamento no arquivo *SISPAG - SISTEMA DE CONTAS A PAGAR ITAÚ*
- Entre com as informações dos favorecidos nas planilha *Lote Datalhe* e *Lote Datalhe (2)*
  - Em *Lote Datalhe* coloque apenas pagamentos para contas do ITAÚ (cód 341)
  - Em *Lote Datalhe (2)* coloque pagamentos de outros bancos
- Volte para *Arquivo* e clique no botão "Gerar Arquivo com dois Lotes"
- Selecione um arquivo que está na pasta que deseja salvar o **Arquivo Layout** e clique em *Abrir*
- Pronto! Seu arquivo está salvo e já pode ser carregado no *app* do Banco

<sub>Obs: A "aba" *Saída* contém uma cópia do último lote gerado</sub>
