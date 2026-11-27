# Apps Script da planilha "LP Carreiras V4 Lima & Co"

O script fica na própria planilha (Extensões > Apps Script), no arquivo `Código.gs`.
Esta pasta guarda uma cópia da parte do formulário de vagas, para histórico.

- `formulario-vagas.gs`: tudo de `abrirPainel` até o fim do `Código.gs` (formulário
  Vagas > Nova vaga / Editar vaga selecionada, `salvarVaga` e o HTML do formulário).
  O restante do arquivo (`doGet`, `doPost`, proteções, currículos e o menu `onOpen`)
  não mudou.

O formulário segue o modelo de vaga do time de People. Ao salvar, o script cria
sozinho na aba Vagas as colunas que ainda não existirem:
Endereço, Remuneração, Horário, Benefícios, Outras informações,
O que a posição oferece e Sobre a empresa.

O site lê todas as colunas da aba Vagas pelo `doGet`, então essas mudanças não
exigem nova implantação do app da web.
