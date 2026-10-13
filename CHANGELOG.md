# Changelog

Todas as mudanças do painel ficam registradas aqui.

O projeto segue o [versionamento semântico](https://semver.org/lang/pt-BR/) (`MAIOR.MENOR.CORREÇÃO`):

- **MAIOR** (2.0.0): mudança que exige ação de quem usa, como reinstalar ou refazer a configuração.
- **MENOR** (1.1.0): função nova que não quebra nada.
- **CORREÇÃO** (1.0.1): correção de erro ou de segurança.

## [Não lançado]

### Segurança

- O ZIP de cada release passa a ser gerado pelo GitHub a partir do código da tag e assinado com Sigstore, com o registro da procedência (`fila-chamados.zip.sigstore.json`). Para lançar uma versão, basta enviar a tag.
- O resultado da análise de segurança (PSScriptAnalyzer) vai para a aba Security do repositório (Code Scanning).
- O branch `main` passou a ser protegido: mudanças só entram por pull request, com os testes e a análise de segurança aprovados.

### Alterado

- Calendário: "chegou" passa a ser o dia em que você resgatou o chamado ou em que ele foi delegado para você, lido nas interações do chamado. Sem essa interação, continua valendo a data de criação.
- A cada leitura, o painel abre também a página de cada chamado da fila, uma de cada vez e por até 45 segundos, para ler as horas e as interações.

### Adicionado

- Horas DEV estimadas e realizadas de cada chamado, na fila, no histórico e na planilha do histórico. O realizado fica em vermelho quando passa do estimado.
- Relatório de horas acima do estimado, no menu do botão direito: chamados com horas realizadas maiores que as estimadas que passaram por "Aguardando Homologação" no período escolhido, em PDF. O PDF é temporário como a planilha do histórico e é impresso por um navegador à parte, sem a sessão da Extranet e sem JavaScript.

## [1.1.0] - 2026-10-10

### Corrigido

- Fila com mais de 20 chamados: o painel lia só a primeira página da Extranet e recusava a leitura ("a página trouxe 20 de 26 chamados"). Agora percorre todas as páginas. O mecanismo de paginação da Extranet (AJAX) foi identificado por José Ramos.

### Alterado

- Barra de rolagem no tema escuro do painel.
- A lista mantém a posição da rolagem quando o painel se atualiza sozinho.
- O projeto passou a se chamar **mininet**: título, janelas, menu e documentação. O endereço do repositório, a pasta de dados e o atalho de inicialização continuam os mesmos.

### Adicionado

- Filtro por status na fila, com a quantidade de chamados de cada status.
- Andamento da leitura no topo do painel quando a fila tem várias páginas.
- Aviso no histórico quando ele tem mais de 60 chamados, indicando o Excel para ver todos.
- Testes e análise de segurança (PSScriptAnalyzer) automáticos num Windows do GitHub, a cada envio.
- OpenSSF Scorecard: nota de segurança do repositório, atualizada toda semana.
- Dependabot para manter as ações dos workflows atualizadas.
- Guia de contribuição ([CONTRIBUTING.md](CONTRIBUTING.md)).
- Aviso privado de vulnerabilidades pelo GitHub, com prazo de resposta, no [SECURITY.md](SECURITY.md).
- README com a segurança em destaque: resumo no topo, como conferir o SHA-256 do ZIP, tabela do que está protegido e os cuidados de quem usa.

## [1.0.0] - 2026-10-10

Primeira versão segura do painel.

### Segurança

- O endereço da Extranet não fica mais no código. O painel pede o endereço na primeira vez e o guarda cifrado no computador de quem usa.
- Os dados locais (`chamados.json`, `estado.json`, `historico.json`, `execucao.json`) passaram a ser cifrados com a conta do Windows (DPAPI).
- O navegador do painel só abre a Extranet. Links e popups para outros sites vão para o navegador padrão, e redirecionamentos para fora são bloqueados.
- A fila só é aceita quando vem de uma página da Extranet, e o código de cada chamado é conferido antes de virar link.
- A planilha do histórico (`historico.csv`) não fica mais aberta na pasta do painel. Ela é gerada só quando você abre no Excel, numa pasta do próprio usuário, e apagada quando o painel fecha. A `historico.csv` deixada por versões anteriores é apagada sozinha.
- A planilha do histórico passou a ser protegida contra fórmulas do Excel.
- O `diagnostico-extranet.html` passou a ser apagado sozinho depois de uma leitura que dá certo.
- Os exemplos, os testes e as imagens agora usam apenas dados fictícios.

### Adicionado

- Item **Trocar endereço da Extranet** no menu do botão direito.
- Versão do painel exibida no menu do botão direito.
- Documento de segurança ([SECURITY.md](SECURITY.md)).
- Licença MIT ([LICENSE](LICENSE)).

### Removido

- O arquivo `config.json`, com o sobrenome e os endereços: o endereço passou a ser pedido no painel, e o sobrenome não tinha efeito.
- O item de menu "Abrir pasta de relatórios".
