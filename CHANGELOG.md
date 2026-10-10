# Changelog

Todas as mudanças do painel ficam registradas aqui.

O projeto segue o [versionamento semântico](https://semver.org/lang/pt-BR/) (`MAIOR.MENOR.CORREÇÃO`):

- **MAIOR** (2.0.0): mudança que exige ação de quem usa, como reinstalar ou refazer a configuração.
- **MENOR** (1.1.0): função nova que não quebra nada.
- **CORREÇÃO** (1.0.1): correção de erro ou de segurança.

## [Não lançado]

### Alterado

- O projeto passou a se chamar **mininet**: título, janelas, menu e documentação. O endereço do repositório, a pasta de dados e o atalho de inicialização continuam os mesmos.

### Adicionado

- Testes e análise de segurança (PSScriptAnalyzer) automáticos num Windows do GitHub, a cada envio.
- OpenSSF Scorecard: nota de segurança do repositório, atualizada toda semana.
- Dependabot para manter as ações dos workflows atualizadas.
- Guia de contribuição ([CONTRIBUTING.md](CONTRIBUTING.md)).
- Aviso privado de vulnerabilidades pelo GitHub, com prazo de resposta, no [SECURITY.md](SECURITY.md).

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
