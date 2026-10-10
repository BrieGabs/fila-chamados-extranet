# Como contribuir

Obrigado por querer melhorar o mininet.

## Relatar um erro ou sugerir melhoria

- Abra uma [issue](https://github.com/BrieGabs/fila-chamados-extranet/issues) e conte:
  - o que aconteceu;
  - o que você esperava que acontecesse;
  - a versão do painel, que aparece no menu do botão direito ("mininet vX.Y.Z").
- **Nunca anexe dados reais:** nada de `historico.csv`, `diagnostico-extranet.html`, prints com clientes ou o endereço da Extranet. Use dados fictícios.
- **Problemas de segurança não vão em issue.** Siga o [SECURITY.md](SECURITY.md).

## Enviar uma mudança

1. Faça um fork e crie uma branch a partir da `main`.
2. Faça a mudança seguindo o estilo do código ao redor:
   - nomes e comentários em português, sem acentos nos comentários do `.ps1`;
   - arquivos `.ps1` em UTF-8 com BOM e quebras de linha CRLF.
3. **Toda função nova ou alterada precisa de teste no `Testar.ps1`.** Mudanças de segurança precisam de um teste que falharia sem a correção.
4. Rode o autoteste no Windows e confira se termina com `OK`:

   ```
   powershell -NoProfile -ExecutionPolicy Bypass -File Testar.ps1
   ```

5. Registre a mudança no [CHANGELOG.md](CHANGELOG.md), na seção "Não lançado".
6. Abra um pull request explicando o que mudou e por quê.

Os testes e a análise de segurança (PSScriptAnalyzer) rodam automaticamente no pull request, num Windows do GitHub. O pull request só é aceito se os dois passarem.

## O que não é aceito

- **Dados da empresa ou de clientes** no código, nos testes, nas imagens ou nos commits. Isso inclui o endereço da Extranet, nomes de clientes e chamados reais.
- **Envio de dados para fora da Extranet,** como telemetria ou serviços externos.
- **Senhas ou credenciais,** em qualquer forma.
- **Desativar ou enfraquecer uma proteção** descrita no [SECURITY.md](SECURITY.md) sem discutir antes numa issue.

## Versões

O projeto segue o versionamento semântico. Quem mantém o repositório decide quando lançar uma versão; o passo a passo está no README, em "Como lançar uma versão".
