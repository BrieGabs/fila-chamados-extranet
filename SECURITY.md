# Segurança

O mininet foi feito e revisado com a segurança como prioridade, para que colaboradores e empresa possam usá-lo com tranquilidade:

- **Colaboradores:** a senha nunca passa pelo painel, e os arquivos de dados guardados no computador ficam cifrados com a conta do Windows de cada pessoa. A planilha do histórico só existe enquanto está aberta no Excel.
- **Empresa:** o repositório não contém o endereço da Extranet, dados de clientes nem chamados reais. O painel só conversa com a Extranet e não dá a ninguém acesso além do que o próprio login já permite.

Este documento detalha o que o painel protege, os cuidados que dependem de quem usa e como avisar sobre um problema.

## Versões com suporte

| Versão | Suporte |
|---|---|
| 1.1.x | ✅ Recebe correções de segurança |
| 1.0.x | ❌ Atualize para a 1.1 |

A versão instalada aparece no menu do botão direito do painel ("mininet vX.Y.Z").

## O que o painel protege

### Login e sessão

- **A senha nunca passa pelo painel.** O login é feito na própria página da Extranet, dentro da janela do painel. O painel não lê, não guarda e não envia a senha.
- **A sessão da Extranet** fica no perfil do navegador embutido (WebView2), em `%LOCALAPPDATA%\FilaChamados`. Os cookies são protegidos pelo próprio WebView2 com a criptografia do Windows.

### Endereço da Extranet

- **O endereço não está no código nem no GitHub.** Na primeira vez, o painel pede o endereço, e ele fica guardado só no computador de quem usa, dentro do `estado.json` cifrado.

### Dados no computador

- **Os arquivos `chamados.json`, `estado.json`, `historico.json` e `execucao.json` ficam cifrados** com a conta do Windows (DPAPI). Uma cópia da pasta em outro computador, outra conta, OneDrive ou pendrive mostra só texto ilegível.
- **Os dados gerados nunca vão para o git:** todos estão no `.gitignore`.
- **O `diagnostico-extranet.html`**, uma cópia da página criada quando a leitura falha, é apagado sozinho na próxima leitura que der certo.

### Navegador do painel

- **Só abre a Extranet.** Links para outros sites abrem no navegador padrão. Redirecionamentos para fora são bloqueados, e o painel avisa.
- **Popups** (`target=_blank`, `window.open`) não abrem janelas com a sua sessão: vão para o navegador padrão.
- **A fila só é aceita se vier de uma página da Extranet**, e o código de cada chamado é conferido (só números) antes de virar link.

### Planilha do histórico

- **Só existe enquanto é usada.** A planilha é gerada na hora em que a pessoa clica em **Abrir histórico no Excel**, em `%LOCALAPPDATA%\FilaChamados`, uma pasta do próprio usuário, fora da pasta do painel e do OneDrive. Ela é apagada quando o painel fecha. Se o Excel ainda estiver com ela aberta, é apagada na próxima vez que o painel abrir. O histórico em si fica guardado no `historico.json`, cifrado.
- **Sem fórmulas:** texto que começa com `=`, `+`, `-` ou `@` é gravado como texto, e o Excel não executa nada que venha de um título de chamado.

### Componentes

- **As DLLs do WebView2** (`webview2/`) são as originais do SDK da Microsoft, versão 1.0.4191.47, com assinatura digital da Microsoft Corporation.

## O que o painel não protege (cuidados)

- **Enquanto o Excel está aberto, a planilha do histórico fica legível**, e uma cópia salva pelo **Salvar como** do Excel não é apagada pelo painel. Ela tem nomes de clientes e títulos de chamados: não envie para fora da empresa.
- **`diagnostico-extranet.html` fica aberto enquanto existir.** Ele tem a página inteira da sua fila. Não anexe em issues nem em lugares públicos.
- **Quem entra na sua conta do Windows lê tudo.** A cifragem protege contra cópias da pasta, não contra quem usa a sua própria conta.
- **O script não é assinado digitalmente.** Ele roda com `-ExecutionPolicy Bypass`. Baixe o painel só pelas releases deste repositório e não use cópias recebidas por outros caminhos.
- **Trocar de computador ou de conta do Windows** torna os `.json` ilegíveis. Apague-os, e o painel recomeça do zero, pedindo o endereço de novo.

## Como avisar sobre um problema de segurança

- **Não abra uma issue pública** com detalhes do problema, e nunca anexe a planilha do histórico, o `diagnostico-extranet.html` ou prints com dados de clientes.
- **Use o aviso privado do GitHub:** [Security → Report a vulnerability](https://github.com/BrieGabs/fila-chamados-extranet/security/advisories/new). Só quem mantém o repositório vê o aviso.
- Conte o que aconteceu, a versão do painel e como reproduzir, com dados fictícios.
- **Prazo:** a primeira resposta sai em até 14 dias. Quando o problema for confirmado, a correção é lançada o quanto antes, e quem avisou recebe o crédito na release, se quiser.
- Correções de segurança saem como versão de correção (por exemplo, 1.0.1) e ficam registradas no [CHANGELOG](CHANGELOG.md).
