<h1 align="center">Mininet</h1>

<p align="center">
  <a href="https://github.com/BrieGabs/fila-chamados-extranet/actions/workflows/testes.yml"><img src="https://github.com/BrieGabs/fila-chamados-extranet/actions/workflows/testes.yml/badge.svg" alt="Testes"></a>
  <a href="https://scorecard.dev/viewer/?uri=github.com/BrieGabs/fila-chamados-extranet"><img src="https://api.scorecard.dev/projects/github.com/BrieGabs/fila-chamados-extranet/badge" alt="OpenSSF Scorecard"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/licen%C3%A7a-MIT-blue" alt="Licença MIT"></a>
  <img src="https://hits.sh/github.com/BrieGabs/fila-chamados-extranet.svg?label=visualiza%C3%A7%C3%B5es&color=007ec6" alt="Visualizações desta página">
</p>

<p align="center">
  Sua fila de chamados da Extranet, sempre visível no canto da tela.
</p>

<p align="center">
  <img src="docs/capa.png" width="900" alt="Painel mininet flutuando no canto inferior direito da área de trabalho, com cinco chamados listados">
</p>

<p align="center">
  <a href="#segurança"><img src="docs/seguranca.svg" width="16" height="16" alt=""> <b>Segurança</b></a> ·
  <a href="#como-instalar"><b>Instalar</b></a> ·
  <a href="#primeiro-uso"><b>Primeiro uso</b></a> ·
  <a href="#conhecendo-o-painel"><b>Conhecendo o painel</b></a> ·
  <a href="#problemas-comuns"><b>Problemas comuns</b></a>
</p>

> **Exclusivo para a Extranet.** O painel só funciona com a Extranet da empresa e exige uma conta com acesso a ela. Não serve para outros sistemas de chamados.

> [!IMPORTANT]
> <img src="docs/seguranca.svg" width="16" height="16" alt=""> **Segurança em primeiro lugar.** O mininet foi feito e revisado com a segurança como prioridade:
>
> - **Sua senha nunca passa pelo painel.** Você entra na própria página da Extranet; o painel não lê, não guarda e não envia a senha.
> - **Nada sai do seu computador.** O painel só conversa com a Extranet, e o navegador dele não abre nenhum outro site.
> - **Tudo o que fica no disco é cifrado** com a sua conta do Windows (DPAPI). Uma cópia da pasta em outro computador, outra conta ou no OneDrive fica ilegível.
> - **O repositório não tem dados da empresa:** nem o endereço da Extranet, nem clientes, nem chamados reais.
> - **Verificado de forma independente:** testes de segurança a cada mudança, análise da Microsoft (PSScriptAnalyzer) e nota pública do [OpenSSF Scorecard](https://scorecard.dev/viewer/?uri=github.com/BrieGabs/fila-chamados-extranet).
>
> Detalhes, os cuidados que dependem de você e como avisar sobre um problema: [Segurança](#segurança) e [SECURITY.md](SECURITY.md).

<sub>As imagens desta página são ilustrações com dados fictícios.</sub>

## O que você ganha

- **Sua fila sempre à vista.** Uma janelinha pequena, por cima das outras, que não ocupa a barra de tarefas.
- **Atualização sozinha.** O painel lê a Extranet a cada 15 minutos enquanto a sua sessão estiver aberta.
- **Um clique para abrir.** Clique em um chamado e ele abre na Extranet.
- **Histórico e calendário.** Veja o que saiu da sua fila, o que chegou em cada dia e em quais dias cada chamado esteve em execução.
- <img src="docs/seguranca.svg" width="16" height="16" alt=""> **Sem senha guardada.** Você entra na própria página da Extranet. O painel não salva sua senha.
- <img src="docs/seguranca.svg" width="16" height="16" alt=""> **Dados cifrados.** O que o painel guarda no seu computador fica cifrado com a sua conta do Windows. Veja [Segurança](#segurança).

## Antes de começar

Você precisa de:

- Windows 10 ou 11
- Uma conta na Extranet

Não é preciso instalar nenhum programa. O painel usa o PowerShell e o WebView2, que já vêm no Windows.

## Como instalar

<p align="center">
  <img src="docs/instalar.png" width="860" alt="Três passos: baixar o ZIP, desbloquear e extrair, abrir o painel">
</p>

### 1. Baixe o ZIP

[**Baixar o ZIP da versão mais recente**](https://github.com/BrieGabs/fila-chamados-extranet/releases/latest/download/fila-chamados.zip)

O link sempre baixa a última versão publicada. As versões anteriores e o que mudou em cada uma estão em [Releases](https://github.com/BrieGabs/fila-chamados-extranet/releases) e no [CHANGELOG](CHANGELOG.md).

#### Versões disponíveis

| Versão | Data | Destaques | Assinada | SHA-256 do ZIP |
|---|---|---|---|---|
| [**v1.1.0**](https://github.com/BrieGabs/fila-chamados-extranet/releases/tag/v1.1.0) (mais recente) | 10/10/2026 | Lê filas com mais de 20 chamados, filtro por status e andamento da leitura. O projeto passa a se chamar mininet. | Sim | `b4d54f9adacdf054d9d96a9b8daf06f0fdeaf1b523b9172f5cc1684cf0aa824a` |
| [v1.0.0](https://github.com/BrieGabs/fila-chamados-extranet/releases/tag/v1.0.0) | 10/10/2026 | Primeira versão segura: endereço da Extranet e dados locais cifrados, navegador restrito à Extranet. | Não | `e002a491fa285fad2dcb42612e8005592da3838525593922b2ca0ba690e71070` |

> [!WARNING]
> **Baixe só por aqui.** O script não é assinado digitalmente e roda com `-ExecutionPolicy Bypass`. Não use cópias recebidas por e-mail, chat, pendrive ou pasta compartilhada: elas podem ter sido alteradas.

<details>
<summary><b>Conferir que o ZIP é o original (recomendado)</b></summary>

Cada release traz a impressão digital SHA-256 do ZIP. No PowerShell, rode:

```
Get-FileHash "$env:USERPROFILE\Downloads\fila-chamados.zip" -Algorithm SHA256
```

O valor mostrado precisa ser igual ao publicado na [página da release](https://github.com/BrieGabs/fila-chamados-extranet/releases/latest). Se for diferente, apague o ZIP e não abra o painel.

O ZIP também é **assinado** pelo GitHub (Sigstore), com o registro de que foi gerado a partir do código deste repositório. Quem tem o [GitHub CLI](https://cli.github.com/) pode conferir a assinatura:

```
gh attestation verify fila-chamados.zip -R BrieGabs/fila-chamados-extranet
```
</details>

### 2. Desbloqueie e extraia

1. Abra a pasta **Downloads**.
2. Clique com o botão direito no ZIP e escolha **Propriedades**.
3. Na parte de baixo da janela, marque **Desbloquear** e clique em **OK**.
4. Clique com o botão direito no ZIP de novo e escolha **Extrair tudo**.
5. Extraia para uma pasta que você não vá apagar, por exemplo `Documentos\fila-chamados`.

> **Por que desbloquear?** O Windows marca arquivos baixados da internet e costuma barrar os componentes do painel. Desbloquear o ZIP antes de extrair evita isso. Se a opção **Desbloquear** não aparecer, pode seguir adiante.

### 3. Abra o painel

Na pasta extraída, dê dois cliques em **`Abrir Fila de Chamados.cmd`**.

O painel aparece no canto inferior direito da tela.

<details>
<summary>Prefere usar o Git?</summary>

```
git clone https://github.com/BrieGabs/fila-chamados-extranet.git
```

Depois, dê dois cliques em `Abrir Fila de Chamados.cmd`. Com o Git não é preciso desbloquear nada.
</details>

## Primeiro uso

Na primeira vez, o painel ainda não tem a sua fila nem o endereço da Extranet.

<p align="center">
  <img src="docs/primeiro-uso.png" width="414" alt="Painel vazio com um aviso e o botão Recarregar destacado">
</p>

1. Clique em **Recarregar** (a seta circular).
2. O painel pede o endereço da Extranet. Abra a Extranet no seu navegador, copie o endereço da barra (pode ser a página de login), cole e clique em **Salvar**.
3. Uma janela da Extranet abre. Entre com seu usuário e senha, como você já faz no navegador.
4. A janela fecha sozinha e a sua fila aparece no painel.

O endereço é pedido só uma vez. Ele não vem no painel nem vai para o GitHub: fica guardado cifrado neste computador. Para trocá-lo, clique com o botão direito no painel e escolha **Trocar endereço da Extranet**.

Pronto. Daqui em diante o painel se atualiza sozinho a cada 15 minutos. Quando a sessão da Extranet expirar, ele avisa e basta clicar em **Recarregar** para entrar de novo.

> **Dica:** para o painel abrir junto com o Windows, clique nele com o botão direito e marque **Iniciar com o Windows**.

## Conhecendo o painel

<p align="center">
  <img src="docs/tour.png" width="470" alt="Painel com nove marcadores numerados sobre os elementos da tela">
</p>

| Nº | O que é | Para que serve |
|:---:|---|---|
| 1 | Contador | Quantos chamados estão na sua fila agora |
| 2 | Última leitura | Quando a fila foi lida da Extranet pela última vez |
| 3 | Calendário | Abre o calendário e o Gantt. Clique de novo para voltar à fila |
| 4 | Histórico | Mostra os chamados que saíram da sua fila. Clique de novo para voltar |
| 5 | Recarregar | Lê a fila da Extranet agora |
| 6 | Recolher | Encolhe o painel para só o cabeçalho, e expande de volta |
| 7 | Fechar | Fecha o painel |
| 8 | Chamado | Clique para abrir na Extranet. Passe o mouse para ver cliente, versão e data prevista |
| 9 | Dias em aberto | Há quantos dias o chamado foi criado |

**Para mover o painel**, arraste pelo cabeçalho. Ele lembra a posição.

### Como ler um chamado

Cada chamado mostra o código, as etiquetas, o título, o cliente e o status atual.

<p align="center">
  <img src="docs/etiquetas.png" width="572" alt="Legenda das etiquetas: Incidência, Novo requisito, Consultoria, em execução, NOVO e atrasado">
</p>

- O status fica **amarelo** quando mudou hoje ou ontem.
- A lista traz primeiro o que está **em execução**, depois as **incidências**, depois o que mudou de status há menos tempo.

### Fila grande

- **Todas as páginas são lidas.** A Extranet mostra a fila de 20 em 20; o painel percorre as páginas sozinho. Enquanto lê, o topo mostra o andamento, por exemplo "atualizando… 40 de 85".
- **Filtro por status.** Quando a fila tem mais de um status, aparecem botões acima da lista, um por status, com a quantidade, por exemplo "Aguardando Análise (12)". Clique para ver só aquele status; clique em **Todos** para voltar. O contador do topo mostra "12 de 85" e o painel lembra o filtro escolhido.
- **A lista não volta para o topo** quando o painel se atualiza sozinho: você continua de onde estava.

### Histórico

Quando um chamado sai da sua fila, ele passa para o histórico, com o dia em que saiu e o último status que tinha com você.

<p align="center">
  <img src="docs/historico.png" width="414" alt="Tela de histórico com quatro chamados que saíram da fila">
</p>

Para ver o histórico no Excel, clique com o botão direito no painel e escolha **Abrir histórico no Excel**. A planilha é gerada na hora em que você abre, numa pasta só sua, e apagada quando você fecha o painel. Se quiser guardar uma cópia, use **Salvar como** no Excel.

### Calendário e Gantt

<p align="center">
  <img src="docs/calendario.png" width="380" alt="Calendário de outubro com barras azuis e verdes nos dias com movimento">
  &nbsp;
  <img src="docs/gantt.png" width="380" alt="Gantt com barras nos dias em que cada chamado esteve em execução">
</p>

**Calendário** mostra o movimento de cada dia do mês.

- Barra **azul**: chamados que chegaram (pela data de criação).
- Barra **verde**: chamados entregues (o dia em que saíram da sua fila).
- Clique em um dia para ver quais foram. Clique no nome do mês para voltar ao mês atual.

**Gantt** mostra em quais dias cada chamado esteve em execução.

- Barra **azul**: o chamado ainda está na sua fila.
- Barra **verde**: o chamado já foi entregue.
- A linha vertical marca o dia de hoje.

Nas duas telas, as etiquetas **Todos**, **Incidência**, **Novo requisito** e **Consultoria** filtram por tipo.

### Menu do botão direito

Clique com o botão direito em qualquer ponto do painel.

| Opção | O que faz |
|---|---|
| **Copiar código** | Copia o código do chamado em que você clicou |
| **Copiar código e título** | Copia código e título, prontos para colar em uma mensagem |
| **Abrir minha fila na Extranet** | Abre a lista completa no seu navegador |
| **Abrir histórico no Excel** | Gera a planilha do histórico e abre no Excel |
| **Recarregar da Extranet** | O mesmo que o botão Recarregar |
| **Trocar endereço da Extranet** | Pede o endereço da Extranet de novo |
| **Voltar para o canto da tela** | Devolve o painel ao canto inferior direito |
| **Iniciar com o Windows** | Liga ou desliga a abertura automática |
| **mininet vX.Y.Z** | Mostra a versão instalada (não é clicável) |
| **Fechar** | Fecha o painel |

As duas opções de copiar só aparecem quando o clique foi em cima de um chamado.

## Problemas comuns

<details>
<summary><b>Dei dois cliques e nada aconteceu</b></summary>

O painel provavelmente já está aberto: ele só abre uma vez. Procure no canto inferior direito da tela. Se você o moveu e não encontra, feche-o pelo Gerenciador de Tarefas (processo **Windows PowerShell**) e abra de novo.
</details>

<details>
<summary><b>O Windows mostra um aviso de segurança, ou o painel diz que não conseguiu abrir o navegador</b></summary>

O ZIP provavelmente foi extraído sem desbloquear. Apague a pasta extraída e refaça o [passo 2](#2-desbloqueie-e-extraia), marcando **Desbloquear** antes de extrair.
</details>

<details>
<summary><b>"Falta o endereço da Extranet"</b></summary>

O painel ainda não sabe o endereço da Extranet. Clique em **Recarregar** e cole o endereço que aparece no seu navegador, começando com `https://`.
</details>

<details>
<summary><b>"A Extranet não abriu a lista de chamados"</b></summary>

O endereço informado provavelmente não é o da Extranet. Clique com o botão direito no painel, escolha **Trocar endereço da Extranet** e cole o endereço da página de login da Extranet.
</details>

<details>
<summary><b>"Sessão da Extranet fechada: clique em Recarregar para entrar."</b></summary>

Sua sessão expirou. Clique em **Recarregar** e entre de novo. Enquanto isso o painel continua mostrando a última leitura.
</details>

<details>
<summary><b>"A fila ainda não foi atualizada hoje."</b></summary>

O painel ainda não conseguiu ler a Extranet hoje. Clique em **Recarregar**.
</details>

<details>
<summary><b>"Falta a pasta webview2 com o navegador do painel"</b></summary>

A pasta `webview2` não veio junto. Extraia o ZIP inteiro de novo, sem tirar arquivos de dentro da pasta.
</details>

<details>
<summary><b>"Não consegui ler a fila na Extranet"</b></summary>

Se a mensagem fala em **páginas** (por exemplo "as páginas trouxeram 38 de 40 chamados" ou "a página 3 da lista não carregou"), a fila provavelmente mudou enquanto o painel lia, ou a Extranet demorou. Clique em **Recarregar** de novo.

A página da Extranet veio diferente do que o painel espera. Tente **Recarregar** mais uma vez. Se continuar, abra uma [issue](https://github.com/BrieGabs/fila-chamados-extranet/issues) com a mensagem que apareceu.

O painel salva uma cópia da página em `diagnostico-extranet.html` e a apaga sozinho na próxima leitura que der certo. **Não anexe esse arquivo em lugares públicos:** ele contém os seus chamados.
</details>

<details>
<summary><b>Mudei a pasta de lugar e o painel parou de abrir com o Windows</b></summary>

O atalho de inicialização aponta para a pasta antiga. Abra o painel pela pasta nova, clique com o botão direito e desmarque e marque de novo **Iniciar com o Windows**.
</details>

<details>
<summary><b>Como desinstalar</b></summary>

1. Botão direito no painel, desmarque **Iniciar com o Windows**.
2. Feche o painel.
3. Apague a pasta do painel.
4. Para apagar também a sessão da Extranet guardada pelo navegador do painel, apague a pasta `%LOCALAPPDATA%\FilaChamados`.
</details>

## Segurança

O painel foi feito e revisado com a segurança como prioridade, para ser usado com tranquilidade pelos colaboradores e pela empresa. A versão 1.0.0 é a primeira com todas as proteções abaixo.

### Para os colaboradores

- **Sua senha fica só com você.** O painel não lê, não guarda e não envia a senha.
- **Seus dados ficam no seu computador.** Nada é enviado para outro lugar, e os arquivos de dados do painel ficam cifrados com a sua conta do Windows.

### Para a empresa

- **O repositório não tem dados da empresa nem de clientes:** nem o endereço da Extranet, nem nomes de clientes, nem chamados reais. Exemplos e imagens usam dados fictícios.
- **O painel só conversa com a Extranet,** e o navegador dele não abre outros sites.
- **O acesso continua sendo o da Extranet:** o painel só mostra o que a pessoa já vê com o próprio login. Ele não amplia permissões nem contorna nenhum controle da Extranet.

### O que está protegido

| Item | Proteção |
|---|---|
| **Senha** | Nunca passa pelo painel. Você entra na própria página da Extranet. |
| **Sessão** | Fica no navegador embutido (WebView2), em `%LOCALAPPDATA%\FilaChamados`, com os cookies protegidos pela criptografia do Windows. |
| **Arquivos de dados** | `chamados.json`, `estado.json`, `execucao.json` e `historico.json` ficam cifrados com a sua conta do Windows (DPAPI). |
| **Endereço da Extranet** | Não está no código nem no GitHub. É pedido no primeiro uso e fica só no seu computador, cifrado. |
| **Navegador do painel** | Só abre a Extranet. Links, popups e redirecionamentos para outros sites são bloqueados ou vão para o seu navegador padrão, sem a sua sessão. |
| **Leitura da fila** | Só é aceita se vier de uma página da Extranet, e o código de cada chamado é conferido antes de virar link. |
| **Planilha do histórico** | Só existe enquanto você usa: é gerada numa pasta só sua e apagada quando o painel fecha. Sai sem fórmulas, então o Excel não executa nada que venha de um chamado. |
| **Componentes** | As DLLs do WebView2 são as originais da Microsoft, com assinatura digital. |

### Seus cuidados

> [!CAUTION]
> Parte da segurança depende de quem usa. Fique atento a estes pontos:
>
> - **Baixe só pelas [releases](https://github.com/BrieGabs/fila-chamados-extranet/releases) deste repositório** e, de preferência, confira o SHA-256. O script não é assinado digitalmente.
> - **Não envie a planilha do histórico para fora da empresa,** nem uma cópia salva pelo **Salvar como**: ela tem nomes de clientes e títulos de chamados, e o painel não apaga cópias salvas.
> - **Não publique nem anexe** `diagnostico-extranet.html`, os `.json` da pasta ou prints com dados reais, nem em issues.
> - **Bloqueie o computador ao sair** (<kbd>Win</kbd> + <kbd>L</kbd>). A cifragem protege contra cópias da pasta, não contra quem usa a sua própria conta do Windows.

### Verificado automaticamente

Os selos de testes e do Scorecard, no topo desta página, não são enfeite: cada um aponta para uma verificação que qualquer pessoa pode abrir.

- **Testes:** a cada mudança, um Windows do GitHub roda o autoteste (incluindo os testes de segurança) e o PSScriptAnalyzer, o analisador de código da Microsoft, com as regras de segurança. O resultado da análise vai para a aba Security do repositório.
- **`main` protegido:** nenhuma mudança entra direto. Tudo passa por pull request, e só é aceito com os testes e a análise de segurança aprovados. O histórico não pode ser reescrito nem apagado.
- **OpenSSF Scorecard:** a Open Source Security Foundation avalia o repositório toda semana e publica uma nota de 0 a 10. Clique no selo para ver cada item avaliado.
- **Releases conferíveis e assinadas:** cada release traz a impressão digital SHA-256 do ZIP e uma assinatura Sigstore, que comprova que o ZIP foi gerado a partir do código deste repositório (veja [como conferir](#1-baixe-o-zip)).

### Encontrou um problema de segurança?

> [!IMPORTANT]
> **Não abra uma issue pública.** Use o aviso privado do GitHub: [Security → Report a vulnerability](https://github.com/BrieGabs/fila-chamados-extranet/security/advisories/new). Só quem mantém o repositório vê o aviso, e a primeira resposta sai em até 14 dias. Detalhes em [SECURITY.md](SECURITY.md).

### Seus dados

- O painel grava seus dados **na própria pasta**: `chamados.json`, `estado.json`, `execucao.json` e `historico.json`.
- Nada é enviado para outro lugar. O painel só conversa com a Extranet.
- Esses arquivos contêm nomes de clientes e títulos de chamados. Eles já estão no `.gitignore`; não os publique.
- **Os `.json` ficam cifrados com a sua conta do Windows.** Se a pasta for copiada para outro computador, aberta por outro usuário ou sincronizada no OneDrive, eles ficam ilegíveis. Só a sua conta, no seu computador, consegue ler.
- A planilha do histórico não fica na pasta do painel. Ela é gerada na hora em que você abre, numa pasta só sua, e apagada quando você fecha o painel. Se o Excel ainda estiver aberto nessa hora, ela é apagada na próxima vez que o painel abrir.
- O endereço da Extranet também fica no `estado.json`, cifrado. Ele não aparece no código nem no GitHub.
- Se você trocar de computador ou de conta do Windows, os `.json` antigos não abrem na conta nova. Apague `chamados.json`, `estado.json`, `execucao.json` e `historico.json` da pasta: o painel recomeça do zero e pede o endereço de novo.
- Sua senha não é salva pelo painel. A sessão da Extranet fica guardada pelo navegador embutido em `%LOCALAPPDATA%\FilaChamados`.

## Para quem quer mexer no código

| Arquivo | O que é |
|---|---|
| `FilaChamados.ps1` | O painel (WPF + WebView2) |
| `Abrir Fila de Chamados.cmd` | Abre o painel sem janela de console |
| `Testar.ps1` | Autoteste das funções (datas, calendário, leitor da lista, segurança); não abre o painel |
| `SECURITY.md` | O que o painel protege, o que não protege e como avisar sobre problemas |
| `CHANGELOG.md` | O que mudou em cada versão |
| `LICENSE` | Licença MIT |
| `CONTRIBUTING.md` | Como relatar erros e enviar mudanças |
| `.github/` | Testes automáticos, assinatura das releases, OpenSSF Scorecard e Dependabot |
| `FilaChamados.ico` | Ícone |
| `webview2/` | DLLs do SDK do Microsoft WebView2 (versão 1.0.4191.47), com a licença da Microsoft |
| `docs/` | Imagens desta página |

Para rodar o autoteste:

```
powershell -NoProfile -ExecutionPolicy Bypass -File Testar.ps1
```

Limitações conhecidas:

- O leitor depende da estrutura da página "Minha responsabilidade" da Extranet. Se ela mudar, o painel mostra a última leitura até o leitor ser ajustado.
- O Gantt só conta os dias em que uma leitura viu o chamado "em execução". Execuções curtas entre duas leituras, ou com o painel fechado, ficam de fora.
- As colunas "Status final" e "Responsável atual" da planilha do histórico ficam em branco: o painel registra quando o chamado saiu da fila, mas não consulta o chamado depois disso.

### Como lançar uma versão

O projeto usa [versionamento semântico](https://semver.org/lang/pt-BR/): `MAIOR.MENOR.CORREÇÃO`. Veja no [CHANGELOG](CHANGELOG.md) quando subir cada número.

O `main` é protegido: a versão entra por pull request, como qualquer outra mudança.

1. Numa branch, atualize `$Versao` no começo do `FilaChamados.ps1`.
2. Registre as mudanças no `CHANGELOG.md`, numa seção nova com a versão e a data.
3. Rode o `Testar.ps1` no Windows e confira se termina com `OK`.
4. Faça o commit, envie a branch e abra o pull request. Faça o merge quando os testes passarem.
5. Atualize o `main` local, crie a tag e envie:

   ```
   git tag -a vX.Y.Z -m "Versão X.Y.Z"
   git push origin vX.Y.Z
   ```

6. Pronto: o workflow **Release** confere se `$Versao` e o `CHANGELOG.md` batem com a tag, gera o `fila-chamados.zip` a partir do código da tag, assina com Sigstore e publica a release com as notas do CHANGELOG, o SHA-256 e a assinatura (`fila-chamados.zip.sigstore.json`). Acompanhe em [Actions](https://github.com/BrieGabs/fila-chamados-extranet/actions/workflows/release.yml).

7. Com a release publicada, acrescente a versão nova no topo da tabela [Versões disponíveis](#versões-disponíveis), com o SHA-256 da página da release, e mande por pull request.

Para publicar de novo uma versão já lançada (por exemplo, se o workflow falhou), rode o **Release** manualmente em Actions, informando a tag.

## Licença

Código aberto sob a [licença MIT](LICENSE). As DLLs da pasta `webview2/` são da Microsoft e seguem a licença dela, em [`webview2/LICENSE.txt`](webview2/LICENSE.txt).
