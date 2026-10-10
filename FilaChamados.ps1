# mininet - janelinha flutuante com a fila de chamados da Extranet
# Mostra os chamados de chamados.json (mesma pasta). O Recarregar (e a busca
# automatica a cada 15 min) le a fila na Extranet por um navegador WebView2 do
# proprio painel: o login e feito pela pessoa na janelinha que abre, sem senha
# salva. Quando um chamado sai da fila, ele
# vai para o historico (historico.json; a planilha para o Excel e gerada so quando e aberta).
# Os dias em que a leitura ve um chamado "em execucao" ficam em execucao.json (Gantt).
# So mostra chamados que passaram pela sua fila.
# Os .json de dados ficam cifrados com a conta do Windows (DPAPI).
# O endereco da Extranet e pedido no primeiro Recarregar e fica no estado.json (cifrado); nao vem no codigo.
# Clique num chamado para abrir no navegador. Botao direito: mais opcoes.

$ErrorActionPreference = 'Stop'
$Versao = '1.0.0'   # versionamento semantico; mudou aqui, atualize o CHANGELOG.md e crie a tag vX.Y.Z
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$ScriptPath    = $MyInvocation.MyCommand.Path
$Dir           = Split-Path -Parent $ScriptPath
$DataFile      = Join-Path $Dir 'chamados.json'
$StateFile     = Join-Path $Dir 'estado.json'
$HistFile      = Join-Path $Dir 'historico.json'
$HistCsv       = Join-Path $env:LOCALAPPDATA 'FilaChamados\historico.csv'   # temporaria: so existe enquanto o Excel usa
$HistCsvAntigo = Join-Path $Dir 'historico.csv'
$ExecFile      = Join-Path $Dir 'execucao.json'
$WebDir        = Join-Path $Dir 'webview2'
$WebPerfil     = Join-Path $env:LOCALAPPDATA 'FilaChamados\WebView2'
$DiagFile      = Join-Path $Dir 'diagnostico-extranet.html'
$StartupLnk    = Join-Path ([Environment]::GetFolderPath('Startup')) 'Fila de Chamados.lnk'

# Uma instancia so
$created = $false
$mutex = New-Object System.Threading.Mutex($true, 'Local\FilaChamadosWidget', [ref]$created)
if (-not $created) { exit }

# Le .json aberto (arquivos de versoes antigas) ou cifrado por Write-Json
function Read-Json($path) {
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try {
        $txt = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        if ($txt.StartsWith('DPAPI:')) {
            Add-Type -AssemblyName System.Security
            $bytes = [Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String($txt.Substring(6).Trim()), $null, 'CurrentUser')
            $txt = [Text.Encoding]::UTF8.GetString($bytes)
        }
        return ($txt | ConvertFrom-Json)
    } catch { return $null }
}

# why: os .json tem nomes de clientes e titulos; cifrados com a conta do Windows (DPAPI), uma copia da pasta
# (outro usuario, OneDrive, pendrive) fica ilegivel. So a mesma conta, na mesma maquina, consegue ler
function Write-Json($obj, [int]$depth, [string]$path) {
    Add-Type -AssemblyName System.Security
    $bytes = [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $obj -Depth $depth))
    $prot = [Security.Cryptography.ProtectedData]::Protect($bytes, $null, 'CurrentUser')
    Set-Content -LiteralPath $path -Value ('DPAPI:' + [Convert]::ToBase64String($prot)) -Encoding ASCII
}

# why: o endereco da Extranet nao vem no codigo; a pessoa cola no painel (Pedir-Extranet) e ele fica no estado.json, cifrado
$ListaUrl = ''
$ChamadoUrl = ''
$ExtranetDominio = ''   # dominio da lista (sem o www): o unico site que entra no navegador do painel

# why: o navegador do painel guarda a sessao da Extranet e repassa as mensagens da pagina; nenhum outro site entra nele
function Test-UrlExtranet([string]$u) {
    $uri = $null
    if (-not $ExtranetDominio) { return $false }
    return [Uri]::TryCreate($u, [UriKind]::Absolute, [ref]$uri) -and $uri.Scheme -eq 'https' -and $uri.Host -match ('(^|\.)' + [regex]::Escape($ExtranetDominio) + '$')
}

# endereco colado (qualquer pagina da Extranet, ex.: a de login) -> endereco da lista "Minha responsabilidade", na mesma pasta do site
function Get-ListaUrl([string]$texto) {
    $uri = $null
    if (-not [Uri]::TryCreate(([string]$texto).Trim(), [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https') { return $null }
    if ($uri.AbsolutePath -notmatch '/$' -and $uri.Segments[-1] -notmatch '\.') { $uri = New-Object System.Uri ($uri.GetLeftPart('Path') + '/') }   # ".../pasta" sem barra final
    return (New-Object System.Uri ($uri, 'Chamados.aspx?TipoBusca=lnkResponsabilidade')).AbsoluteUri
}

function Set-Extranet([string]$lista) {
    $script:ListaUrl = ''; $script:ChamadoUrl = ''; $script:ExtranetDominio = ''
    $uri = $null
    if (-not [Uri]::TryCreate($lista, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https') { return }
    $script:ListaUrl = $uri.AbsoluteUri
    $script:ChamadoUrl = (New-Object System.Uri ($uri, 'Chamado.aspx?CODIGO=')).AbsoluteUri
    $script:ExtranetDominio = $uri.Host -replace '^www\.', ''
}

$state = Read-Json $StateFile
if (-not $state) { $state = New-Object psobject }
foreach ($p in @('left', 'top', 'recolhido', 'modo', 'ultimoAtualizadoEm', 'ultimos', 'ultimosCodigos', 'codigosBase', 'extranet')) {
    if ($state.PSObject.Properties.Name -notcontains $p) { $state | Add-Member -NotePropertyName $p -NotePropertyValue $null }
}
Set-Extranet ([string]$state.extranet)
function Save-State {
    try { Write-Json $state 5 $StateFile } catch {}
}

function Get-DataLocal($v) {
    if ($v -is [datetime]) { return $v.ToLocalTime() }
    if (-not $v) { return $null }
    try { return [datetime]::Parse([string]$v, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind).ToLocalTime() } catch { return $null }
}

function Get-Stamp($v) {
    if ($v -is [datetime]) { return $v.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ') }
    return [string]$v
}

# ---------------------------------------------------------------- historico
function Get-Historico {
    $h = Read-Json $HistFile
    $list = New-Object System.Collections.ArrayList
    if ($h) { foreach ($e in @($h)) { if ($e -and $e.codigo) { [void]$list.Add($e) } } }
    return , $list
}

# hazard: o Excel roda como formula o texto que comeca com = + - @ (titulo e cliente vem da Extranet, escritos por clientes)
function Get-TextoCsv([string]$s) {
    if ($s -match '^[=+\-@\t\r]') { return "'" + $s }
    return $s
}

function Save-Historico($list) {
    $arr = @($list | Sort-Object { Get-Stamp $_.saiuEm } -Descending)
    try { Write-Json $arr 4 $HistFile } catch {}
}

function Export-HistoricoCsv($arr, [string]$path) {
    $arr | Select-Object `
            @{ n = 'Código'; e = { Get-TextoCsv $_.codigo } },
            @{ n = 'Fora da fila desde'; e = { $d = Get-DataLocal $_.saiuEm; if ($d) { $d.ToString('dd/MM/yyyy') } else { '' } } },
            @{ n = 'Status final'; e = { Get-TextoCsv $_.statusFinal } },
            @{ n = 'Responsável atual'; e = { Get-TextoCsv $_.responsavelAtual } },
            @{ n = 'Último status na minha fila'; e = { Get-TextoCsv $_.ultimoStatus } },
            @{ n = 'Tipo'; e = { Get-TextoCsv $_.tipo } },
            @{ n = 'Cliente'; e = { Get-TextoCsv $_.cliente } },
            @{ n = 'Título'; e = { Get-TextoCsv $_.titulo } },
            @{ n = 'Criado em'; e = { Get-TextoCsv $_.criacao } },
            @{ n = 'Dias em aberto'; e = { $_.diasAberto } },
            @{ n = 'Link'; e = { $_.url } } |
            Export-Csv -LiteralPath $path -NoTypeInformation -Delimiter ';' -Encoding UTF8
}

# why: a planilha aberta so existe enquanto e usada; o historico fica guardado no historico.json, cifrado.
# Ela e gerada em %LOCALAPPDATA% (fora da pasta do painel e do OneDrive) e apagada quando o painel fecha
function Open-HistoricoExcel {
    $h = Get-Historico
    $hist = @($h | Sort-Object { Get-Stamp $_.saiuEm } -Descending)
    if ($hist.Count -eq 0) {
        [void][System.Windows.MessageBox]::Show('O histórico ainda está vazio. Ele começa a encher quando um chamado sair da sua fila.', 'mininet')
        return
    }
    try {
        [void](New-Item -ItemType Directory -Force -Path (Split-Path -Parent $HistCsv))
        Export-HistoricoCsv $hist $HistCsv
    } catch {
        # hazard: o Excel trava a planilha enquanto ela esta aberta; ai so traz de volta a que ja esta aberta
        if (-not (Test-Path -LiteralPath $HistCsv)) {
            [void][System.Windows.MessageBox]::Show("Não consegui gerar a planilha do histórico:`n$($_.Exception.Message)", 'mininet')
            return
        }
    }
    Start-Process $HistCsv
}

# apaga a planilha temporaria e a historico.csv que versoes antes da 1.0.0 deixavam aberta na pasta do painel
function Remove-PlanilhaHistorico {
    foreach ($p in @($HistCsv, $HistCsvAntigo)) { Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue }
}

# prev = fila anterior, atual = fila nova, saidosData = saidos com status final (vindos da verificacao)
function Update-Historico($prev, $atual, $saidosData, [string]$stamp) {
    $hist = Get-Historico
    $codAtual = @($atual | ForEach-Object { [string]$_.codigo })
    $mudou = $false
    # chamado que voltou para a fila sai do historico
    for ($i = $hist.Count - 1; $i -ge 0; $i--) {
        if ($codAtual -contains [string]$hist[$i].codigo) { $hist.RemoveAt($i); $mudou = $true }
    }
    $saidos = [ordered]@{}
    foreach ($s in @($saidosData)) { if ($s -and $s.codigo) { $saidos[[string]$s.codigo] = $s } }
    foreach ($p in @($prev)) {
        $k = [string]$p.codigo
        if ($k -and ($codAtual -notcontains $k) -and -not $saidos.Contains($k)) { $saidos[$k] = $p }
    }
    foreach ($k in @($saidos.Keys)) {
        $s = $saidos[$k]
        $det = @($prev | Where-Object { [string]$_.codigo -eq $k }) | Select-Object -First 1
        if (-not $det -or -not $det.titulo) { $det = $s }
        for ($i = $hist.Count - 1; $i -ge 0; $i--) { if ([string]$hist[$i].codigo -eq $k) { $hist.RemoveAt($i) } }
        $url = [string]$det.url
        if (-not $url) { $url = $ChamadoUrl + $k }
        [void]$hist.Add([pscustomobject][ordered]@{
                codigo           = $k
                saiuEm           = $stamp
                statusFinal      = [string]$s.statusFinal
                responsavelAtual = [string]$s.responsavelAtual
                ultimoStatus     = [string]$det.status
                tipo             = [string]$det.tipo
                cliente          = [string]$det.cliente
                titulo           = [string]$det.titulo
                categoria        = [string]$det.categoria
                criacao          = [string]$det.criacao
                diasAberto       = [int]$det.diasAberto
                url              = $url
            })
        $mudou = $true
    }
    if ($mudou) { Save-Historico $hist }
}

# hazard: so entra o dia em que uma leitura da fila (a cada 15 min, com a sessao aberta) viu o chamado "em execucao";
# execucao curta entre duas leituras, ou com o painel fechado, fica de fora do Gantt
function Update-Execucao($ch, $quando) {
    if (-not $quando) { return }
    $exec = New-Object psobject
    if (Test-Path -LiteralPath $ExecFile) {
        $exec = Read-Json $ExecFile
        if (-not $exec) { return }   # hazard: arquivo ilegivel; gravar por cima perderia os dias ja anotados
    }
    $dia = $quando.ToString('dd/MM/yyyy', [Globalization.CultureInfo]::InvariantCulture)
    $mudou = $false
    foreach ($c in @($ch | Where-Object { $_.emExecucao })) {
        $k = [string]$c.codigo
        $dias = @($exec.$k | Where-Object { $_ })
        if ($dias -notcontains $dia) {
            Add-Member -InputObject $exec -MemberType NoteProperty -Name $k -Value @($dias + $dia) -Force
            $mudou = $true
        }
    }
    if ($mudou) { try { Write-Json $exec 3 $ExecFile } catch {} }
}

function Process-Data($data, $ch) {
    Update-Execucao $ch (Get-DataLocal $data.atualizadoEm)
    $codes = @($ch | ForEach-Object { [string]$_.codigo })
    $stamp = Get-Stamp $data.atualizadoEm
    if ($stamp -ne (Get-Stamp $state.ultimoAtualizadoEm)) {
        $prev = @($state.ultimos | Where-Object { $_ })
        if ($prev.Count -eq 0) { $prev = @($state.ultimosCodigos | Where-Object { $_ } | ForEach-Object { [pscustomobject]@{ codigo = [string]$_ } }) }
        Update-Historico $prev $ch @($data.saidos | Where-Object { $_ }) $stamp
        # base do "NOVO" = fila do fim do dia anterior (a busca roda varias vezes por dia)
        $antes = Get-DataLocal $state.ultimoAtualizadoEm
        $agora = Get-DataLocal $data.atualizadoEm
        $semBase = @($state.codigosBase | Where-Object { $_ }).Count -eq 0
        if ($semBase -or -not $antes -or -not $agora -or $antes.Date -lt $agora.Date) { $state.codigosBase = @($prev | ForEach-Object { [string]$_.codigo }) }
        $state.ultimos = $ch
        $state.ultimosCodigos = $codes
        $state.ultimoAtualizadoEm = $stamp
        Save-State
    } elseif ($null -eq $state.ultimos -and $ch.Count -gt 0) {
        $state.ultimos = $ch
        Save-State
    }
}

# ---------------------------------------------------------------- datas e tipos
function Get-DataBr([string]$s) {
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($s.Trim(), [string[]]@('dd/MM/yyyy HH:mm', 'dd/MM/yyyy'), [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$d)) { return $d }
    return $null
}

function Get-TipoCategoria([string]$cat) {
    if ($cat -match 'INCID') { return 'Incidência' }
    if ($cat -match 'NOVO REQUISITO') { return 'Novo requisito' }
    if ($cat -match 'CONSULTORIA') { return 'Consultoria' }
    return 'Outro'
}

# ---------------------------------------------------------------- extranet
# Le a lista "Minha responsabilidade" (Chamados.aspx?TipoBusca=lnkResponsabilidade) e devolve JSON por postMessage.
# ponytail: depende dos ids do Repeater da Extranet (rptChamados_*); se a pagina mudar, a busca avisa, salva a pagina e nao grava nada
$ScriptFila = @'
(() => {
  // em falha manda a pagina junto, para ajustar o leitor (vira diagnostico-extranet.html)
  const post = o => window.chrome.webview.postMessage(JSON.stringify(o.ok ? o : { ...o, html: document.documentElement.outerHTML }));
  const t = el => (el ? el.textContent : '').replace(/\s+/g, ' ').trim();
  try {
    const total = (t(document.body).match(/Lista de Chamados \((\d+)\)/) || [])[1];
    if (total === undefined) return post({ ok: false, erro: 'a página não mostrou a Lista de Chamados' });
    const linhas = [...document.querySelectorAll('tr[id^="ContentPlaceHolder1_rptChamados_divChamado_"]')];
    if (linhas.length !== +total) return post({ ok: false, erro: `a página trouxe ${linhas.length} de ${total} chamados` });
    const chamados = linhas.map(tr => {
      const oculto = n => { const i = tr.querySelector(`input[id^="ContentPlaceHolder1_rptChamados_${n}_"]`); return i ? i.value.trim() : ''; };
      const td = n => tr.querySelector(`td[id^="ContentPlaceHolder1_rptChamados_${n}_"]`);
      const painel = tr.querySelector('.panel-heading');
      return {
        codigo: oculto('hddCodigo'),
        cliente: t(td('td1')),
        status: t(td('td2').querySelector('span')),
        diasStatus: parseInt(t(td('td2').querySelector('a')), 10) || 0,
        versao: t(td('td3').querySelector('b')),
        dataPrevista: t(td('td3').querySelector('a')),
        titulo: t(painel.querySelector('h5 a')),
        categoria: t([...painel.childNodes].find(n => n.nodeType === 3 && n.textContent.trim())),   // texto solto logo abaixo do cliente
        criacao: (t(painel).match(/Cria\S*:\s*(\d{2}\/\d{2}\/\d{4})/) || [])[1] || '',
        emExecucao: oculto('HddEmExecucao') === 'S',
        prioridade: oculto('hddPrioridade')
      };
    });
    if (chamados.some(c => !/^\d+$/.test(c.codigo) || !c.titulo)) return post({ ok: false, erro: 'chamado sem código ou título na lista' });
    post({ ok: true, chamados });
  } catch (e) { post({ ok: false, erro: 'formato da lista mudou: ' + e.message }); }
})();
'@

$script:web = $null
$script:webWin = $null
$script:busca = $null          # 'manual' ou 'auto' enquanto busca na Extranet
$script:buscaDesde = [datetime]::MinValue
$script:tentativas = 0
$script:esperandoLogin = $false
$script:sessaoOk = $false
$script:ultimaBusca = [datetime]::MinValue
$script:avisoBusca = ''

# link para outro site: clique da pessoa abre no navegador padrao; redirecionamento da pagina e barrado
function On-NavegacaoIniciada($e) {
    $u = [string]$e.Uri
    if ($u -eq 'about:blank' -or (Test-UrlExtranet $u)) { return }
    $e.Cancel = $true
    if ($u -match '^https?://') {
        if ($e.IsUserInitiated) { Start-Process $u; return }
        if ($script:busca) { Stop-Busca "A Extranet tentou abrir outro site ($(([Uri]$u).Host)): a fila mostra a última leitura." }
    }
}

# Navegador do painel (criado no primeiro uso, depois que a janela principal ja abriu)
function Initialize-Web {
    if ($script:web) { return }
    Add-Type -Path (Join-Path $WebDir 'Microsoft.Web.WebView2.Core.dll')
    Add-Type -Path (Join-Path $WebDir 'Microsoft.Web.WebView2.Wpf.dll')
    [Microsoft.Web.WebView2.Core.CoreWebView2Environment]::SetLoaderDllFolderPath($WebDir)
    $w = New-Object System.Windows.Window
    $w.Title = 'Entrar na Extranet - mininet'
    $w.Width = 520; $w.Height = 700
    $w.Left = -20000; $w.Top = -20000   # fora da tela: so aparece quando precisa de login
    $w.ShowInTaskbar = $false; $w.ShowActivated = $false
    $wv = New-Object Microsoft.Web.WebView2.Wpf.WebView2
    $cp = New-Object Microsoft.Web.WebView2.Wpf.CoreWebView2CreationProperties
    $cp.UserDataFolder = $WebPerfil
    $wv.CreationProperties = $cp
    $wv.Add_CoreWebView2InitializationCompleted({
            param($s, $e)
            if ($e.IsSuccess) {
                # popup (target=_blank, window.open): abre no navegador padrao em vez de outra janela com a sessao da Extranet
                $s.CoreWebView2.Add_NewWindowRequested({
                        param($s2, $e2)
                        $e2.Handled = $true
                        if ($e2.IsUserInitiated -and [string]$e2.Uri -match '^https?://') { Start-Process ([string]$e2.Uri) }
                    })
                if (-not $script:esperandoLogin) { $script:webWin.Hide() }
                return
            }
            $script:web = $null   # tenta criar de novo no proximo Recarregar
            Stop-Busca "Não consegui abrir o navegador do painel: $($e.InitializationException.Message)"
        })
    $wv.Add_NavigationStarting({ param($s, $e) try { On-NavegacaoIniciada $e } catch { $e.Cancel = $true } })
    $wv.Add_NavigationCompleted({ param($s, $e) try { On-Navegou $e } catch { Stop-Busca "Erro ao buscar na Extranet: $($_.Exception.Message)" } })
    $wv.Add_WebMessageReceived({
            param($s, $e)
            if (-not (Test-UrlExtranet $e.Source)) { return }   # so aceita mensagem de pagina da Extranet
            try { On-Fila ($e.TryGetWebMessageAsString()) } catch { Stop-Busca "Erro ao ler a fila: $($_.Exception.Message)" }
        })
    $w.Add_Closing({ param($s, $e) $e.Cancel = $true; Stop-Busca 'Login cancelado: a fila mostra a última leitura.' })
    $w.Content = $wv
    $script:web = $wv
    $script:webWin = $w
    $w.Show()
}

function Show-Login {
    $w = $script:webWin
    $wa = [System.Windows.SystemParameters]::WorkArea
    $w.Height = [Math]::Min(700, $wa.Height - 40)
    $w.Left = $wa.Left + ($wa.Width - $w.Width) / 2
    $w.Top = $wa.Top + ($wa.Height - $w.Height) / 2
    $w.ShowInTaskbar = $true
    $w.Topmost = $true
    $script:esperandoLogin = $true
    $w.Show()
    [void]$w.Activate()
}

function Hide-Login {
    $script:esperandoLogin = $false
    if ($script:webWin) { $script:webWin.Topmost = $false; $script:webWin.ShowInTaskbar = $false; $script:webWin.Hide() }
}

# Janelinha que pede o endereco da Extranet; devolve $true se a pessoa informou um endereco valido
function Pedir-Extranet {
    $d = New-Object System.Windows.Window
    $d.Title = 'Endereço da Extranet - mininet'
    $d.SizeToContent = 'WidthAndHeight'; $d.ResizeMode = 'NoResize'
    $d.WindowStartupLocation = 'CenterScreen'; $d.Topmost = $true
    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Margin = T 16 14 16 14
    $txt = New-Object System.Windows.Controls.TextBlock
    $txt.Text = "Cole o endereço da Extranet, do jeito que aparece no navegador (pode ser a página de login).`nEle fica guardado só neste computador, cifrado com a sua conta do Windows."
    $txt.TextWrapping = 'Wrap'; $txt.Width = 440
    [void]$sp.Children.Add($txt)
    $tb = New-Object System.Windows.Controls.TextBox
    $tb.Margin = T 0 10 0 0; $tb.Padding = T 4 3 4 3
    if ($ListaUrl) { $tb.Text = $ListaUrl }
    [void]$sp.Children.Add($tb)
    $erro = New-Object System.Windows.Controls.TextBlock
    $erro.Text = 'Endereço inválido: ele precisa começar com https://'
    $erro.Foreground = New-Brush '#FFDC2626'; $erro.Margin = T 0 6 0 0; $erro.Visibility = 'Collapsed'
    [void]$sp.Children.Add($erro)
    $bts = New-Object System.Windows.Controls.StackPanel
    $bts.Orientation = 'Horizontal'; $bts.HorizontalAlignment = 'Right'; $bts.Margin = T 0 12 0 0
    $ok = New-Object System.Windows.Controls.Button
    $ok.Content = 'Salvar'; $ok.IsDefault = $true; $ok.Width = 84; $ok.Height = 26
    $cancelar = New-Object System.Windows.Controls.Button
    $cancelar.Content = 'Cancelar'; $cancelar.IsCancel = $true; $cancelar.Width = 84; $cancelar.Height = 26; $cancelar.Margin = T 8 0 0 0
    [void]$bts.Children.Add($ok); [void]$bts.Children.Add($cancelar)
    [void]$sp.Children.Add($bts)
    $d.Content = $sp
    $ok.Add_Click({
            $lista = Get-ListaUrl $tb.Text
            if (-not $lista) { $erro.Visibility = 'Visible'; return }
            $d.Tag = $lista
            $d.DialogResult = $true
        })
    $d.Add_ContentRendered({ [void]$tb.Focus(); $tb.SelectAll() })
    if (-not $d.ShowDialog()) { return $false }
    Set-Extranet ([string]$d.Tag)
    $state.extranet = $ListaUrl
    Save-State
    return $true
}

# modo 'manual' (Recarregar) pede o endereco e abre o login se preciso; 'auto' so avisa
function Start-Busca([string]$modo) {
    if ($script:busca) { if ($modo -eq 'manual' -and $script:esperandoLogin) { Show-Login }; return }
    if (-not $ExtranetDominio) {
        if ($modo -ne 'manual' -or -not (Pedir-Extranet)) {
            $script:avisoBusca = 'Falta o endereço da Extranet: clique em Recarregar para informar.'
            Render
            return
        }
    }
    if (-not (Test-Path -LiteralPath (Join-Path $WebDir 'Microsoft.Web.WebView2.Wpf.dll'))) {
        if ($modo -eq 'manual') { $script:avisoBusca = 'Falta a pasta webview2 com o navegador do painel: a fila não foi buscada na Extranet.'; Render }
        return
    }
    $script:busca = $modo
    $script:tentativas = 0
    $script:avisoBusca = ''
    try { Initialize-Web; Open-Lista } catch { Stop-Busca "Não consegui abrir o navegador do painel: $($_.Exception.Message)"; return }
    $UpdatedText.Text = 'atualizando…'
}

function Open-Lista {
    $script:tentativas++
    $script:buscaDesde = Get-Date
    if ($script:web.CoreWebView2) { $script:web.CoreWebView2.Navigate($ListaUrl) } else { $script:web.Source = [Uri]$ListaUrl }
}

function Stop-Busca([string]$aviso) {
    $script:busca = $null
    $script:avisoBusca = $aviso
    Hide-Login
    Render
}

function On-Navegou($e) {
    if (-not $script:busca) { return }
    if (-not $e.IsSuccess) {
        if ([string]$e.WebErrorStatus -eq 'OperationCanceled') { return }   # outra navegacao tomou o lugar
        Stop-Busca "A Extranet não abriu ($($e.WebErrorStatus)): a fila mostra a última leitura."
        return
    }
    $url = [string]$script:web.Source
    if ($url -match 'Login\.aspx') {
        if ($script:busca -eq 'manual') { Show-Login; return }
        $script:sessaoOk = $false
        Stop-Busca 'Sessão da Extranet fechada: clique em Recarregar para entrar.'
    } elseif ($url -match 'Chamados\.aspx') {
        Hide-Login
        [void]$script:web.ExecuteScriptAsync($ScriptFila)
    } elseif ($script:tentativas -lt 3) {
        Hide-Login   # pagina inicial logo depois do login: segue para a fila
        Open-Lista
    } else { Stop-Busca 'A Extranet não abriu a lista de chamados: a fila mostra a última leitura.' }
}

function On-Fila([string]$json) {
    if (-not $script:busca) { return }
    $r = $json | ConvertFrom-Json
    if (-not $r.ok) {
        if ($r.html) { Set-Content -LiteralPath $DiagFile -Value $r.html -Encoding UTF8 }
        Stop-Busca "Não consegui ler a fila na Extranet ($($r.erro)): a fila mostra a última leitura. Cópia da página em diagnostico-extranet.html."
        return
    }
    # o codigo vira parte do link que o clique abre: confere de novo aqui, nao so no script da pagina
    if (@($r.chamados | Where-Object { [string]$_.codigo -notmatch '^\d+$' }).Count) {
        Stop-Busca 'Não consegui ler a fila na Extranet (chamado com código inválido): a fila mostra a última leitura.'
        return
    }
    $hoje = (Get-Date).Date
    $ch = @(foreach ($c in @($r.chamados)) {
            $cad = Get-DataBr $c.criacao
            $dias = 0
            if ($cad) { $dias = ($hoje - $cad.Date).Days }
            [pscustomobject][ordered]@{
                codigo       = [string]$c.codigo
                cliente      = $c.cliente
                status       = $c.status
                diasStatus   = [int]$c.diasStatus
                versao       = $c.versao
                dataPrevista = $c.dataPrevista
                titulo       = $c.titulo
                categoria    = $c.categoria
                tipo         = Get-TipoCategoria $c.categoria
                criacao      = $c.criacao
                diasAberto   = $dias
                emExecucao   = [bool]$c.emExecucao
                prioridade   = $c.prioridade
                url          = $ChamadoUrl + $c.codigo
            }
        })
    $novo = [pscustomobject][ordered]@{ atualizadoEm = Get-Stamp (Get-Date); chamados = $ch }
    Write-Json $novo 4 $DataFile
    Remove-Item -LiteralPath $DiagFile -ErrorAction SilentlyContinue   # a copia da pagina tem seus chamados: nao fica depois que o leitor volta a funcionar
    $script:sessaoOk = $true
    $script:ultimaBusca = Get-Date
    Stop-Busca ''
}

# ---------------------------------------------------------------- visual
function New-Brush([string]$hex) {
    $b = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($hex))
    $b.Freeze()
    return $b
}
function T([double]$l, [double]$t, [double]$r, [double]$b) { New-Object System.Windows.Thickness($l, $t, $r, $b) }

$BrushText    = New-Brush '#FFF1F3F5'
$BrushTitle   = New-Brush '#FFDDE1E6'
$BrushMuted   = New-Brush '#FF9AA0A6'
$BrushIcon    = New-Brush '#FFB0B4BA'
$BrushRecent  = New-Brush '#FFFCD34D'
$BrushRow     = New-Brush '#14FFFFFF'
$BrushRowOver = New-Brush '#26FFFFFF'
$BrushBadge   = New-Brush '#FF3B82F6'
$BrushBadgeH  = New-Brush '#FF475569'
$Mono         = New-Object System.Windows.Media.FontFamily('Cascadia Mono, Consolas')
$TextInfo     = (Get-Culture).TextInfo

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="mininet" Width="370" SizeToContent="Height"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        Topmost="True" ShowInTaskbar="False" ResizeMode="NoResize"
        WindowStartupLocation="Manual" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="FontFamily" Value="Segoe MDL2 Assets"/>
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="Foreground" Value="#FFB0B4BA"/>
      <Setter Property="Width" Value="26"/>
      <Setter Property="Height" Value="22"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="B" Background="Transparent" CornerRadius="4">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="B" Property="Background" Value="#26FFFFFF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>
  <Border x:Name="Root" Margin="8" CornerRadius="10" Background="#F21F2125" BorderBrush="#33FFFFFF" BorderThickness="1">
    <Border.Effect>
      <DropShadowEffect BlurRadius="12" ShadowDepth="2" Opacity="0.4"/>
    </Border.Effect>
    <StackPanel>
      <Grid x:Name="Header" Background="Transparent" Margin="12,7,6,7" Cursor="SizeAll">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock x:Name="TitleText" Text="Minha fila" Foreground="#FFF1F3F5" FontWeight="SemiBold" FontSize="13"/>
          <Border x:Name="CountBadge" Background="#FF3B82F6" CornerRadius="8" Padding="7,0,7,1" Margin="8,0,0,0" VerticalAlignment="Center">
            <TextBlock x:Name="CountText" Text="-" Foreground="White" FontWeight="Bold" FontSize="11"/>
          </Border>
          <TextBlock x:Name="UpdatedText" Foreground="#FF9AA0A6" FontSize="11" Margin="8,1,0,0" VerticalAlignment="Center"/>
        </StackPanel>
        <StackPanel Grid.Column="1" Orientation="Horizontal" Cursor="Arrow">
          <Button x:Name="BtnCal" Content="&#xE787;" ToolTip="Calendário / fila"/>
          <Button x:Name="BtnHist" Content="&#xE81C;" ToolTip="Histórico / fila"/>
          <Button x:Name="BtnRefresh" Content="&#xE72C;" ToolTip="Recarregar da Extranet"/>
          <Button x:Name="BtnCollapse" Content="&#xE70D;" ToolTip="Recolher / expandir"/>
          <Button x:Name="BtnClose" Content="&#xE8BB;" ToolTip="Fechar"/>
        </StackPanel>
      </Grid>
      <TextBlock x:Name="Warn" Visibility="Collapsed" Foreground="#FFFBBF24" FontSize="11" TextWrapping="Wrap" Margin="12,0,12,7"/>
      <ScrollViewer x:Name="Scroll" MaxHeight="520" VerticalScrollBarVisibility="Auto" Margin="6,0,6,7">
        <StackPanel x:Name="List"/>
      </ScrollViewer>
    </StackPanel>
  </Border>
</Window>
'@

function New-Tag([string]$text, [string]$bg, [string]$fg) {
    $b = New-Object System.Windows.Controls.Border
    $b.Background = New-Brush $bg
    $b.CornerRadius = New-Object System.Windows.CornerRadius(4)
    $b.Padding = T 5 0 5 1
    $b.Margin = T 6 1 0 1
    $b.VerticalAlignment = 'Center'
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = $text
    $t.FontSize = 10
    $t.FontWeight = [System.Windows.FontWeights]::SemiBold
    $t.Foreground = New-Brush $fg
    $b.Child = $t
    return $b
}

function New-TipoTag([string]$tipo) {
    switch ($tipo) {
        'Incidência'     { return New-Tag 'Incidência' '#33EF4444' '#FFFCA5A5' }
        'Novo requisito' { return New-Tag 'Novo requisito' '#333B82F6' '#FF93C5FD' }
        'Consultoria'    { return New-Tag 'Consultoria' '#338B5CF6' '#FFC4B5FD' }
        default          { if ($tipo -and $tipo -ne 'Outro') { return New-Tag $tipo '#26FFFFFF' '#FFD1D5DB' } }
    }
    return $null
}

function New-StatusTag([string]$s) {
    if (-not $s) { return $null }
    if ($s -match 'Atualizado no Cliente|^Homologado$|^Aprovado$') { return New-Tag $s '#3322C55E' '#FF86EFAC' }
    if ($s -match 'Cancelado') { return New-Tag $s '#33EF4444' '#FFFCA5A5' }
    if ($s -match 'Devolvido|Retido') { return New-Tag $s '#40F59E0B' '#FFFCD34D' }
    if ($s -match 'Homologa|Aguardando Atualiza') { return New-Tag $s '#3314B8A6' '#FF5EEAD4' }
    return New-Tag $s '#26FFFFFF' '#FFD1D5DB'
}

function Get-ClienteCurto([string]$c) {
    $s = ($c -replace '\s*\(.*$', '').Trim()
    if ($s.Length -gt 22 -and $s.Contains(' - ')) {
        $p = $s -split ' - '
        $s = $p[-1]
        # "CIDADE - SIGLA - INSTITUTO DE ..." fica "CIDADE - SIGLA"
        if ($s.Length -gt 22) { $s = $p[0..($p.Count - 2)] -join ' - ' }
    }
    return $s
}

function Get-NomeCurto([string]$n) {
    $p = @($n.Trim() -split '\s+' | Where-Object { $_ })
    if ($p.Count -eq 0) { return '' }
    return $TextInfo.ToTitleCase((($p | Select-Object -First 2) -join ' ').ToLower())
}

# "na fila de atualização" para filas, "com Fulano Tal" para pessoas
function Get-ComQuem([string]$r) {
    if ($r -match '^\s*FILA\b') { return 'na ' + $r.Trim().ToLower() }
    return 'com ' + (Get-NomeCurto $r)
}

function New-RowBorder($item, [string]$tip) {
    $row = New-Object System.Windows.Controls.Border
    $row.CornerRadius = New-Object System.Windows.CornerRadius(7)
    $row.Padding = T 9 6 9 7
    $row.Margin = T 0 2 0 2
    $row.Background = $BrushRow
    $row.Cursor = [System.Windows.Input.Cursors]::Hand
    $row.Tag = $item
    $row.ToolTip = $tip
    $row.Add_MouseEnter({ $this.Background = $BrushRowOver })
    $row.Add_MouseLeave({ $this.Background = $BrushRow })
    $row.Add_MouseLeftButtonUp({ if (Test-UrlExtranet $this.Tag.url) { Start-Process ([string]$this.Tag.url) } })
    $row.Add_MouseRightButtonDown({ $script:alvo = $this.Tag })   # chamado do menu do botao direito
    return $row
}

function New-Line1([string]$codigo, $tags, [string]$direita) {
    $l1 = New-Object System.Windows.Controls.DockPanel
    $r = New-Object System.Windows.Controls.TextBlock
    $r.Text = $direita
    $r.FontSize = 10.5
    $r.Foreground = $BrushMuted
    $r.VerticalAlignment = 'Center'
    $r.Margin = T 6 0 0 0
    [System.Windows.Controls.DockPanel]::SetDock($r, 'Right')
    [void]$l1.Children.Add($r)
    $left = New-Object System.Windows.Controls.WrapPanel   # com muitas tags, quebra a linha em vez de cortar
    $cod = New-Object System.Windows.Controls.TextBlock
    $cod.Text = $codigo
    $cod.FontFamily = $Mono
    $cod.FontWeight = [System.Windows.FontWeights]::Bold
    $cod.Foreground = $BrushText
    $cod.VerticalAlignment = 'Center'
    [void]$left.Children.Add($cod)
    foreach ($tg in @($tags)) { if ($tg) { [void]$left.Children.Add($tg) } }
    [void]$l1.Children.Add($left)
    return $l1
}

function New-TextLine([string]$text, $brush, [double]$size, [double]$top) {
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = $text
    $t.Foreground = $brush
    $t.FontSize = $size
    $t.TextTrimming = 'CharacterEllipsis'
    $t.Margin = T 0 $top 0 0
    return $t
}

function New-Row($c, [bool]$novo) {
    $tip = "$($c.codigo) - $($c.titulo)`n$($c.cliente)`n$($c.categoria)"
    if ($c.versao) { $tip += "`nVersão: $($c.versao)" }
    if ($c.dataPrevista) { $tip += "`nPrevisto para $($c.dataPrevista)" }
    $row = New-RowBorder $c "$tip`nCriado em $($c.criacao)`nClique para abrir na Extranet"
    $sp = New-Object System.Windows.Controls.StackPanel
    $tags = @((New-TipoTag ([string]$c.tipo)))
    if ($c.emExecucao) { $tags += (New-Tag 'em execução' '#3322C55E' '#FF86EFAC') }
    if ($novo)         { $tags += (New-Tag 'NOVO' '#40F59E0B' '#FFFCD34D') }
    $prev = Get-DataBr ([string]$c.dataPrevista)
    if ($prev -and $prev.Date -lt (Get-Date).Date) { $tags += (New-Tag 'atrasado' '#FFB91C1C' '#FFFFFFFF') }
    [void]$sp.Children.Add((New-Line1 ([string]$c.codigo) $tags "$([int]$c.diasAberto) d em aberto"))
    [void]$sp.Children.Add((New-TextLine ([string]$c.titulo) $BrushTitle 12 3))
    $meta = New-Object System.Windows.Controls.TextBlock
    $meta.FontSize = 11
    $meta.Margin = T 0 2 0 0
    $meta.TextTrimming = 'CharacterEllipsis'
    $r1 = New-Object System.Windows.Documents.Run ("$(Get-ClienteCurto ([string]$c.cliente))  ·  ")
    $r1.Foreground = $BrushMuted
    $ds = [int]$c.diasStatus
    if ($ds -eq 0) { $quando = 'hoje' } elseif ($ds -eq 1) { $quando = 'há 1 dia' } else { $quando = "há $ds dias" }
    $r2 = New-Object System.Windows.Documents.Run ("$($c.status) $quando")
    if ($ds -le 1) { $r2.Foreground = $BrushRecent } else { $r2.Foreground = $BrushMuted }
    [void]$meta.Inlines.Add($r1)
    [void]$meta.Inlines.Add($r2)
    [void]$sp.Children.Add($meta)
    $row.Child = $sp
    return $row
}

function New-HistRow($e) {
    $tip = "$($e.codigo) - $($e.titulo)`n$($e.cliente)`nStatus final: $($e.statusFinal)`nResponsável atual: $($e.responsavelAtual)"
    if ($e.ultimoStatus) { $tip += "`nÚltimo status na sua fila: $($e.ultimoStatus)" }
    $row = New-RowBorder $e "$tip`nClique para abrir na Extranet"
    $sp = New-Object System.Windows.Controls.StackPanel
    $d = Get-DataLocal $e.saiuEm
    $quando = ''
    if ($d) { $quando = 'fora desde ' + $d.ToString('dd/MM') }
    $tags = @((New-TipoTag ([string]$e.tipo)), (New-StatusTag ([string]$e.statusFinal)))
    [void]$sp.Children.Add((New-Line1 ([string]$e.codigo) $tags $quando))
    if ($e.titulo) { [void]$sp.Children.Add((New-TextLine ([string]$e.titulo) $BrushTitle 12 3)) }
    $partes = @()
    if ($e.cliente) { $partes += (Get-ClienteCurto ([string]$e.cliente)) }
    if ($e.responsavelAtual) {
        $partes += ('agora ' + (Get-ComQuem ([string]$e.responsavelAtual)))
    } elseif ($e.ultimoStatus) {
        $partes += ('saiu de: ' + [string]$e.ultimoStatus)
    }
    if ($partes.Count -gt 0) { [void]$sp.Children.Add((New-TextLine ($partes -join '  ·  ') $BrushMuted 11 2)) }
    $row.Child = $sp
    return $row
}

function Show-Warn([string]$msg) {
    if ($msg) { $Warn.Text = $msg; $Warn.Visibility = 'Visible' } else { $Warn.Text = ''; $Warn.Visibility = 'Collapsed' }
}

function New-Vazio([string]$msg) {
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = $msg
    $t.Foreground = $BrushMuted
    $t.TextWrapping = 'Wrap'
    $t.Margin = T 6 2 6 4
    return $t
}

function New-Secao([string]$texto, [string]$tip) {
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = $texto
    $t.ToolTip = $tip
    $t.Foreground = $BrushMuted
    $t.FontSize = 11
    $t.FontWeight = [System.Windows.FontWeights]::SemiBold
    $t.Margin = T 6 8 6 2
    return $t
}

function Update-Warn($data) {
    if ($script:avisoBusca) { Show-Warn $script:avisoBusca; return }
    if (-not $data) { Show-Warn 'Ainda sem dados. Clique em Recarregar para buscar a fila na Extranet.'; return }
    $quando = Get-DataLocal $data.atualizadoEm
    $agora = Get-Date
    $diaUtil = $agora.DayOfWeek -ne 'Saturday' -and $agora.DayOfWeek -ne 'Sunday'
    if ($quando -and $quando.Date -lt $agora.Date -and $diaUtil -and $agora.TimeOfDay -gt [timespan]'08:15') {
        Show-Warn 'A fila ainda não foi atualizada hoje. Clique em Recarregar para buscar na Extranet.'
    } else { Show-Warn '' }
}

function Draw-Fila($data, $ch) {
    $TitleText.Text = 'Minha fila'
    $CountBadge.Background = $BrushBadge
    $List.Children.Clear()
    if (-not $data) { $CountText.Text = '-'; $UpdatedText.Text = '' }
    else {
        $CountText.Text = [string]$ch.Count
        $quando = Get-DataLocal $data.atualizadoEm
        if ($quando) {
            if ($quando.Date -eq (Get-Date).Date) { $UpdatedText.Text = 'atualizado às ' + $quando.ToString('HH:mm') }
            else { $UpdatedText.Text = 'atualizado em ' + $quando.ToString('dd/MM HH:mm') }
        } else { $UpdatedText.Text = '' }
        if ($script:busca) { $UpdatedText.Text = 'atualizando…' }
        if ($ch.Count -eq 0) { [void]$List.Children.Add((New-Vazio 'Nenhum chamado na sua fila.')) }
        $base = @($state.codigosBase | Where-Object { $_ } | ForEach-Object { [string]$_ })
        $ordenados = $ch | Sort-Object `
            @{ Expression = { if ($_.emExecucao) { 0 } else { 1 } } }, `
            @{ Expression = { if ([string]$_.tipo -eq 'Incidência') { 0 } else { 1 } } }, `
            @{ Expression = { [int]$_.diasStatus } }
        foreach ($c in $ordenados) {
            $novo = ($base.Count -gt 0) -and ($base -notcontains [string]$c.codigo)
            [void]$List.Children.Add((New-Row $c $novo))
        }
    }
}

function Draw-Historico {
    $TitleText.Text = 'Histórico'
    $CountBadge.Background = $BrushBadgeH
    $List.Children.Clear()
    $h = Get-Historico
    $hist = @($h | Sort-Object { Get-Stamp $_.saiuEm } -Descending)
    $CountText.Text = [string]$hist.Count
    $UpdatedText.Text = 'saíram da sua fila'
    if ($hist.Count -eq 0) {
        [void]$List.Children.Add((New-Vazio 'Nenhum chamado saiu da fila ainda. Quando um sair, ele aparece aqui com o status final e com quem ficou.'))
        return
    }
    foreach ($e in ($hist | Select-Object -First 60)) { [void]$List.Children.Add((New-HistRow $e)) }
}

# ---------------------------------------------------------------- calendario
$BrushChegou  = New-Brush '#FF60A5FA'
$BrushEntrega = New-Brush '#FF4ADE80'
$BrushDiaSel  = New-Brush '#403B82F6'
$PtBr         = [Globalization.CultureInfo]::GetCultureInfo('pt-BR')
$script:calMes  = (Get-Date).Date.AddDays(1 - (Get-Date).Day)
$script:calTipo = ''      # '' = todos os tipos
$script:calDia  = $null
$script:calVista = 'calendario'   # ou 'gantt'

# why: o painel nao guarda o dia em que o chamado entrou na fila, entao "chegou" usa a data de criacao
function Get-EventosCalendario($ch, $hist) {
    foreach ($c in @(@($ch) + @($hist) | Where-Object { $_ })) {
        $d = Get-DataBr ([string]$c.criacao)
        if ($d) { [pscustomobject]@{ dia = $d.Date; evento = 'chegou'; item = $c } }
    }
    foreach ($e in @($hist | Where-Object { $_ })) {
        $d = Get-DataLocal $e.saiuEm
        if ($d) { [pscustomobject]@{ dia = $d.Date; evento = 'entregue'; item = $e } }
    }
}

# invariant: dias seguidos em execucao viram um trecho so; chamado sem dia anotado fica de fora
function Get-PeriodosGantt($ch, $hist, $exec) {
    foreach ($c in @(@($ch) + @($hist) | Where-Object { $_ })) {
        $trechos = @()
        foreach ($d in @($exec.([string]$c.codigo) | ForEach-Object { Get-DataBr $_ } | Where-Object { $_ } | Sort-Object -Unique)) {
            if ($trechos.Count -and $trechos[-1].fim.AddDays(1) -eq $d) { $trechos[-1].fim = $d }
            else { $trechos += [pscustomobject]@{ ini = $d; fim = $d } }
        }
        if ($trechos.Count) { [pscustomobject]@{ item = $c; trechos = $trechos } }
    }
}

function Draw-Gantt($ch, [datetime]$ini, [datetime]$fim) {
    $hoje = (Get-Date).Date
    $per = @(foreach ($p in @(Get-PeriodosGantt $ch (Get-Historico) (Read-Json $ExecFile))) {
            $t = @($p.trechos | Where-Object { $_.ini -lt $fim -and $_.fim -ge $ini })
            if ($t.Count -and (-not $script:calTipo -or [string]$p.item.tipo -eq $script:calTipo)) { [pscustomobject]@{ item = $p.item; trechos = $t } }
        })
    if ($per.Count -eq 0) { [void]$List.Children.Add((New-Vazio 'Nenhum dia em execução anotado neste mês. O painel anota os chamados "em execução" a cada leitura da fila.')); return }
    $per = @($per | Sort-Object { $_.trechos[0].ini }, { [string]$_.item.codigo })
    $larg = 250.0
    $px = $larg / ($fim - $ini).Days
    $xHoje = -1
    if ($hoje -ge $ini -and $hoje -lt $fim) { $xHoje = ($hoje - $ini).Days * $px + $px / 2 }

    # invariant: margem 60 = padding da linha (6) + coluna do codigo (48) + folga, para o eixo ficar sobre as barras
    $eixo = New-Object System.Windows.Controls.Canvas
    $eixo.Width = $larg; $eixo.Height = 14
    $eixo.HorizontalAlignment = 'Left'
    $eixo.Margin = T 60 2 0 0
    foreach ($d in @(1, 8, 15, 22, 29)) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = [string]$d; $t.FontSize = 9.5; $t.Foreground = $BrushMuted
        [System.Windows.Controls.Canvas]::SetLeft($t, ($d - 1) * $px)
        [void]$eixo.Children.Add($t)
    }
    [void]$List.Children.Add($eixo)

    foreach ($p in $per) {
        $c = $p.item
        $cor = $BrushChegou
        if ($c.saiuEm) { $cor = $BrushEntrega }
        $quando = @($p.trechos | ForEach-Object { if ($_.ini -eq $_.fim) { $_.ini.ToString('dd/MM') } else { "$($_.ini.ToString('dd/MM')) a $($_.fim.ToString('dd/MM'))" } }) -join ', '
        $row = New-RowBorder $c "$($c.codigo) - $($c.titulo)`n$($c.cliente)`nEm execução: $quando`nClique para abrir na Extranet"
        $row.Padding = T 6 3 6 3
        $row.Margin = T 0 1 0 1
        $dp = New-Object System.Windows.Controls.DockPanel
        $cod = New-Object System.Windows.Controls.TextBlock
        $cod.Text = [string]$c.codigo; $cod.Width = 48; $cod.FontSize = 10.5
        $cod.FontFamily = $Mono; $cod.Foreground = $BrushText; $cod.VerticalAlignment = 'Center'
        [void]$dp.Children.Add($cod)
        $cv = New-Object System.Windows.Controls.Canvas
        $cv.Width = $larg; $cv.Height = 12
        $cv.HorizontalAlignment = 'Left'
        if ($xHoje -ge 0) {
            $l = New-Object System.Windows.Controls.Border
            $l.Width = 1; $l.Height = 12; $l.Background = $BrushMuted
            [System.Windows.Controls.Canvas]::SetLeft($l, $xHoje)
            [void]$cv.Children.Add($l)
        }
        foreach ($t in $p.trechos) {
            $s = $t.ini; if ($s -lt $ini) { $s = $ini }
            $e = $t.fim; if ($e -ge $fim) { $e = $fim.AddDays(-1) }
            $b = New-Object System.Windows.Controls.Border
            $b.Height = 8
            $b.Width = (($e - $s).Days + 1) * $px
            $b.CornerRadius = New-Object System.Windows.CornerRadius(3)
            $b.Background = $cor
            [System.Windows.Controls.Canvas]::SetLeft($b, ($s - $ini).Days * $px)
            [System.Windows.Controls.Canvas]::SetTop($b, 2)
            [void]$cv.Children.Add($b)
        }
        [void]$dp.Children.Add($cv)
        $row.Child = $dp
        [void]$List.Children.Add($row)
    }
    [void]$List.Children.Add((New-Vazio 'Barras: dias em que o chamado estava em execução. Azul: ainda na fila. Verde: entregue. Linha: hoje.'))
}

function New-Barra($brush, [int]$n, [int]$max) {
    $b = New-Object System.Windows.Controls.Border
    $b.Width = 7
    $b.Margin = T 1 0 1 0
    $b.CornerRadius = New-Object System.Windows.CornerRadius(2)
    $b.VerticalAlignment = 'Bottom'
    $b.Background = $brush
    $b.Height = 0
    if ($n -gt 0) { $b.Height = [Math]::Max(3, 20 * $n / $max) }
    return $b
}

function Draw-Calendario($ch) {
    $TitleText.Text = 'Calendário'
    $CountBadge.Background = $BrushBadgeH
    $CountText.Text = [string]$ch.Count
    $UpdatedText.Text = 'na fila agora'
    $List.Children.Clear()
    $ini = $script:calMes
    $fim = $ini.AddMonths(1)
    $ev = @(Get-EventosCalendario $ch (Get-Historico) | Where-Object {
            $_.dia -ge $ini -and $_.dia -lt $fim -and (-not $script:calTipo -or [string]$_.item.tipo -eq $script:calTipo)
        })
    $porDia = @{}
    foreach ($e in $ev) {
        if (-not $porDia.ContainsKey($e.dia.Day)) { $porDia[$e.dia.Day] = @{ chegou = @(); entregue = @() } }
        $porDia[$e.dia.Day][$e.evento] += $e.item
    }
    $max = 1
    foreach ($v in $porDia.Values) { $max = [Math]::Max($max, [Math]::Max($v.chegou.Count, $v.entregue.Count)) }

    $vistas = New-Object System.Windows.Controls.WrapPanel
    $vistas.HorizontalAlignment = 'Center'
    foreach ($f in @(@('Calendário', 'calendario'), @('Gantt', 'gantt'))) {
        if ($script:calVista -eq $f[1]) { $chip = New-Tag $f[0] '#FF475569' '#FFFFFFFF' } else { $chip = New-Tag $f[0] '#26FFFFFF' '#FFD1D5DB' }
        $chip.Tag = $f[1]
        $chip.Cursor = [System.Windows.Input.Cursors]::Hand
        $chip.Add_MouseLeftButtonUp({ $script:calVista = [string]$this.Tag; Render })
        [void]$vistas.Children.Add($chip)
    }
    [void]$List.Children.Add($vistas)

    # mes: < Setembro 2026 >  (clique no nome volta para o mes atual)
    $nav = New-Object System.Windows.Controls.DockPanel
    $nav.Margin = T 2 2 2 2
    $bAnt = New-Object System.Windows.Controls.Button
    $bAnt.Content = [string][char]0xE76B; $bAnt.ToolTip = 'Mês anterior'
    $bAnt.Add_Click({ $script:calMes = $script:calMes.AddMonths(-1); $script:calDia = $null; Render })
    $bProx = New-Object System.Windows.Controls.Button
    $bProx.Content = [string][char]0xE76C; $bProx.ToolTip = 'Próximo mês'
    $bProx.Add_Click({ $script:calMes = $script:calMes.AddMonths(1); $script:calDia = $null; Render })
    [System.Windows.Controls.DockPanel]::SetDock($bAnt, 'Left')
    [System.Windows.Controls.DockPanel]::SetDock($bProx, 'Right')
    [void]$nav.Children.Add($bAnt)
    [void]$nav.Children.Add($bProx)
    $mes = New-Object System.Windows.Controls.TextBlock
    $mes.Text = $PtBr.TextInfo.ToTitleCase($ini.ToString('MMMM yyyy', $PtBr))
    $mes.Foreground = $BrushText
    $mes.FontWeight = [System.Windows.FontWeights]::SemiBold
    $mes.HorizontalAlignment = 'Center'
    $mes.VerticalAlignment = 'Center'
    $mes.Cursor = [System.Windows.Input.Cursors]::Hand
    $mes.ToolTip = 'Voltar para o mês atual'
    $mes.Add_MouseLeftButtonUp({ $script:calMes = (Get-Date).Date.AddDays(1 - (Get-Date).Day); $script:calDia = $null; Render })
    [void]$nav.Children.Add($mes)
    [void]$List.Children.Add($nav)

    # resumo do mes (serve de legenda das barras)
    $cheg = @($ev | Where-Object { $_.evento -eq 'chegou' }).Count
    $ent = @($ev | Where-Object { $_.evento -eq 'entregue' }).Count
    $saldo = $cheg - $ent
    $res = New-Object System.Windows.Controls.TextBlock
    $res.FontSize = 11
    $res.Margin = T 6 0 6 4
    $res.HorizontalAlignment = 'Center'
    foreach ($p in @(@("■ $cheg chegaram", $BrushChegou), @('   ', $BrushMuted), @("■ $ent entregues", $BrushEntrega), @("   saldo $(if ($saldo -gt 0) { '+' })$saldo", $BrushMuted))) {
        $run = New-Object System.Windows.Documents.Run ($p[0])
        $run.Foreground = $p[1]
        [void]$res.Inlines.Add($run)
    }
    $res.ToolTip = "Chegou = data de criação do chamado. Entregue = dia em que saiu da sua fila.`nSaldo positivo: a fila cresceu no mês."
    [void]$List.Children.Add($res)

    # filtro por tipo
    $filtro = New-Object System.Windows.Controls.WrapPanel
    $filtro.HorizontalAlignment = 'Center'
    $filtro.Margin = T 0 0 6 6
    foreach ($f in @(@('Todos', ''), @('Incidência', 'Incidência'), @('Novo requisito', 'Novo requisito'), @('Consultoria', 'Consultoria'))) {
        if ($script:calTipo -eq $f[1]) { $chip = New-Tag $f[0] '#FF3B82F6' '#FFFFFFFF' } else { $chip = New-Tag $f[0] '#26FFFFFF' '#FFD1D5DB' }
        $chip.Tag = $f[1]
        $chip.Cursor = [System.Windows.Input.Cursors]::Hand
        $chip.Add_MouseLeftButtonUp({ $script:calTipo = [string]$this.Tag; Render })
        [void]$filtro.Children.Add($chip)
    }
    [void]$List.Children.Add($filtro)
    if ($script:calVista -eq 'gantt') { Draw-Gantt $ch $ini $fim; return }

    $sem =New-Object System.Windows.Controls.Primitives.UniformGrid
    $sem.Columns = 7
    foreach ($d in @('D', 'S', 'T', 'Q', 'Q', 'S', 'S')) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = $d; $t.FontSize = 10; $t.Foreground = $BrushMuted; $t.HorizontalAlignment = 'Center'
        [void]$sem.Children.Add($t)
    }
    [void]$List.Children.Add($sem)

    $grade = New-Object System.Windows.Controls.Primitives.UniformGrid
    $grade.Columns = 7
    for ($i = 0; $i -lt [int]$ini.DayOfWeek; $i++) { [void]$grade.Children.Add((New-Object System.Windows.Controls.Border)) }
    $hoje = (Get-Date).Date
    for ($dia = $ini; $dia -lt $fim; $dia = $dia.AddDays(1)) {
        $v = $porDia[$dia.Day]
        $nC = 0; $nE = 0
        if ($v) { $nC = $v.chegou.Count; $nE = $v.entregue.Count }
        $cel = New-Object System.Windows.Controls.Border
        $cel.Height = 42
        $cel.Margin = T 1 1 1 1
        $cel.Padding = T 3 1 3 3
        $cel.CornerRadius = New-Object System.Windows.CornerRadius(5)
        $cel.Background = $BrushRow
        if ($script:calDia -eq $dia) { $cel.Background = $BrushDiaSel }
        if ($dia -eq $hoje) { $cel.BorderBrush = $BrushBadge; $cel.BorderThickness = T 1 1 1 1 }
        $cel.Cursor = [System.Windows.Input.Cursors]::Hand
        $cel.Tag = $dia
        $tip = $dia.ToString('dddd, dd/MM', $PtBr)
        if ($v) {
            if ($nC) { $tip += "`nChegaram ($nC): " + (@($v.chegou | ForEach-Object { $_.codigo }) -join ', ') }
            if ($nE) { $tip += "`nEntregues ($nE): " + (@($v.entregue | ForEach-Object { $_.codigo }) -join ', ') }
        }
        $cel.ToolTip = $tip
        $cel.Add_MouseLeftButtonUp({ if ($script:calDia -eq $this.Tag) { $script:calDia = $null } else { $script:calDia = $this.Tag }; Render })
        $g = New-Object System.Windows.Controls.Grid
        $num = New-Object System.Windows.Controls.TextBlock
        $num.Text = [string]$dia.Day
        $num.FontSize = 10
        $num.Foreground = $BrushMuted
        if ($v) { $num.Foreground = $BrushText }
        [void]$g.Children.Add($num)
        $barras = New-Object System.Windows.Controls.StackPanel
        $barras.Orientation = 'Horizontal'
        $barras.HorizontalAlignment = 'Center'
        $barras.VerticalAlignment = 'Bottom'
        [void]$barras.Children.Add((New-Barra $BrushChegou $nC $max))
        [void]$barras.Children.Add((New-Barra $BrushEntrega $nE $max))
        [void]$g.Children.Add($barras)
        $cel.Child = $g
        [void]$grade.Children.Add($cel)
    }
    [void]$List.Children.Add($grade)

    $sel = $script:calDia
    if (-not $sel -or $sel -lt $ini -or $sel -ge $fim) {
        [void]$List.Children.Add((New-Vazio 'Clique num dia para ver os chamados que chegaram e os que foram entregues.'))
        return
    }
    $v = $porDia[$sel.Day]
    if (-not $v) { [void]$List.Children.Add((New-Vazio "Nada em $($sel.ToString('dd/MM')).")); return }
    if ($v.chegou.Count) {
        [void]$List.Children.Add((New-Secao "Chegaram em $($sel.ToString('dd/MM')) · $($v.chegou.Count)" 'Pela data de criação do chamado'))
        foreach ($c in $v.chegou) { if ($c.saiuEm) { [void]$List.Children.Add((New-HistRow $c)) } else { [void]$List.Children.Add((New-Row $c $false)) } }
    }
    if ($v.entregue.Count) {
        [void]$List.Children.Add((New-Secao "Entregues em $($sel.ToString('dd/MM')) · $($v.entregue.Count)" 'Saíram da sua fila neste dia'))
        foreach ($e in $v.entregue) { [void]$List.Children.Add((New-HistRow $e)) }
    }
}

$script:lastWrite = [datetime]::MinValue
$script:lastRender = [datetime]::MinValue
$script:alvo = $null   # chamado sob o clique direito (menu)

function Render {
    $script:lastRender = Get-Date
    $data = $null
    if (Test-Path -LiteralPath $DataFile) {
        $data = Read-Json $DataFile
        if (-not $data) { return }   # arquivo sendo gravado: tenta de novo no proximo ciclo
        $script:lastWrite = (Get-Item -LiteralPath $DataFile).LastWriteTimeUtc
    }
    $ch = @()
    if ($data) {
        $ch = @($data.chamados | Where-Object { $_ })
        Process-Data $data $ch
    }
    Update-Warn $data
    $BtnHist.Foreground = $BrushIcon
    $BtnCal.Foreground = $BrushIcon
    if ($state.modo -eq 'historico') {
        Draw-Historico
        $BtnHist.Foreground = $BrushText
    } elseif ($state.modo -eq 'calendario') {
        Draw-Calendario $ch
        $BtnCal.Foreground = $BrushText
    } else {
        Draw-Fila $data $ch
    }
}

function Set-Recolhido([bool]$v) {
    $state.recolhido = $v
    if ($v) { $Scroll.Visibility = 'Collapsed'; $BtnCollapse.Content = [string][char]0xE70E }
    else    { $Scroll.Visibility = 'Visible';   $BtnCollapse.Content = [string][char]0xE70D }
}

function Place-Window {
    $w = $win.ActualWidth; $h = $win.ActualHeight
    if ($w -le 0) { $w = $win.Width }
    if ($h -le 0) { $h = 200 }
    if ($null -ne $state.left -and $null -ne $state.top) {
        $vl = [System.Windows.SystemParameters]::VirtualScreenLeft
        $vt = [System.Windows.SystemParameters]::VirtualScreenTop
        $vr = $vl + [System.Windows.SystemParameters]::VirtualScreenWidth
        $vb = $vt + [System.Windows.SystemParameters]::VirtualScreenHeight
        $win.Left = [Math]::Max($vl, [Math]::Min([double]$state.left, $vr - $w))
        $win.Top  = [Math]::Max($vt, [Math]::Min([double]$state.top,  $vb - $h))
    } else {
        $wa = [System.Windows.SystemParameters]::WorkArea
        $win.Left = $wa.Right - $w - 4
        $win.Top  = $wa.Bottom - $h - 4
    }
}

function Toggle-Startup {
    if (Test-Path -LiteralPath $StartupLnk) {
        Remove-Item -LiteralPath $StartupLnk -Force
    } else {
        $ws = New-Object -ComObject WScript.Shell
        $lnk = $ws.CreateShortcut($StartupLnk)
        $lnk.TargetPath = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $lnk.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""
        $lnk.WorkingDirectory = $Dir
        $lnk.WindowStyle = 7
        $lnk.Save()
    }
    $miStartup.IsChecked = (Test-Path -LiteralPath $StartupLnk)
}

Remove-PlanilhaHistorico   # sobra de quando o painel fechou com o Excel aberto

try {
    $win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$xaml)))
    foreach ($n in @('Root', 'Header', 'TitleText', 'CountBadge', 'CountText', 'UpdatedText', 'BtnCal', 'BtnHist', 'BtnRefresh', 'BtnCollapse', 'BtnClose', 'Warn', 'Scroll', 'List')) {
        Set-Variable -Name $n -Value $win.FindName($n) -Scope Script
    }

    $Header.Add_MouseLeftButtonDown({
            try { $win.DragMove() } catch {}
            $state.left = $win.Left; $state.top = $win.Top
            Save-State
        })
    $BtnHist.Add_Click({
            if ($state.modo -eq 'historico') { $state.modo = 'fila' } else { $state.modo = 'historico' }
            if ($state.recolhido) { Set-Recolhido $false }
            Save-State
            Render
        })
    $BtnCal.Add_Click({
            if ($state.modo -eq 'calendario') { $state.modo = 'fila' } else { $state.modo = 'calendario' }
            if ($state.recolhido) { Set-Recolhido $false }
            Save-State
            Render
        })
    $BtnRefresh.Add_Click({ Render; Start-Busca 'manual' })
    $BtnCollapse.Add_Click({ Set-Recolhido (-not [bool]$state.recolhido); Save-State })
    $BtnClose.Add_Click({ $win.Close() })

    # Menu do botao direito (as opcoes de copiar so aparecem quando o clique foi num chamado)
    $menu = New-Object System.Windows.Controls.ContextMenu
    $miCopiar = New-Object System.Windows.Controls.MenuItem; $miCopiar.Header = 'Copiar código'
    $miCopiar.Add_Click({ try { Set-Clipboard -Value ([string]$script:alvo.codigo) } catch {} }); [void]$menu.Items.Add($miCopiar)
    $miCopiarT = New-Object System.Windows.Controls.MenuItem; $miCopiarT.Header = 'Copiar código e título'
    $miCopiarT.Add_Click({ try { Set-Clipboard -Value "$($script:alvo.codigo) - $($script:alvo.titulo)" } catch {} }); [void]$menu.Items.Add($miCopiarT)
    $sepCopiar = New-Object System.Windows.Controls.Separator; [void]$menu.Items.Add($sepCopiar)
    $Root.Add_PreviewMouseRightButtonDown({ $script:alvo = $null })
    $menu.Add_Opened({
            if ($script:alvo) { $v = 'Visible' } else { $v = 'Collapsed' }
            foreach ($i in @($miCopiar, $miCopiarT, $sepCopiar)) { $i.Visibility = $v }
        })
    $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = 'Abrir minha fila na Extranet'
    $mi.Add_Click({ if ($ExtranetDominio) { Start-Process $ListaUrl } else { Start-Busca 'manual' } }); [void]$menu.Items.Add($mi)   # sem endereco: pede
    $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = 'Trocar endereço da Extranet'
    $mi.Add_Click({ if (-not $script:busca -and (Pedir-Extranet)) { Start-Busca 'manual' } }); [void]$menu.Items.Add($mi)
    $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = 'Abrir histórico no Excel'
    $mi.Add_Click({ Open-HistoricoExcel }); [void]$menu.Items.Add($mi)
    $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = 'Recarregar da Extranet'
    $mi.Add_Click({ Render; Start-Busca 'manual' }); [void]$menu.Items.Add($mi)
    $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = 'Voltar para o canto da tela'
    $mi.Add_Click({ $state.left = $null; $state.top = $null; Save-State; Place-Window }); [void]$menu.Items.Add($mi)
    $script:miStartup = New-Object System.Windows.Controls.MenuItem; $miStartup.Header = 'Iniciar com o Windows'
    $miStartup.IsChecked = (Test-Path -LiteralPath $StartupLnk)
    $miStartup.Add_Click({ Toggle-Startup }); [void]$menu.Items.Add($miStartup)
    [void]$menu.Items.Add((New-Object System.Windows.Controls.Separator))
    $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = "mininet v$Versao"; $mi.IsEnabled = $false
    [void]$menu.Items.Add($mi)
    $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = 'Fechar'
    $mi.Add_Click({ $win.Close() }); [void]$menu.Items.Add($mi)
    $Root.ContextMenu = $menu

    $win.Add_SizeChanged({ Place-Window })

    # Confere o arquivo a cada minuto; redesenha quando muda (ou a cada 30 min)
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [timespan]::FromSeconds(60)
    $timer.Add_Tick({
            try {
                $mudou = (Test-Path -LiteralPath $DataFile) -and ((Get-Item -LiteralPath $DataFile).LastWriteTimeUtc -ne $script:lastWrite)
                if ($mudou -or ((Get-Date) - $script:lastRender).TotalMinutes -ge 30) { Render }
                # Extranet: desiste apos 90 s sem resposta; com sessao aberta, busca de novo a cada 15 min
                if ($script:busca -and -not $script:esperandoLogin -and ((Get-Date) - $script:buscaDesde).TotalSeconds -ge 90) {
                    Stop-Busca 'A Extranet não respondeu: a fila mostra a última leitura.'
                } elseif (-not $script:busca -and $script:sessaoOk -and ((Get-Date) - $script:ultimaBusca).TotalMinutes -ge 15) { Start-Busca 'auto' }
            } catch {}
        })
    $win.Add_ContentRendered({ try { Start-Busca 'auto' } catch {} })   # tenta com a sessao que ja existir

    Set-Recolhido ([bool]$state.recolhido)
    Render
    Place-Window
    $timer.Start()
    [void]$win.ShowDialog()
    Save-State
} catch {
    [void][System.Windows.MessageBox]::Show("O mininet encontrou um erro:`n`n$($_.Exception.Message)`n`nLinha: $($_.InvocationInfo.ScriptLineNumber)", 'mininet')
} finally {
    Remove-PlanilhaHistorico
    $mutex.ReleaseMutex()
}
