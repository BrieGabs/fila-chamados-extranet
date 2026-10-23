# mininet - janelinha flutuante com a fila de chamados da Extranet
# Mostra os chamados de chamados.json (mesma pasta). O Recarregar (e a busca
# automatica a cada 15 min) le a fila na Extranet por um navegador WebView2 do
# proprio painel: o login e feito pela pessoa na janelinha que abre, sem senha
# salva. Quando um chamado sai da fila, ele
# vai para o historico (historico.json; a planilha para o Excel e gerada so quando e aberta).
# Os dias em que a leitura ve um chamado "em execucao" ficam em execucao.json (Gantt).
# As horas DEV estimadas e realizadas vem da pagina de cada chamado (Chamado.aspx), na fila e no historico.
# Botao direito, "Relatorio de horas acima do estimado": PDF dos chamados que passaram por Aguardando Homologacao no periodo.
# So mostra chamados que passaram pela sua fila.
# Os .json de dados ficam cifrados com a conta do Windows (DPAPI).
# O endereco da Extranet e pedido no primeiro Recarregar e fica no estado.json (cifrado); nao vem no codigo.
# Clique num chamado para abrir no navegador. Botao direito: mais opcoes.

$ErrorActionPreference = 'Stop'
$Versao = '1.1.0'   # versionamento semantico; mudou aqui, atualize o CHANGELOG.md e crie a tag vX.Y.Z
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$ScriptPath    = $MyInvocation.MyCommand.Path
$Dir           = Split-Path -Parent $ScriptPath
$DataFile      = Join-Path $Dir 'chamados.json'
$StateFile     = Join-Path $Dir 'estado.json'
$HistFile      = Join-Path $Dir 'historico.json'
$HistCsv       = Join-Path $env:LOCALAPPDATA 'FilaChamados\historico.csv'   # temporaria: so existe enquanto o Excel usa
$HistCsvAntigo = Join-Path $Dir 'historico.csv'
$RelDir        = Join-Path $env:LOCALAPPDATA 'FilaChamados'                      # PDFs do relatorio de horas: temporarios como a planilha
$RelPerfil     = Join-Path $env:LOCALAPPDATA 'FilaChamados\WebView2-relatorio'   # navegador que so imprime o relatorio, sem a sessao da Extranet
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
foreach ($p in @('left', 'top', 'recolhido', 'modo', 'ultimoAtualizadoEm', 'ultimos', 'ultimosCodigos', 'codigosBase', 'extranet', 'filtroStatus', 'usuario')) {
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
            @{ n = 'Entrou na minha fila'; e = { $cheg = @($_.chegadas | Where-Object { $_ }); if ($cheg.Count) { Get-TextoCsv $cheg[-1] } else { '' } } },
            @{ n = 'Fora da fila desde'; e = { $d = Get-DataLocal $_.saiuEm; if ($d) { $d.ToString('dd/MM/yyyy') } else { '' } } },
            @{ n = 'Status final'; e = { Get-TextoCsv $_.statusFinal } },
            @{ n = 'Responsável atual'; e = { Get-TextoCsv $_.responsavelAtual } },
            @{ n = 'Último status na minha fila'; e = { Get-TextoCsv $_.ultimoStatus } },
            @{ n = 'Tipo'; e = { Get-TextoCsv $_.tipo } },
            @{ n = 'Cliente'; e = { Get-TextoCsv $_.cliente } },
            @{ n = 'Título'; e = { Get-TextoCsv $_.titulo } },
            @{ n = 'Criado em'; e = { Get-TextoCsv $_.criacao } },
            @{ n = 'Dias em aberto'; e = { $_.diasAberto } },
            @{ n = 'Horas DEV estimadas'; e = { Get-TextoCsv $_.horasEstimadas } },
            @{ n = 'Horas DEV realizadas'; e = { Get-TextoCsv $_.horasRealizadas } },
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
        $horas = $s
        if (-not $s.horasEstimadas -and -not $s.horasRealizadas) { $horas = $det }   # sem leitura das horas na saida: vale a ultima feita na fila
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
                chegadas         = @($det.chegadas | Where-Object { $_ })
                saidas           = @($det.saidas | Where-Object { $_ })
                diasAberto       = [int]$det.diasAberto
                horasEstimadas   = [string]$horas.horasEstimadas
                horasRealizadas  = [string]$horas.horasRealizadas
                url              = $url
            })
        $mudou = $true
    }
    if ($mudou) { Save-Historico $hist }
}

# Idas e vindas do chamado com $nome, pelas trocas de responsavel lidas na pagina do chamado: cada vez que ele veio para $nome
# (resgate feito por ele, ou delegacao para ele) e uma chegada; cada vez que saiu (delegado para outra pessoa ou fila, ou
# resgatado por outra pessoa) e uma saida. Chamado que vai e volta conta mais de uma vez.
# why: a lista escreve o nome curto ("FULANO TAL (EQUIPE 1)") e as interacoes o nome inteiro; valem as palavras do curto, na ordem
function Get-Movimentos($pagina, [string]$nome) {
    $mov = [pscustomobject]@{ chegadas = @(); saidas = @() }
    $p = @(($nome -replace '\(.*?\)', '') -split '\s+' | Where-Object { $_ })
    if ($p.Count -eq 0) { return $mov }
    $re = '^\s*' + (@($p | ForEach-Object { [regex]::Escape($_) }) -join '\s(.*\s)?') + '(\s|$)'
    $comigo = $false
    foreach ($tr in @($pagina.transferencias | Where-Object { $_ -and $_.dia })) {
        $meu = [string]$tr.para -match $re
        if ($meu -and -not $comigo) { $mov.chegadas += [string]$tr.dia; $comigo = $true }
        elseif (-not $meu -and $comigo) { $mov.saidas += [string]$tr.dia; $comigo = $false }
    }
    return $mov
}

# As horas e as idas e vindas lidas agora na pagina do chamado valem tambem para quem ja esta no historico (inclusive quem saiu
# antes de o painel ler isso). horasLidasEm marca a entrada como lida: o leitor so volta a ela quando um relatorio de horas pede o historico inteiro
function Update-HorasHistorico($paginas) {
    $pag = @{}
    foreach ($p in @($paginas | Where-Object { $_ })) { $pag[[string]$p.codigo] = $p }
    $hist = Get-Historico
    $mudou = $false
    foreach ($e in $hist) {
        $p = $pag[[string]$e.codigo]
        if (-not $p) { continue }
        $e | Add-Member -NotePropertyName horasEstimadas -NotePropertyValue ([string]$p.horasEstimadas) -Force
        $e | Add-Member -NotePropertyName horasRealizadas -NotePropertyValue ([string]$p.horasRealizadas) -Force
        $e | Add-Member -NotePropertyName horasLidasEm -NotePropertyValue (Get-Stamp (Get-Date)) -Force
        # sem saber quem e voce (a fila ainda nao foi lida com chamados), as idas e vindas ficam para uma proxima leitura
        if ($state.usuario) {
            $mov = Get-Movimentos $p ([string]$state.usuario)
            $e | Add-Member -NotePropertyName chegadas -NotePropertyValue @($mov.chegadas) -Force
            $e | Add-Member -NotePropertyName saidas -NotePropertyValue @($mov.saidas) -Force
            $e.PSObject.Properties.Remove('chegou')   # campo antigo, so com a primeira chegada
            # importado da busca: o dia em que saiu de voce so aparece agora, na ultima saida
            if (-not $e.saiuEm -and @($mov.saidas).Count) { $d = Get-DataBr ([string]@($mov.saidas)[-1]); if ($d) { $e.saiuEm = Get-Stamp $d } }
        }
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

# horas da Extranet vem como "2,68"; vazio, zero ou ilegivel devolve $null
function Get-Horas([string]$s) {
    $d = 0.0
    if ([double]::TryParse($s.Trim().Replace(',', '.'), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d) -and $d -gt 0) { return $d }
    return $null
}

function Get-TextoHoras([string]$s) {
    if (Get-Horas $s) { return $s.Trim() + ' h' }
    return '-'
}

function Get-TipoCategoria([string]$cat) {
    if ($cat -match 'INCID') { return 'Incidência' }
    if ($cat -match 'NOVO REQUISITO') { return 'Novo requisito' }
    if ($cat -match 'CONSULTORIA') { return 'Consultoria' }
    return 'Outro'
}

# ---------------------------------------------------------------- extranet
# Le a lista "Minha responsabilidade" (Chamados.aspx?TipoBusca=lnkResponsabilidade) e devolve JSON por postMessage.
# A lista vem em paginas de 20; a troca de pagina e AJAX (UpdatePanel do ASP.NET, sem navegar): o leitor le a pagina,
# aciona o "proxima" da propria Extranet, espera o fim do AJAX e le de novo, ate juntar o total que a pagina informa.
# ponytail: depende dos ids do Repeater da Extranet (rptChamados_*, lnkNext); se a pagina mudar, a busca avisa, salva a pagina e nao grava nada
$ScriptFila = @'
(async () => {
  // em falha manda a pagina junto, para ajustar o leitor (vira diagnostico-extranet.html)
  const post = o => window.chrome.webview.postMessage(JSON.stringify(o.ok ? o : { ...o, html: document.documentElement.outerHTML }));
  const progresso = (lidos, total) => window.chrome.webview.postMessage(JSON.stringify({ progresso: true, lidos, total }));
  const t = el => (el ? el.textContent : '').replace(/\s+/g, ' ').trim();
  const MAX_PAGINAS = 50, ESPERA_MS = 20000;
  // ponytail: horas DEV e interacoes so existem na pagina do chamado (Chamado.aspx: campos txtHorasCodes*, Repeater rptInteracoes_*); null = a pagina nao abriu e o painel fica com a leitura anterior
  const lerChamado = async codigo => {
    try {
      const r = await fetch('Chamado.aspx?CODIGO=' + codigo, { credentials: 'same-origin', signal: AbortSignal.timeout(20000) });
      if (!r.ok) return null;
      const doc = new DOMParser().parseFromString(await r.text(), 'text/html');
      const campo = n => { const i = doc.getElementById('ContentPlaceHolder1_' + n); return i ? i.value.trim() : ''; };
      if (campo('hddCodigo') !== codigo) return null;   // login ou pagina de erro no lugar do chamado
      // interacoes na ordem em que aconteceram: o que foi feito (h4), quem fez e o dia (hddData de mesmo numero)
      const interacoes = [...doc.querySelectorAll('[id^="ContentPlaceHolder1_rptInteracoes_trInteracao_"]')].map(b => ({
        acoes: [...b.querySelectorAll('h4')].map(h => t(h)),
        autor: t(b.querySelector('.int-header b')),
        dia: (campo('rptInteracoes_hddData_' + b.id.split('_').pop()).match(/\d{2}\/\d{2}\/\d{4}/) || [''])[0]
      })).filter(i => i.dia);
      // dias em que alguem pos o chamado em "Aguardando Homologacao"
      const homologacoes = interacoes.filter(i => i.acoes.some(a => /^Alterou Status para:\s*Aguardando Homologa/i.test(a))).map(i => i.dia);
      // com quem o chamado ficou a cada troca de responsavel: quem resgatou, ou para quem foi delegado
      const transferencias = [];
      for (const i of interacoes) for (const a of i.acoes) {
        if (/^Resgatou chamado para sua responsabilidade/i.test(a)) transferencias.push({ dia: i.dia, para: i.autor });
        else if (/^Chamado delegado para:/i.test(a)) transferencias.push({ dia: i.dia, para: a.replace(/^Chamado delegado para:\s*/i, '') });
      }
      return { codigo, horasEstimadas: campo('txtHorasCodes'), horasRealizadas: campo('txtHorasCodesRealizada'), homologacoes, transferencias };
    } catch (e) { return null; }
  };
  // Depois da lista: abre a pagina de cada chamado, uma de cada vez - a fila, quem saiu dela desde a leitura anterior
  // (filaAnteriores) e o historico ainda sem horas (filaHistorico).
  // hazard: passou de 45 s, o resto fica para a proxima leitura (a busca inteira desiste aos 90 s sem recado)
  const detalhar = async (chamados, paginas) => {
    const detalhes = [];   // pagina de cada chamado aberta nesta busca
    const abrir = async codigo => {
      if (detalhes.some(p => p.codigo === codigo)) return;
      const p = await lerChamado(codigo);
      if (p) detalhes.push(p);
    };
    const pedidos = n => (window[n] || []).map(String).filter(k => /^\d+$/.test(k));
    const limite = Date.now() + 45000;
    for (const codigo of [...chamados.map(c => c.codigo), ...pedidos('filaAnteriores'), ...pedidos('filaHistorico')]) if (Date.now() < limite) await abrir(codigo);
    post({ ok: true, chamados, paginas, detalhes });
    // relatorio de horas: precisa da pagina de todos (filaRelatorio = historico inteiro), sem limite de tempo; avisa o andamento
    if (!window.filaRelatorio) return;
    const falta = [...new Set([...chamados.map(c => c.codigo), ...pedidos('filaRelatorio')])].filter(k => !detalhes.some(p => p.codigo === k));
    for (const [i, codigo] of falta.entries()) {
      post({ ok: true, andamento: `${i + 1} de ${falta.length}` });
      await abrir(codigo);
    }
    post({ ok: true, relatorio: detalhes });
  };
  // le a pagina que esta na tela
  const ler = () => {
    const total = (t(document.body).match(/Lista de Chamados \((\d+)\)/) || [])[1];
    if (total === undefined) return { erro: 'a página não mostrou a Lista de Chamados' };
    const linhas = [...document.querySelectorAll('tr[id^="ContentPlaceHolder1_rptChamados_divChamado_"]')];
    if (linhas.length === 0 && +total > 0) return { erro: `a página não trouxe nenhum dos ${total} chamados` };
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
        prioridade: oculto('hddPrioridade'),
        responsavel: (t(painel).match(/Respons\S*:\s*(.*?)\s*-\s*Cria/) || [])[1] || ''   // voce, do jeito que a lista escreve: "FULANO TAL (EQUIPE 1)"
      };
    });
    if (chamados.some(c => !/^\d+$/.test(c.codigo) || !c.titulo)) return { erro: 'chamado sem código ou título na lista' };
    // "proxima" habilitado: link com __doPostBack (o desabilitado vem sem href ou com a classe aspNetDisabled)
    const prox = document.getElementById('ContentPlaceHolder1_lnkNext');
    const alvo = prox && !prox.classList.contains('aspNetDisabled') && ((prox.getAttribute('href') || '').match(/__doPostBack\('([^']+)','([^']*)'\)/));
    return { total: +total, chamados, alvo };
  };
  const todos = new Map();
  let paginas = 0;
  const passo = () => {
    try {
      const r = ler();
      if (r.erro) return post({ ok: false, erro: r.erro });
      const antes = todos.size;
      r.chamados.forEach(c => todos.set(c.codigo, c));
      paginas++;
      if (todos.size < r.total && r.alvo) {
        // hazard: pagina que nao trouxe nenhum chamado novo faria o leitor girar sem sair do lugar
        if (paginas > 1 && todos.size === antes) return post({ ok: false, erro: `a página ${paginas} repetiu os chamados da anterior` });
        if (paginas >= MAX_PAGINAS) return post({ ok: false, erro: `a lista passou de ${MAX_PAGINAS} páginas` });
        const prm = window.Sys && Sys.WebForms && Sys.WebForms.PageRequestManager && Sys.WebForms.PageRequestManager.getInstance();
        if (!prm || typeof window.__doPostBack !== 'function') return post({ ok: false, erro: `a lista tem ${r.total} chamados em mais de uma página e não deu para trocar de página` });
        progresso(todos.size, r.total);
        const espera = setTimeout(() => { prm.remove_endRequest(fim); post({ ok: false, erro: `a página ${paginas + 1} da lista não carregou` }); }, ESPERA_MS);
        const fim = (s, a) => {
          prm.remove_endRequest(fim);
          clearTimeout(espera);
          const e = a && a.get_error && a.get_error();
          if (e) { if (a.set_errorHandled) a.set_errorHandled(true); return post({ ok: false, erro: 'erro ao trocar de página: ' + e.message }); }
          setTimeout(passo, 0);
        };
        prm.add_endRequest(fim);
        window.__doPostBack(r.alvo[1], r.alvo[2]);
        return;
      }
      const chamados = [...todos.values()];
      if (chamados.length !== r.total) return post({ ok: false, erro: `as páginas trouxeram ${chamados.length} de ${r.total} chamados` });
      detalhar(chamados, paginas).catch(e => post({ ok: false, erro: 'erro ao abrir as páginas dos chamados: ' + e.message }));
    } catch (e) { post({ ok: false, erro: 'formato da lista mudou: ' + e.message }); }
  };
  passo();
})();
'@

# Importacao para o historico: busca na propria lista da Extranet os chamados em que voce e o desenvolvedor, abertos no periodo.
# Etapa 'buscar': preenche o filtro (Desenvolvedor = voce, Data abertura de/ate) e aciona o Procurar, que recarrega a pagina.
# Etapa 'ler': na pagina do resultado, junta todas as paginas (a troca de pagina e AJAX, pelo "proxima" da propria Extranet).
# ponytail: depende dos ids do filtro (ddlResponsavel, ddlDesenvolvedor, txtDataIni, txtDataFim, lnkProcurar, lnkNext) e das linhas rptChamados_*
$ScriptImportacao = @'
(() => {
  const post = o => window.chrome.webview.postMessage(JSON.stringify(o.ok ? o : { ...o, html: document.documentElement.outerHTML }));
  const progresso = (lidos, total) => window.chrome.webview.postMessage(JSON.stringify({ progresso: true, lidos, total }));
  const t = el => (el ? el.textContent : '').replace(/\s+/g, ' ').trim();
  const pedido = window.filaImportar || {};
  if (pedido.etapa === 'buscar') {
    try {
      const resp = document.getElementById('ContentPlaceHolder1_ddlResponsavel');
      const dev = document.getElementById('ContentPlaceHolder1_ddlDesenvolvedor');
      const ini = document.getElementById('ContentPlaceHolder1_txtDataIni');
      const fim = document.getElementById('ContentPlaceHolder1_txtDataFim');
      if (!resp || !dev || !ini || !fim || typeof window.__doPostBack !== 'function') return post({ ok: false, erro: 'o filtro da lista mudou' });
      // voce = o responsavel que a lista "Minha responsabilidade" ja traz marcado, com o nome inteiro; o filtro de
      // desenvolvedor escreve o nome curto ("FULANO TAL - EQUIPE 1"): valem as palavras do curto, na ordem, no nome inteiro
      const eu = resp.selectedOptions[0] && resp.selectedOptions[0].value ? t(resp.selectedOptions[0]) : '';
      const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const mesmo = o => {
        const p = o.text.replace(/\s+-\s+.*$/, '').trim().split(/\s+/).filter(x => x);
        return p.length > 0 && new RegExp('^\\s*' + p.map(esc).join('\\s(.*\\s)?') + '(\\s|$)', 'i').test(eu);
      };
      const ops = [...dev.options].filter(o => o.value && mesmo(o));
      if (!eu || ops.length !== 1) return post({ ok: false, erro: `não achei você no filtro de desenvolvedor (${ops.length} opções parecidas)` });
      dev.value = ops[0].value;
      resp.value = '';
      ini.value = pedido.ini;
      fim.value = pedido.fim;
      const prm = window.Sys && Sys.WebForms && Sys.WebForms.PageRequestManager && Sys.WebForms.PageRequestManager.getInstance();
      // o Procurar recarrega a pagina; se um dia virar AJAX, o fim do AJAX avisa
      if (prm) { const fimBusca = () => { prm.remove_endRequest(fimBusca); post({ ok: true, buscaPronta: true }); }; prm.add_endRequest(fimBusca); }
      post({ ok: true, buscando: true });
      window.__doPostBack('ctl00$ContentPlaceHolder1$lnkProcurar', '');
    } catch (e) { post({ ok: false, erro: 'o filtro da lista mudou: ' + e.message }); }
    return;
  }
  const MAX_PAGINAS = 50, ESPERA_MS = 20000;
  const ler = () => {
    const total = (t(document.body).match(/Lista de Chamados \((\d+)\)/) || [])[1];
    if (total === undefined) return { erro: 'a página não mostrou a Lista de Chamados' };
    const linhas = [...document.querySelectorAll('tr[id^="ContentPlaceHolder1_rptChamados_divChamado_"]')];
    if (linhas.length === 0 && +total > 0) return { erro: `a página não trouxe nenhum dos ${total} chamados` };
    const chamados = linhas.map(tr => {
      const oculto = n => { const i = tr.querySelector(`input[id^="ContentPlaceHolder1_rptChamados_${n}_"]`); return i ? i.value.trim() : ''; };
      const td = n => tr.querySelector(`td[id^="ContentPlaceHolder1_rptChamados_${n}_"]`);
      const painel = tr.querySelector('.panel-heading');
      return {
        codigo: oculto('hddCodigo'),
        cliente: t(td('td1')),
        status: t(td('td2').querySelector('span')),
        titulo: t(painel.querySelector('h5 a')),
        categoria: t([...painel.childNodes].find(n => n.nodeType === 3 && n.textContent.trim())),
        criacao: (t(painel).match(/Cria\S*:\s*(\d{2}\/\d{2}\/\d{4})/) || [])[1] || '',
        responsavel: (t(painel).match(/Respons\S*:\s*(.*?)\s*-\s*Cria/) || [])[1] || ''
      };
    });
    if (chamados.some(c => !/^\d+$/.test(c.codigo) || !c.titulo)) return { erro: 'chamado sem código ou título na lista' };
    // "proxima" habilitado: link com __doPostBack (o desabilitado vem sem href ou com a classe aspNetDisabled)
    const prox = document.getElementById('ContentPlaceHolder1_lnkNext');
    const alvo = prox && !prox.classList.contains('aspNetDisabled') && ((prox.getAttribute('href') || '').match(/__doPostBack\('([^']+)','([^']*)'\)/));
    return { total: +total, chamados, alvo };
  };
  const todos = new Map();
  let paginas = 0;
  const passo = () => {
    try {
      const r = ler();
      if (r.erro) return post({ ok: false, erro: r.erro });
      const antes = todos.size;
      r.chamados.forEach(c => todos.set(c.codigo, c));
      paginas++;
      if (todos.size < r.total && r.alvo) {
        // hazard: pagina que nao trouxe nenhum chamado novo faria o leitor girar sem sair do lugar
        if (paginas > 1 && todos.size === antes) return post({ ok: false, erro: `a página ${paginas} repetiu os chamados da anterior` });
        if (paginas >= MAX_PAGINAS) return post({ ok: false, erro: `a busca passou de ${MAX_PAGINAS} páginas` });
        const prm = window.Sys && Sys.WebForms && Sys.WebForms.PageRequestManager && Sys.WebForms.PageRequestManager.getInstance();
        if (!prm || typeof window.__doPostBack !== 'function') return post({ ok: false, erro: `a busca trouxe ${r.total} chamados em mais de uma página e não deu para trocar de página` });
        progresso(todos.size, r.total);
        const espera = setTimeout(() => { prm.remove_endRequest(fim); post({ ok: false, erro: `a página ${paginas + 1} da busca não carregou` }); }, ESPERA_MS);
        const fim = (s, a) => {
          prm.remove_endRequest(fim);
          clearTimeout(espera);
          const e = a && a.get_error && a.get_error();
          if (e) { if (a.set_errorHandled) a.set_errorHandled(true); return post({ ok: false, erro: 'erro ao trocar de página: ' + e.message }); }
          setTimeout(passo, 0);
        };
        prm.add_endRequest(fim);
        window.__doPostBack(r.alvo[1], r.alvo[2]);
        return;
      }
      const chamados = [...todos.values()];
      if (chamados.length !== r.total) return post({ ok: false, erro: `as páginas da busca trouxeram ${chamados.length} de ${r.total} chamados` });
      post({ ok: true, importacao: chamados });
    } catch (e) { post({ ok: false, erro: 'formato da busca mudou: ' + e.message }); }
  };
  passo();
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
$script:relatorio = $null      # periodo (@{ ini; fim }) do relatorio de horas pedido, enquanto a busca dele roda
$script:importacao = $null     # @{ ini; etapa = 'buscar' | 'ler' | 'lendo' } enquanto importa para o historico os chamados em que voce e o desenvolvedor
$script:acao = $null           # @{ codigo; iniciar } enquanto o painel inicia ou pausa a execucao de um chamado na Extranet
$script:acaoFeita = $null      # a mesma, depois de acionada: a leitura da fila seguinte confere se a Extranet aceitou
$script:pdf = $null            # navegador a parte (@{ janela; web; html; arq; tarefa; desde }) enquanto imprime o relatorio em PDF
$script:pdfRelogio = $null

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

# "100001,100002" para montar uma lista de codigos no JavaScript do leitor
function Get-Codigos($lista) {
    return @($lista | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^\d+$' }) -join ','
}

function Open-Lista {
    $script:tentativas++
    $script:buscaDesde = Get-Date
    $u = $ListaUrl
    if ($script:acao) { $u = Get-UrlAcao $script:acao }   # iniciar/pausar: o mesmo link do botao da lista da Extranet; depois volta para a lista
    if ($script:web.CoreWebView2) { $script:web.CoreWebView2.Navigate($u) } else { $script:web.Source = [Uri]$u }
}

# Link de iniciar ou pausar a execucao, como o da propria lista da Extranet (Chamado.aspx?CODIGO=...&inicia_execucao=S)
function Get-UrlAcao($acao) {
    if ([string]$acao.codigo -notmatch '^\d+$') { return $null }
    $p = 'pausa_execucao'
    if ($acao.iniciar) { $p = 'inicia_execucao' }
    return "$ChamadoUrl$($acao.codigo)&$p=S"
}

# Botao direito num chamado da fila: inicia ou pausa a execucao dele na Extranet e le a fila de novo
function Start-Acao([string]$codigo, [bool]$iniciar) {
    if ($codigo -notmatch '^\d+$') { return }
    $script:acao = @{ codigo = $codigo; iniciar = $iniciar }
    $script:acaoFeita = $null
    if ($script:busca -and -not $script:esperandoLogin) { $script:busca = $null }   # leitura em andamento: recomeca pela acao
    Start-Busca 'manual'
    if (-not $script:busca) { $script:acao = $null }
}

# Depois da leitura da fila: avisa se o chamado acionado nao ficou como pedido (em execucao ou nao)
function Get-AvisoAcao($ch) {
    $f = $script:acaoFeita
    $script:acaoFeita = $null
    if (-not $f) { return '' }
    $c = @($ch | Where-Object { [string]$_.codigo -eq [string]$f.codigo }) | Select-Object -First 1
    $verbo = 'pausou'
    if ($f.iniciar) { $verbo = 'iniciou' }
    if (-not $c) { return "O chamado $($f.codigo) não está mais na sua fila." }
    if ([bool]$c.emExecucao -ne [bool]$f.iniciar) { return "A Extranet não $verbo a execução do $($f.codigo). Abra o chamado para ver o motivo." }
    return ''
}

function Stop-Busca([string]$aviso) {
    $script:busca = $null
    $script:relatorio = $null
    $script:importacao = $null
    $script:acao = $null
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
    } elseif ($script:acao -and $url -match 'Chamado\.aspx' -and $url -match ('CODIGO=' + $script:acao.codigo + '(&|$)')) {
        # a Extranet abriu o chamado com o iniciar/pausar: segue para a lista, que mostra se ficou em execucao
        Hide-Login
        $script:acaoFeita = $script:acao
        $script:acao = $null
        Open-Lista
    } elseif ($url -match 'Chamados\.aspx') {
        Hide-Login
        if ($script:importacao) { Invoke-Importacao; return }
        # pedidos ao leitor: alem da fila, a pagina de quem estava na leitura anterior (pode ter saido), do historico
        # ainda sem horas e, no relatorio de horas, do historico inteiro
        $hist = Get-Historico
        $pedido = "window.filaAnteriores = [$(Get-Codigos $state.ultimosCodigos)];"
        $semLeitura = @($hist | Where-Object { -not $_.horasLidasEm -or ($state.usuario -and $_.PSObject.Properties.Name -notcontains 'chegadas') })
        $pedido += " window.filaHistorico = [$(Get-Codigos @($semLeitura | ForEach-Object { $_.codigo }))];"
        if ($script:relatorio) { $pedido += " window.filaRelatorio = [$(Get-Codigos @($hist | ForEach-Object { $_.codigo }))];" }
        [void]$script:web.ExecuteScriptAsync("$pedido`n$ScriptFila")
    } elseif ($script:tentativas -lt 3) {
        Hide-Login   # pagina inicial logo depois do login: segue para a fila
        Open-Lista
    } else { Stop-Busca 'A Extranet não abriu a lista de chamados: a fila mostra a última leitura.' }
}

function On-Fila([string]$json) {
    if (-not $script:busca) { return }
    $r = $json | ConvertFrom-Json
    if ($r.progresso) {
        # fila (ou busca da importacao) em varias paginas: mostra o andamento e renova o prazo de 90 s (o prazo e "sem resposta", nao a busca inteira)
        $script:buscaDesde = Get-Date
        if ($script:importacao) { $UpdatedText.Text = "importando… $([int]$r.lidos) de $([int]$r.total)" }
        else { $UpdatedText.Text = "atualizando… $([int]$r.lidos) de $([int]$r.total)" }
        return
    }
    if (-not $r.ok) {
        if ($r.html) { Set-Content -LiteralPath $DiagFile -Value $r.html -Encoding UTF8 }
        Stop-Busca "Não consegui ler a fila na Extranet ($($r.erro)): a fila mostra a última leitura. Cópia da página em diagnostico-extranet.html."
        return
    }
    # recados do relatorio de horas: andamento (renova o prazo de 90 s) e, no fim, as paginas lidas
    if ($r.andamento) { $script:buscaDesde = Get-Date; $UpdatedText.Text = "relatório: lendo $($r.andamento)…"; return }
    if ($r.PSObject.Properties.Name -contains 'relatorio') { Write-Relatorio @($r.relatorio | Where-Object { $_ -and [string]$_.codigo -match '^\d+$' }); return }
    # recados da importacao: nenhum deles e a fila, entao nada aqui pode seguir para a gravacao da fila abaixo
    if ($script:importacao -or $r.buscando -or $r.buscaPronta -or $r.PSObject.Properties.Name -contains 'importacao') {
        if ($r.buscaPronta) { Invoke-Importacao }
        elseif ($script:importacao -and $r.PSObject.Properties.Name -contains 'importacao') { Import-Historico @($r.importacao | Where-Object { $_ -and [string]$_.codigo -match '^\d+$' }) }
        return
    }
    # o codigo vira parte do link que o clique abre: confere de novo aqui, nao so no script da pagina
    if (@($r.chamados | Where-Object { [string]$_.codigo -notmatch '^\d+$' }).Count) {
        Stop-Busca 'Não consegui ler a fila na Extranet (chamado com código inválido): a fila mostra a última leitura.'
        return
    }
    $hoje = (Get-Date).Date
    $pag = @{}     # pagina de cada chamado que o leitor abriu nesta busca
    foreach ($p in @($r.detalhes | Where-Object { $_ -and [string]$_.codigo -match '^\d+$' })) { $pag[[string]$p.codigo] = $p }
    $antes = @{}   # horas e chegada da leitura anterior: continuam valendo quando a pagina do chamado nao abriu agora
    foreach ($a in @($state.ultimos | Where-Object { $_ })) { $antes[[string]$a.codigo] = $a }
    # usuario logado = o responsavel de qualquer linha da lista "Minha responsabilidade"; fica guardado para o historico
    $eu = @($r.chamados | ForEach-Object { [string]$_.responsavel } | Where-Object { $_ }) | Select-Object -First 1
    if ($eu) { $state.usuario = $eu }
    $ch = @(foreach ($c in @($r.chamados)) {
            $cad = Get-DataBr $c.criacao
            $dias = 0
            if ($cad) { $dias = ($hoje - $cad.Date).Days }
            $h = $pag[[string]$c.codigo]
            if ($h) { $mov = Get-Movimentos $h ([string]$c.responsavel) }
            else { $h = $antes[[string]$c.codigo]; $mov = $h }
            [pscustomobject][ordered]@{
                codigo          = [string]$c.codigo
                cliente         = $c.cliente
                status          = $c.status
                diasStatus      = [int]$c.diasStatus
                versao          = $c.versao
                dataPrevista    = $c.dataPrevista
                titulo          = $c.titulo
                categoria       = $c.categoria
                tipo            = Get-TipoCategoria $c.categoria
                criacao         = $c.criacao
                chegadas        = @($mov.chegadas | Where-Object { $_ })
                saidas          = @($mov.saidas | Where-Object { $_ })
                diasAberto      = $dias
                emExecucao      = [bool]$c.emExecucao
                prioridade      = $c.prioridade
                horasEstimadas  = [string]$h.horasEstimadas
                horasRealizadas = [string]$h.horasRealizadas
                url             = $ChamadoUrl + $c.codigo
            }
        })
    # saidos = horas de quem saiu da fila desde a leitura anterior, lidas agora (vao para o historico junto com o chamado)
    $naFila = @($ch | ForEach-Object { $_.codigo })
    $saidos = @($state.ultimosCodigos | ForEach-Object { [string]$_ } | Where-Object { $pag[$_] -and $naFila -notcontains $_ } | ForEach-Object {
            [pscustomobject][ordered]@{ codigo = $_; horasEstimadas = [string]$pag[$_].horasEstimadas; horasRealizadas = [string]$pag[$_].horasRealizadas }
        })
    $novo = [pscustomobject][ordered]@{ atualizadoEm = Get-Stamp (Get-Date); chamados = $ch; saidos = $saidos }
    Write-Json $novo 4 $DataFile
    Remove-Item -LiteralPath $DiagFile -ErrorAction SilentlyContinue   # a copia da pagina tem seus chamados: nao fica depois que o leitor volta a funcionar
    Update-HorasHistorico @($pag.Values)
    $script:sessaoOk = $true
    $script:ultimaBusca = Get-Date
    if ($script:relatorio) { $script:buscaDesde = Get-Date; Render; return }   # a busca segue: o leitor ainda manda as paginas do relatorio
    Stop-Busca (Get-AvisoAcao $ch)
}

# ---------------------------------------------------------------- relatorio de horas
# Entra o chamado com horas DEV realizadas acima das estimadas que teve "Aguardando Homologacao" dentro do periodo.
# why: chamado sem estimativa (0,00) com horas realizadas tambem entra; a coluna das estimadas mostra o zero
function Get-LinhasRelatorio($paginas, $det, [datetime]$ini, [datetime]$fim) {
    foreach ($p in @($paginas | Where-Object { $_ })) {
        $est = [double](Get-Horas ([string]$p.horasEstimadas))
        $real = [double](Get-Horas ([string]$p.horasRealizadas))
        if ($real -le $est) { continue }
        $datas = @($p.homologacoes | ForEach-Object { Get-DataBr ([string]$_) } | Where-Object { $_ -and $_.Date -ge $ini.Date -and $_.Date -le $fim.Date } | Sort-Object)
        if ($datas.Count -eq 0) { continue }
        $k = [string]$p.codigo
        $d = $det[$k]
        [pscustomobject][ordered]@{
            codigo      = $k
            cliente     = [string]$d.cliente
            titulo      = [string]$d.titulo
            tipo        = [string]$d.tipo
            estimadas   = $est
            realizadas  = $real
            acima       = $real - $est
            homologacao = @($datas | ForEach-Object { $_.ToString('dd/MM/yyyy') } | Select-Object -Unique) -join ', '
            criacao     = [string]$d.criacao
            url         = $ChamadoUrl + $k
        }
    }
}

# O relatorio em HTML, de folha clara para imprimir; o navegador do painel transforma em PDF
function Get-HtmlRelatorio($linhas, [datetime]$ini, [datetime]$fim, [string]$aviso) {
    $br = [Globalization.CultureInfo]::GetCultureInfo('pt-BR')
    $e = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }
    $n = { param($v) ([double]$v).ToString('0.00', $br) }
    $tr = @($linhas | ForEach-Object {
            '<tr><td><a href="{0}">{1}</a></td><td>{2}</td><td>{3}</td><td>{4}</td><td class="n">{5}</td><td class="n">{6}</td><td class="n x">+{7}</td><td>{8}</td><td>{9}</td></tr>' -f
            (& $e $_.url), (& $e $_.codigo), (& $e $_.cliente), (& $e $_.titulo), (& $e $_.tipo), (& $n $_.estimadas), (& $n $_.realizadas), (& $n $_.acima), (& $e $_.homologacao), (& $e $_.criacao)
        })
    $soma = { param($p) & $n (@($linhas | ForEach-Object { $_.$p }) | Measure-Object -Sum).Sum }
    return @"
<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Relatório de horas acima do estimado</title><style>
body { font: 10pt 'Segoe UI', Arial, sans-serif; color: #1f2125; background: #fff; margin: 0; -webkit-print-color-adjust: exact; }
h1 { font-size: 15pt; margin: 0 0 2pt; }
.sub { color: #5f6368; margin-bottom: 10pt; }
.res { display: flex; gap: 8pt; margin-bottom: 10pt; }
.res div { border: 1px solid #d0d4da; border-radius: 6pt; padding: 5pt 10pt; color: #5f6368; }
.res b { display: block; font-size: 13pt; color: #1f2125; }
table { width: 100%; border-collapse: collapse; }
thead { display: table-header-group; }
th { background: #1f2125; color: #fff; text-align: left; font-weight: 600; padding: 5pt 6pt; }
td { padding: 4pt 6pt; border-bottom: 1px solid #e3e6ea; vertical-align: top; }
tr { break-inside: avoid; }
.n, td:nth-child(4), td:nth-child(9) { white-space: nowrap; }
.n { text-align: right; }
.res b.x, td.x { color: #b91c1c; font-weight: 600; }
a { color: #1d4ed8; text-decoration: none; }
.obs { color: #5f6368; font-size: 8.5pt; margin-top: 10pt; }
</style></head><body>
<h1>Relatório de horas acima do estimado</h1>
<div class="sub">Chamados com horas DEV realizadas acima das estimadas que tiveram Aguardando Homologação de $($ini.ToString('dd/MM/yyyy')) a $($fim.ToString('dd/MM/yyyy'))</div>
<div class="res"><div><b>$(@($linhas).Count)</b>chamados</div><div><b>$(& $soma 'estimadas') h</b>estimadas</div><div><b>$(& $soma 'realizadas') h</b>realizadas</div><div><b class="x">+$(& $soma 'acima') h</b>acima do estimado</div></div>
<table><thead><tr><th>Código</th><th>Cliente</th><th>Título</th><th>Tipo</th><th class="n">Estimadas</th><th class="n">Realizadas</th><th class="n">Acima</th><th>Aguardando Homologação em</th><th>Criado em</th></tr></thead>
<tbody>
$($tr -join "`n")
</tbody></table>
<div class="obs">Gerado em $((Get-Date).ToString('dd/MM/yyyy HH:mm')) pelo mininet, com os chamados da fila e do histórico do painel. $(& $e $aviso)</div>
</body></html>
"@
}

# paginas = leitura da pagina de cada chamado (fila + historico) que o leitor conseguiu abrir
function Write-Relatorio($paginas) {
    $per = $script:relatorio
    Update-HorasHistorico $paginas
    $hist = Get-Historico
    $det = @{}
    foreach ($c in @($hist) + @($state.ultimos | Where-Object { $_ })) { $det[[string]$c.codigo] = $c }
    $lidos = @($paginas | ForEach-Object { [string]$_.codigo })
    $faltou = @($det.Keys | Where-Object { $lidos -notcontains $_ } | Sort-Object)
    $aviso = ''
    if ($faltou.Count) { $aviso = "Relatório sem os chamados que não abriram na Extranet: $($faltou -join ', ')." }
    $periodo = "$($per.ini.ToString('dd/MM/yyyy')) a $($per.fim.ToString('dd/MM/yyyy'))"
    $linhas = @(Get-LinhasRelatorio $paginas $det $per.ini $per.fim | Sort-Object acima -Descending)
    if ($linhas.Count -eq 0) {
        Stop-Busca ("Relatório de $periodo`: nenhum chamado com horas realizadas acima das estimadas e Aguardando Homologação no período (chamados lidos: $($lidos.Count)). $aviso".Trim())
        return
    }
    $arq = Join-Path $RelDir "relatorio-horas-$($per.ini.ToString('yyyy-MM-dd'))-a-$($per.fim.ToString('yyyy-MM-dd')).pdf"
    $html = Get-HtmlRelatorio $linhas $per.ini $per.fim $aviso
    Stop-Busca $aviso   # a leitura da Extranet acabou aqui; o PDF sai em outro navegador
    try { Start-Pdf $html $arq } catch { Close-Pdf "Não consegui gerar o PDF do relatório: $($_.Exception.Message)" }
}

# why: o PDF e impresso num navegador a parte, com perfil proprio, sem a sessao da Extranet e sem JavaScript; o navegador da
# fila continua abrindo so a Extranet. O arquivo e temporario como a planilha do historico: fica em %LOCALAPPDATA%
# (fora da pasta do painel e do OneDrive) e e apagado quando o painel fecha
function Start-Pdf([string]$html, [string]$arq) {
    Close-Pdf ''
    [void](New-Item -ItemType Directory -Force -Path (Split-Path -Parent $arq))
    $w = New-Object System.Windows.Window
    $w.Width = 400; $w.Height = 300
    $w.Left = -20000; $w.Top = -20000   # fora da tela: so imprime
    $w.ShowInTaskbar = $false; $w.ShowActivated = $false
    $wv = New-Object Microsoft.Web.WebView2.Wpf.WebView2
    $cp = New-Object Microsoft.Web.WebView2.Wpf.CoreWebView2CreationProperties
    $cp.UserDataFolder = $RelPerfil
    $wv.CreationProperties = $cp
    $script:pdf = @{ janela = $w; web = $wv; html = $html; arq = $arq; tarefa = $null; desde = Get-Date }
    $wv.Add_CoreWebView2InitializationCompleted({
            param($s, $e)
            try {
                if (-not $e.IsSuccess) { Close-Pdf "Não consegui gerar o PDF do relatório: $($e.InitializationException.Message)"; return }
                $script:pdf.janela.Hide()
                $s.CoreWebView2.Settings.IsScriptEnabled = $false
                $s.NavigateToString($script:pdf.html)
            } catch { Close-Pdf "Não consegui gerar o PDF do relatório: $($_.Exception.Message)" }
        })
    $wv.Add_NavigationCompleted({ param($s, $e) try { Save-Pdf $e } catch { Close-Pdf "Não consegui gerar o PDF do relatório: $($_.Exception.Message)" } })
    $w.Content = $wv
    $w.Add_ContentRendered({ if ($script:pdf) { [void]$script:pdf.web.EnsureCoreWebView2Async($null) } })
    # why: a impressao e assincrona e esperar por ela travaria a janela; um relogio confere ate terminar
    if (-not $script:pdfRelogio) {
        $script:pdfRelogio = New-Object System.Windows.Threading.DispatcherTimer
        $script:pdfRelogio.Interval = [timespan]::FromMilliseconds(200)
        $script:pdfRelogio.Add_Tick({ try { Complete-Pdf } catch { Close-Pdf "Não consegui gerar o PDF do relatório: $($_.Exception.Message)" } })
    }
    $script:pdfRelogio.Start()
    $w.Show()
}

function Save-Pdf($e) {
    if (-not $script:pdf -or $script:pdf.tarefa) { return }
    if (-not $e.IsSuccess) { Close-Pdf "Não consegui montar o relatório ($($e.WebErrorStatus))."; return }
    $core = $script:pdf.web.CoreWebView2
    $cfg = $core.Environment.CreatePrintSettings()
    $cfg.Orientation = 'Landscape'
    $cfg.PageWidth = 8.27; $cfg.PageHeight = 11.69   # A4, em polegadas
    $cfg.ShouldPrintBackgrounds = $true
    $cfg.ShouldPrintHeaderAndFooter = $false
    $script:pdf.tarefa = $core.PrintToPdfAsync($script:pdf.arq, $cfg)
}

# Abre o PDF assim que a impressao termina
function Complete-Pdf {
    $pdf = $script:pdf
    if (-not $pdf) { $script:pdfRelogio.Stop(); return }
    $t = $pdf.tarefa
    if (-not $t -or -not $t.IsCompleted) {
        if (((Get-Date) - $pdf.desde).TotalSeconds -ge 60) { Close-Pdf 'O PDF do relatório não ficou pronto em 60 segundos. Gere de novo.' }
        return
    }
    if ($t.IsFaulted -or -not $t.Result) { Close-Pdf 'Não consegui gravar o PDF do relatório. Se ele está aberto, feche e gere de novo.'; return }
    Close-Pdf ''
    Start-Process $pdf.arq
}

# Fecha o navegador do PDF; com $erro, avisa no painel
function Close-Pdf([string]$erro) {
    if ($script:pdfRelogio) { $script:pdfRelogio.Stop() }
    $pdf = $script:pdf
    $script:pdf = $null
    if ($pdf) { try { $pdf.web.Dispose(); $pdf.janela.Close() } catch {} }
    if ($erro) { $script:avisoBusca = $erro; Render }
}

# os PDFs do relatorio de horas sao temporarios: somem quando o painel fecha (e na abertura seguinte, se um leitor de PDF segurou o arquivo)
function Remove-RelatoriosPdf {
    Get-ChildItem -LiteralPath $RelDir -Filter 'relatorio-horas-*.pdf' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
}
# Proximo passo da importacao, a cada pagina da lista que abre (ou quando a busca avisa que terminou por AJAX)
function Invoke-Importacao {
    $imp = $script:importacao
    $br = [Globalization.CultureInfo]::InvariantCulture
    if ($imp.etapa -eq 'buscar') {
        $imp.etapa = 'ler'
        $pedido = "window.filaImportar = { etapa: 'buscar', ini: '$($imp.ini.ToString('dd/MM/yyyy', $br))', fim: '$((Get-Date).ToString('dd/MM/yyyy', $br))' };"
    } elseif ($imp.etapa -eq 'ler') {
        $imp.etapa = 'lendo'
        $pedido = "window.filaImportar = { etapa: 'ler' };"
    } else { return }
    $script:buscaDesde = Get-Date
    $UpdatedText.Text = 'importando…'
    [void]$script:web.ExecuteScriptAsync("$pedido`n$ScriptImportacao")
}

# Chamados em que voce e o desenvolvedor, achados pela busca da Extranet, entram no historico (menos os que estao na fila
# e os que ja estao la). Status e responsavel vem da propria busca; horas e idas e vindas, nas leituras seguintes
# (filaHistorico, ate 45 s por leitura), e com elas o dia em que o chamado saiu de voce
function Import-Historico($chamados) {
    $hist = Get-Historico
    $naFila = @($state.ultimosCodigos | ForEach-Object { [string]$_ })
    $ja = New-Object System.Collections.Generic.HashSet[string]
    foreach ($e in $hist) { [void]$ja.Add([string]$e.codigo) }
    $hoje = (Get-Date).Date
    $novos = 0
    foreach ($c in @($chamados | Where-Object { $_ -and [string]$_.codigo -match '^\d+$' })) {
        $k = [string]$c.codigo
        if ($naFila -contains $k -or -not $ja.Add($k)) { continue }
        $cad = Get-DataBr ([string]$c.criacao)
        $dias = 0
        if ($cad) { $dias = ($hoje - $cad.Date).Days }
        [void]$hist.Add([pscustomobject][ordered]@{
                codigo           = $k
                saiuEm           = ''
                statusFinal      = [string]$c.status
                responsavelAtual = [string]$c.responsavel
                ultimoStatus     = ''
                tipo             = Get-TipoCategoria ([string]$c.categoria)
                cliente          = [string]$c.cliente
                titulo           = [string]$c.titulo
                categoria        = [string]$c.categoria
                criacao          = [string]$c.criacao
                diasAberto       = $dias
                url              = $ChamadoUrl + $k
                importado        = $true
            })
        $novos++
    }
    if ($novos) { Save-Historico $hist }
    $total = @($chamados).Count
    $periodo = "abertos desde $($script:importacao.ini.ToString('dd/MM/yyyy'))"
    if ($novos) { Stop-Busca "Importei para o histórico $novos de $total chamados em que você é o desenvolvedor ($periodo). As horas e as idas e vindas de cada um são lidas aos poucos, nas próximas leituras." }
    else { Stop-Busca "Nenhum chamado novo para o histórico: os $total chamados em que você é o desenvolvedor ($periodo) já estão na fila ou no histórico." }
}

# Pede confirmacao e comeca a importacao: chamados em que voce e o desenvolvedor, abertos desde o inicio do ano
function Start-Importacao {
    $ini = (Get-Date).Date.AddDays(1 - (Get-Date).DayOfYear)
    $msg = "Buscar na Extranet os chamados em que você é o desenvolvedor, abertos desde $($ini.ToString('dd/MM/yyyy')), e colocar no histórico?`n`nOs que já estão na fila ou no histórico ficam como estão. As horas e as idas e vindas de cada chamado são lidas aos poucos, nas próximas leituras."
    if ([System.Windows.MessageBox]::Show($win, $msg, 'mininet', 'YesNo', 'Question') -ne 'Yes') { return }
    $script:importacao = @{ ini = $ini; etapa = 'buscar' }
    if ($script:busca -and -not $script:esperandoLogin) { $script:busca = $null }   # leitura em andamento: recomeca pela busca
    Start-Busca 'manual'
    if (-not $script:busca) { $script:importacao = $null }
}

function Start-Relatorio([datetime]$ini, [datetime]$fim) {
    $script:relatorio = @{ ini = $ini.Date; fim = $fim.Date }
    if ($script:busca -and -not $script:esperandoLogin) { $script:busca = $null }   # leitura em andamento nao abre o historico: recomeca
    Start-Busca 'manual'
    if (-not $script:busca) { $script:relatorio = $null }
}

$xamlPeriodo = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Relatório de horas" Width="356" SizeToContent="Height"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        Topmost="True" ShowInTaskbar="False" ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="12" xml:lang="pt-BR">
  <Window.Resources>
    <Style x:Key="Chip" TargetType="Button">
      <Setter Property="Foreground" Value="#FFD1D5DB"/>
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="B" Background="#26FFFFFF" CornerRadius="4" Padding="6,1,6,2">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="B" Property="Background" Value="#40FFFFFF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Principal" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="B" Background="#FF3B82F6" CornerRadius="7" Padding="12,4,12,5">
              <ContentPresenter HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="B" Property="Background" Value="#FF2563EB"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="DatePickerTextBox">
      <Setter Property="Foreground" Value="#FFF1F3F5"/>
      <Setter Property="CaretBrush" Value="#FFF1F3F5"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DatePickerTextBox">
            <ScrollViewer x:Name="PART_ContentHost" Margin="3,1,0,1" VerticalAlignment="Center"/>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="DatePicker">
      <Setter Property="Width" Value="124"/>
      <Setter Property="Background" Value="#14FFFFFF"/>
      <Setter Property="BorderBrush" Value="#33FFFFFF"/>
      <Setter Property="Foreground" Value="#FFF1F3F5"/>
    </Style>
  </Window.Resources>
  <Border Margin="8" CornerRadius="10" Background="#F21F2125" BorderBrush="#33FFFFFF" BorderThickness="1">
    <Border.Effect>
      <DropShadowEffect BlurRadius="12" ShadowDepth="2" Opacity="0.4"/>
    </Border.Effect>
    <StackPanel>
      <Grid x:Name="Cabecalho" Background="Transparent" Margin="12,7,6,7" Cursor="SizeAll">
        <TextBlock Text="Relatório de horas" Foreground="#FFF1F3F5" FontWeight="SemiBold" FontSize="13" VerticalAlignment="Center"/>
        <Button Content="&#xE8BB;" ToolTip="Cancelar" IsCancel="True" HorizontalAlignment="Right" Cursor="Arrow"/>
      </Grid>
      <StackPanel Margin="12,0,12,12">
        <TextBlock TextWrapping="Wrap" Foreground="#FF9AA0A6" FontSize="11" Text="Chamados com horas DEV realizadas acima das estimadas que tiveram Aguardando Homologação no período. O relatório abre em PDF."/>
        <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
          <TextBlock Text="De" Foreground="#FFDDE1E6" VerticalAlignment="Center" Margin="0,0,8,0"/>
          <DatePicker x:Name="De"/>
          <TextBlock Text="até" Foreground="#FFDDE1E6" VerticalAlignment="Center" Margin="10,0,8,0"/>
          <DatePicker x:Name="Ate"/>
        </StackPanel>
        <DockPanel Margin="0,12,0,0">
          <Button x:Name="Gerar" DockPanel.Dock="Right" Style="{StaticResource Principal}" Content="Gerar PDF" IsDefault="True"/>
          <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
            <Button x:Name="EsteMes" Style="{StaticResource Chip}" Content="Este mês"/>
            <Button x:Name="MesPassado" Style="{StaticResource Chip}" Content="Mês passado" Margin="6,0,0,0"/>
          </StackPanel>
        </DockPanel>
      </StackPanel>
    </StackPanel>
  </Border>
</Window>
'@

# Janela do periodo, no visual do painel: devolve @{ ini; fim } ou $null (cancelado)
function Show-PeriodoRelatorio {
    $d = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$xamlPeriodo)))
    $d.Owner = $win
    $d.Resources.Add([System.Windows.Controls.Button], $win.Resources[[System.Windows.Controls.Button]])   # botao de icone do painel (o de fechar)
    $d.FindName('Cabecalho').Add_MouseLeftButtonDown({ try { [System.Windows.Window]::GetWindow($this).DragMove() } catch {} })
    $mes = (Get-Date).Date.AddDays(1 - (Get-Date).Day)
    $de = $d.FindName('De'); $ate = $d.FindName('Ate')
    $de.SelectedDate = $mes
    $ate.SelectedDate = (Get-Date).Date
    # os botoes levam o periodo no Tag: o clique nao enxerga as variaveis desta funcao
    $d.FindName('EsteMes').Tag = @($mes, (Get-Date).Date)
    $d.FindName('MesPassado').Tag = @($mes.AddMonths(-1), $mes.AddDays(-1))
    foreach ($n in @('EsteMes', 'MesPassado')) {
        $d.FindName($n).Add_Click({
                $j = [System.Windows.Window]::GetWindow($this)
                $j.FindName('De').SelectedDate = $this.Tag[0]
                $j.FindName('Ate').SelectedDate = $this.Tag[1]
            })
    }
    $d.FindName('Gerar').Add_Click({ [System.Windows.Window]::GetWindow($this).DialogResult = $true })
    if (-not $d.ShowDialog() -or -not $de.SelectedDate -or -not $ate.SelectedDate) { return $null }
    $ini = [datetime]$de.SelectedDate; $fim = [datetime]$ate.SelectedDate
    if ($ini -gt $fim) { $ini, $fim = $fim, $ini }
    return @{ ini = $ini; fim = $fim }
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
$BrushEstouro = New-Brush '#FFFCA5A5'
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
    <!-- barra de rolagem no tema do painel: fina, sem setas, polegar arredondado que clareia no mouse -->
    <Style TargetType="ScrollBar">
      <Setter Property="Width" Value="8"/>
      <Setter Property="MinWidth" Value="8"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Grid Background="Transparent">
              <Track x:Name="PART_Track" IsDirectionReversed="True" Focusable="False">
                <Track.DecreaseRepeatButton>
                  <RepeatButton Command="ScrollBar.PageUpCommand" Opacity="0" Focusable="False" IsTabStop="False"/>
                </Track.DecreaseRepeatButton>
                <Track.IncreaseRepeatButton>
                  <RepeatButton Command="ScrollBar.PageDownCommand" Opacity="0" Focusable="False" IsTabStop="False"/>
                </Track.IncreaseRepeatButton>
                <Track.Thumb>
                  <Thumb>
                    <Thumb.Template>
                      <ControlTemplate TargetType="Thumb">
                        <Border x:Name="P" Background="#33FFFFFF" CornerRadius="3" Margin="2,2,2,2"/>
                        <ControlTemplate.Triggers>
                          <Trigger Property="IsMouseOver" Value="True">
                            <Setter TargetName="P" Property="Background" Value="#66FFFFFF"/>
                          </Trigger>
                          <Trigger Property="IsDragging" Value="True">
                            <Setter TargetName="P" Property="Background" Value="#88FFFFFF"/>
                          </Trigger>
                        </ControlTemplate.Triggers>
                      </ControlTemplate>
                    </Thumb.Template>
                  </Thumb>
                </Track.Thumb>
              </Track>
            </Grid>
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
      <StackPanel x:Name="Corpo">
        <WrapPanel x:Name="Filtros" Visibility="Collapsed" Margin="4,0,4,5"/>
        <ScrollViewer x:Name="Scroll" MaxHeight="520" VerticalScrollBarVisibility="Auto" Margin="6,0,4,7">
          <StackPanel x:Name="List" Margin="0,0,2,0"/>
        </ScrollViewer>
      </StackPanel>
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

# "Estimado 1,50 h  ·  Realizado 2,68 h" (horas DEV da Extranet); o realizado fica vermelho quando passa do estimado
function New-LinhaHoras($c) {
    $est = Get-Horas ([string]$c.horasEstimadas)
    $real = Get-Horas ([string]$c.horasRealizadas)
    if (-not $est -and -not $real) { return $null }
    $t = New-Object System.Windows.Controls.TextBlock
    $t.FontSize = 11
    $t.Margin = T 0 2 0 0
    $r1 = New-Object System.Windows.Documents.Run ("Estimado $(Get-TextoHoras ([string]$c.horasEstimadas))  ·  ")
    $r1.Foreground = $BrushMuted
    $r2 = New-Object System.Windows.Documents.Run ("Realizado $(Get-TextoHoras ([string]$c.horasRealizadas))")
    if ($est -and $real -gt $est) { $r2.Foreground = $BrushEstouro } else { $r2.Foreground = $BrushMuted }
    [void]$t.Inlines.Add($r1)
    [void]$t.Inlines.Add($r2)
    return $t
}

function Get-TipHoras($c) {
    if (-not (Get-Horas ([string]$c.horasEstimadas)) -and -not (Get-Horas ([string]$c.horasRealizadas))) { return '' }
    return "`nHoras DEV: $(Get-TextoHoras ([string]$c.horasEstimadas)) estimadas, $(Get-TextoHoras ([string]$c.horasRealizadas)) realizadas"
}

function New-Row($c, [bool]$novo) {
    $tip = "$($c.codigo) - $($c.titulo)`n$($c.cliente)`n$($c.categoria)"
    if ($c.versao) { $tip += "`nVersão: $($c.versao)" }
    if ($c.dataPrevista) { $tip += "`nPrevisto para $($c.dataPrevista)" }
    $tip += Get-TipHoras $c
    $cheg = @($c.chegadas | Where-Object { $_ })
    if ($cheg.Count) { $tip += "`nCom você desde $($cheg[-1])" }
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
    $lh = New-LinhaHoras $c
    if ($lh) { [void]$sp.Children.Add($lh) }
    $row.Child = $sp
    return $row
}

function New-HistRow($e) {
    $tip = "$($e.codigo) - $($e.titulo)`n$($e.cliente)`nStatus final: $($e.statusFinal)`nResponsável atual: $($e.responsavelAtual)"
    if ($e.ultimoStatus) { $tip += "`nÚltimo status na sua fila: $($e.ultimoStatus)" }
    $cheg = @($e.chegadas | Where-Object { $_ })
    if ($cheg.Count) { $tip += "`nEntrou na sua fila em $($cheg[-1])" }
    $tip += Get-TipHoras $e
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
    $lh = New-LinhaHoras $e
    if ($lh) { [void]$sp.Children.Add($lh) }
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

# status presentes na fila, do mais numeroso para o menos (empate: ordem alfabetica)
function Get-GruposStatus($ch) {
    return @($ch | Group-Object { [string]$_.status } | Sort-Object @{ Expression = 'Count'; Descending = $true }, Name)
}

# why: fila grande (atendimento, suporte, varias areas) fica ilegivel sem recorte; um chip por status, com a quantidade
function Draw-FiltroStatus($ch) {
    $Filtros.Children.Clear()
    $grupos = Get-GruposStatus $ch
    $sel = [string]$state.filtroStatus
    if ($sel -and @($grupos | Where-Object { $_.Name -eq $sel }).Count -eq 0) { $sel = ''; $state.filtroStatus = $null }   # status que saiu da fila: volta para Todos
    if ($grupos.Count -lt 2) { $Filtros.Visibility = 'Collapsed'; return '' }
    foreach ($f in @(, @("Todos ($($ch.Count))", '')) + @($grupos | ForEach-Object { , @("$($_.Name) ($($_.Count))", [string]$_.Name) })) {
        if ($sel -eq $f[1]) { $chip = New-Tag $f[0] '#FF3B82F6' '#FFFFFFFF' } else { $chip = New-Tag $f[0] '#26FFFFFF' '#FFD1D5DB' }
        $chip.Tag = $f[1]
        $chip.Margin = T 2 2 2 2
        $chip.Cursor = [System.Windows.Input.Cursors]::Hand
        $chip.Add_MouseLeftButtonUp({ $state.filtroStatus = [string]$this.Tag; Save-State; Render })
        [void]$Filtros.Children.Add($chip)
    }
    $Filtros.Visibility = 'Visible'
    return $sel
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
        $sel = Draw-FiltroStatus $ch
        if ($sel) {
            $ch = @($ch | Where-Object { [string]$_.status -eq $sel })
            $CountText.Text = "$($ch.Count) de $($CountText.Text)"
        }
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
    if ($hist.Count -gt 60) { [void]$List.Children.Add((New-Vazio "Mostrando os 60 mais recentes de $($hist.Count). Para ver todos, clique com o botão direito e escolha Abrir histórico no Excel.")) }
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

# Cada vez que o chamado veio para voce e um "chegou", e cada vez que saiu de voce e um "entregue": chamado que vai e volta
# conta mais de uma vez. As idas e vindas vem das trocas de responsavel lidas na pagina do chamado (Get-Movimentos).
# Sem nenhuma, "chegou" e a data de criacao e "entregue" e o dia em que o chamado saiu da fila (saiuEm); o saiuEm tambem
# vale quando o chamado saiu da fila sem troca de responsavel (ex.: encerrado), ou antes de a pagina ser lida de novo
function Get-EventosCalendario($ch, $hist) {
    foreach ($c in @(@($ch) + @($hist) | Where-Object { $_ })) {
        $cheg = @($c.chegadas | ForEach-Object { Get-DataBr ([string]$_) } | Where-Object { $_ })
        $sai = @($c.saidas | ForEach-Object { Get-DataBr ([string]$_) } | Where-Object { $_ })
        $semMov = $cheg.Count -eq 0 -and $sai.Count -eq 0
        if ($cheg.Count -eq 0) { $cheg = @(Get-DataBr ([string]$c.criacao) | Where-Object { $_ }) }
        $d = Get-DataLocal $c.saiuEm
        if ($d -and ($semMov -or $cheg.Count -gt $sai.Count)) { $sai += $d }
        foreach ($d in $cheg) { [pscustomobject]@{ dia = $d.Date; evento = 'chegou'; item = $c } }
        foreach ($d in $sai) { [pscustomobject]@{ dia = $d.Date; evento = 'entregue'; item = $c } }
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
    $res.ToolTip = "Chegou = cada vez que o chamado veio para você (resgate ou delegação), inclusive quando volta; sem isso, a data de criação.`nEntregue = cada vez que saiu de você (delegado ou resgatado por outra pessoa); sem isso, o dia em que saiu da sua fila.`nSaldo positivo: a fila cresceu no mês."
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
        [void]$List.Children.Add((New-Secao "Chegaram em $($sel.ToString('dd/MM')) · $($v.chegou.Count)" 'Cada vez que você resgatou o chamado ou ele foi delegado para você'))
        foreach ($c in $v.chegou) { if ($c.saiuEm) { [void]$List.Children.Add((New-HistRow $c)) } else { [void]$List.Children.Add((New-Row $c $false)) } }
    }
    if ($v.entregue.Count) {
        [void]$List.Children.Add((New-Secao "Entregues em $($sel.ToString('dd/MM')) · $($v.entregue.Count)" 'Cada vez que o chamado saiu de você neste dia'))
        foreach ($e in $v.entregue) { if ($e.saiuEm) { [void]$List.Children.Add((New-HistRow $e)) } else { [void]$List.Children.Add((New-Row $e $false)) } }   # pode ter saido e voltado: ainda na fila
    }
}

$script:lastWrite = [datetime]::MinValue
$script:lastRender = [datetime]::MinValue
$script:ultimaTela = ''
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
    # why: com fila longa, a atualizacao automatica (15 min) nao pode jogar a lista de volta para o topo;
    # a posicao so volta ao topo quando muda a tela (fila, historico, calendario, filtro, mes)
    $tela = "$($state.modo)|$($state.filtroStatus)|$($script:calVista)|$($script:calMes)|$($script:calTipo)|$($script:calDia)"
    $rolagem = 0
    if ($tela -eq $script:ultimaTela) { $rolagem = $Scroll.VerticalOffset }
    $script:ultimaTela = $tela
    $Filtros.Visibility = 'Collapsed'
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
    if ($rolagem -gt 0) { $Scroll.UpdateLayout(); $Scroll.ScrollToVerticalOffset($rolagem) }
}

function Set-Recolhido([bool]$v) {
    $state.recolhido = $v
    if ($v) { $Corpo.Visibility = 'Collapsed'; $BtnCollapse.Content = [string][char]0xE70E }
    else    { $Corpo.Visibility = 'Visible';   $BtnCollapse.Content = [string][char]0xE70D }
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
Remove-RelatoriosPdf

try {
    $win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$xaml)))
    foreach ($n in @('Root', 'Header', 'TitleText', 'CountBadge', 'CountText', 'UpdatedText', 'BtnCal', 'BtnHist', 'BtnRefresh', 'BtnCollapse', 'BtnClose', 'Warn', 'Corpo', 'Filtros', 'Scroll', 'List')) {
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
    $miExecutar = New-Object System.Windows.Controls.MenuItem
    $miExecutar.Add_Click({ if ($script:alvo) { Start-Acao ([string]$script:alvo.codigo) (-not [bool]$script:alvo.emExecucao) } }); [void]$menu.Items.Add($miExecutar)
    $sepCopiar = New-Object System.Windows.Controls.Separator; [void]$menu.Items.Add($sepCopiar)
    $Root.Add_PreviewMouseRightButtonDown({ $script:alvo = $null })
    $menu.Add_Opened({
            if ($script:alvo) { $v = 'Visible' } else { $v = 'Collapsed' }
            foreach ($i in @($miCopiar, $miCopiarT, $sepCopiar)) { $i.Visibility = $v }
            # iniciar/pausar so vale para chamado que esta na fila agora (o do historico ja saiu de voce)
            $naFila = $script:alvo -and -not $script:alvo.saiuEm -and @($state.ultimosCodigos | ForEach-Object { [string]$_ }) -contains [string]$script:alvo.codigo
            if ($naFila) {
                if ($script:alvo.emExecucao) { $miExecutar.Header = "Pausar execução do $($script:alvo.codigo)" } else { $miExecutar.Header = "Iniciar execução do $($script:alvo.codigo)" }
                $miExecutar.Visibility = 'Visible'
            } else { $miExecutar.Visibility = 'Collapsed' }
        })
    $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = 'Abrir minha fila na Extranet'
    $mi.Add_Click({ if ($ExtranetDominio) { Start-Process $ListaUrl } else { Start-Busca 'manual' } }); [void]$menu.Items.Add($mi)   # sem endereco: pede
    $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = 'Trocar endereço da Extranet'
    $mi.Add_Click({ if (-not $script:busca -and (Pedir-Extranet)) { Start-Busca 'manual' } }); [void]$menu.Items.Add($mi)
    $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = 'Abrir histórico no Excel'
    $mi.Add_Click({ Open-HistoricoExcel }); [void]$menu.Items.Add($mi)
    $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = 'Relatório de horas acima do estimado…'
    $mi.Add_Click({ $per = Show-PeriodoRelatorio; if ($per) { Start-Relatorio $per.ini $per.fim } }); [void]$menu.Items.Add($mi)
    $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = 'Importar meus chamados da Extranet para o histórico…'
    $mi.Add_Click({ Start-Importacao }); [void]$menu.Items.Add($mi)
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
    Close-Pdf ''
    Remove-RelatoriosPdf
    $mutex.ReleaseMutex()
}
