# Autoteste da FilaChamados.ps1 (datas, calendario e leitor da lista da Extranet); nao abre o painel
# Uso: powershell -NoProfile -ExecutionPolicy Bypass -File Testar.ps1
$ErrorActionPreference = 'Stop'
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'FilaChamados.ps1'), [ref]$null, [ref]$null)
foreach ($f in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) { . ([scriptblock]::Create($f.Extent.Text)) }
$ChamadoUrl = 'https://x/Chamado.aspx?CODIGO='
$ExtranetDominio = 'extranet.exemplo.com.br'
$TextInfo = (Get-Culture).TextInfo
function Confere($ok, [string]$msg) { if (-not $ok) { throw "FALHOU: $msg" } }

# ---------------------------------------------------------------- datas, tipos e responsavel
Confere ((Get-TipoCategoria 'APP (PRODUTO - INCIDÊNCIA)') -eq 'Incidência' -and (Get-TipoCategoria 'SISTEMA WEB - Consultoria') -eq 'Consultoria') 'tipo pela categoria'
Confere ((Get-DataBr '15/09/2026 07:32') -eq [datetime]'2026-09-15 07:32') 'data com hora'
Confere ($null -eq (Get-DataBr '')) 'data vazia'
Confere ((Get-ComQuem 'FILA DE ATUALIZAÇÃO') -eq 'na fila de atualização') 'responsavel fila'
Confere ((Get-ComQuem 'FULANA BELTRANA DE TAL') -eq 'com Fulana Beltrana') 'responsavel pessoa'
Confere ((Get-Horas ' 2,68 ') -eq 2.68 -and $null -eq (Get-Horas '') -and $null -eq (Get-Horas '0,00') -and $null -eq (Get-Horas 'x')) 'horas: virgula decimal; vazio, zero e ilegivel sem valor'
Confere ((Get-TextoHoras '1,50') -eq '1,50 h' -and (Get-TextoHoras '') -eq '-') 'horas: texto'
Confere ((Get-TipHoras ([pscustomobject]@{ horasEstimadas = '1,50'; horasRealizadas = '' })) -match '1,50 h estimadas, - realizadas' -and (Get-TipHoras ([pscustomobject]@{ codigo = '1' })) -eq '') 'horas: dica so quando ha horas'

# ---------------------------------------------------------------- seguranca
Confere ((Get-TextoCsv '=HYPERLINK("x")') -eq "'=HYPERLINK(`"x`")" -and (Get-TextoCsv '@SUM(1)') -eq "'@SUM(1)" -and (Get-TextoCsv 'Erro ao salvar') -eq 'Erro ao salvar') 'csv: texto com cara de formula vira texto'
Confere ((Test-UrlExtranet 'https://www.extranet.exemplo.com.br/Chamado.aspx?CODIGO=1') -and (Test-UrlExtranet 'https://extranet.exemplo.com.br/')) 'url: Extranet aceita'
Confere (-not (Test-UrlExtranet 'http://www.extranet.exemplo.com.br/') -and -not (Test-UrlExtranet 'https://extranet.exemplo.com.br.outro.com/') -and -not (Test-UrlExtranet 'https://xextranet.exemplo.com.br/') -and -not (Test-UrlExtranet 'about:blank') -and -not (Test-UrlExtranet '')) 'url: outro site, http e vazio recusados'
Confere ((Get-ListaUrl ' https://www.extranet.exemplo.com.br/sistema/Login.aspx?ReturnUrl=%2f ') -eq 'https://www.extranet.exemplo.com.br/sistema/Chamados.aspx?TipoBusca=lnkResponsabilidade' -and (Get-ListaUrl 'https://extranet.exemplo.com.br/sistema') -eq 'https://extranet.exemplo.com.br/sistema/Chamados.aspx?TipoBusca=lnkResponsabilidade') 'endereco: qualquer pagina da Extranet vira a lista'
Confere ($null -eq (Get-ListaUrl 'http://extranet.exemplo.com.br/') -and $null -eq (Get-ListaUrl 'extranet') -and $null -eq (Get-ListaUrl '')) 'endereco: sem https recusado'
Set-Extranet 'https://www.extranet.exemplo.com.br/sistema/Chamados.aspx?TipoBusca=lnkResponsabilidade'
Confere ($ChamadoUrl -eq 'https://www.extranet.exemplo.com.br/sistema/Chamado.aspx?CODIGO=' -and $ExtranetDominio -eq 'extranet.exemplo.com.br') 'endereco: chamado e dominio saem da lista'
$teste = Join-Path $env:TEMP 'FilaChamadosTeste-dpapi.json'
Write-Json ([pscustomobject]@{ cliente = 'CLIENTE EXEMPLO A'; dias = @(1, 2) }) 3 $teste
$lido = Read-Json $teste
Confere ((Get-Content -LiteralPath $teste -Raw).StartsWith('DPAPI:') -and (Get-Content -LiteralPath $teste -Raw) -notmatch 'CLIENTE' -and $lido.cliente -eq 'CLIENTE EXEMPLO A' -and @($lido.dias).Count -eq 2) 'dpapi: grava cifrado e le de volta'
Set-Content -LiteralPath $teste -Value '{ "cliente": "aberto" }' -Encoding UTF8
Confere ((Read-Json $teste).cliente -eq 'aberto') 'dpapi: le .json aberto de versao antiga'
Remove-Item -LiteralPath $teste
$planilha = Join-Path $env:TEMP 'FilaChamadosTeste-historico.csv'
Export-HistoricoCsv @([pscustomobject]@{ codigo = '1'; saiuEm = '2026-09-24T17:55:00.000Z'; titulo = '=1+1'; cliente = 'CLIENTE EXEMPLO A'; tipo = 'Incidência'; diasAberto = 3; url = 'https://x/1' }) $planilha
$csv = Get-Content -LiteralPath $planilha -Raw -Encoding UTF8
Confere ($csv -match "'=1\+1" -and $csv -match 'Incidência' -and $csv -match 'Título') 'planilha: gerada na hora, sem formula e com acentos'
Remove-Item -LiteralPath $planilha

# ---------------------------------------------------------------- historico
$HistFile = Join-Path $env:TEMP 'FilaChamadosTeste-historico.json'
Remove-Item -LiteralPath $HistFile -ErrorAction SilentlyContinue
$antes = @(
    [pscustomobject]@{ codigo = '1'; titulo = 'fica'; horasEstimadas = '1,00'; horasRealizadas = '0,50' },
    [pscustomobject]@{ codigo = '2'; titulo = 'saiu com leitura'; horasEstimadas = '2,00'; horasRealizadas = '1,00' },
    [pscustomobject]@{ codigo = '3'; titulo = 'saiu sem leitura'; horasEstimadas = '3,00'; horasRealizadas = '2,50' })
Update-Historico $antes @($antes[0]) @([pscustomobject]@{ codigo = '2'; horasEstimadas = '2,00'; horasRealizadas = '3,25' }) '2026-09-24T17:55:00.000Z'
$h = Get-Historico
Confere ($h.Count -eq 2) "historico: os dois que sairam (veio $($h.Count))"
$h2 = $h | Where-Object { $_.codigo -eq '2' }; $h3 = $h | Where-Object { $_.codigo -eq '3' }
Confere ($h2.horasRealizadas -eq '3,25' -and $h2.titulo -eq 'saiu com leitura') 'historico: horas lidas na saida da fila'
Confere ($h3.horasEstimadas -eq '3,00' -and $h3.horasRealizadas -eq '2,50') 'historico: sem leitura na saida, valem as horas da ultima leitura na fila'
Confere ((Get-Content -LiteralPath $HistFile -Raw).StartsWith('DPAPI:')) 'historico: gravado cifrado, com as horas dentro'
Export-HistoricoCsv @([pscustomobject]@{ codigo = '2'; saiuEm = '2026-09-24T17:55:00.000Z'; titulo = 'saiu com leitura'; horasEstimadas = '=2+2'; horasRealizadas = '3,25' }) $planilha
$csv = Get-Content -LiteralPath $planilha -Raw -Encoding UTF8
Confere ($csv -match 'Horas DEV estimadas";"Horas DEV realizadas' -and $csv -match '"3,25"' -and $csv -match "'=2\+2") 'planilha: colunas de horas, sem formula'
Remove-Item -LiteralPath $planilha
$state = [pscustomobject]@{ usuario = 'ANA LIMA (EQUIPE 2)' }
Update-HorasHistorico @([pscustomobject]@{ codigo = '3'; horasEstimadas = '3,00'; horasRealizadas = '5,75'; transferencias = @([pscustomobject]@{ dia = '05/09/2026'; para = 'ANA LIMA SOUZA' }) }, [pscustomobject]@{ codigo = '9'; horasEstimadas = '1,00'; horasRealizadas = '1,00' })
$h = Get-Historico
$h2 = $h | Where-Object { $_.codigo -eq '2' }; $h3 = $h | Where-Object { $_.codigo -eq '3' }
Confere ($h.Count -eq 2 -and $h3.horasRealizadas -eq '5,75' -and $h3.horasLidasEm -and $h3.titulo -eq 'saiu sem leitura') 'historico: pagina lida depois atualiza as horas de quem ja saiu e marca como lida'
Confere ($h3.chegou -eq '05/09/2026') "historico: pagina lida depois traz o dia em que o chamado chegou para o usuario logado (veio '$($h3.chegou)')"
Confere ($h2.horasRealizadas -eq '3,25' -and -not $h2.horasLidasEm) 'historico: quem nao teve a pagina lida fica como estava'
Confere ((Get-Codigos @('12', $null, 'x', 34)) -eq '12,34') 'codigos para o leitor: so numeros'
Remove-Item -LiteralPath $HistFile

# ---------------------------------------------------------------- relatorio de horas
$det = @{ '1' = [pscustomobject]@{ cliente = 'C1'; titulo = 'T1'; tipo = 'Incidência'; criacao = '01/09/2026' } }
$pags = @(
    [pscustomobject]@{ codigo = '1'; horasEstimadas = '1,50'; horasRealizadas = '2,68'; homologacoes = @('20/09/2026', '08/10/2026', '08/10/2026') },   # entra
    [pscustomobject]@{ codigo = '2'; horasEstimadas = '4,00'; horasRealizadas = '2,00'; homologacoes = @('08/10/2026') },                             # realizado abaixo do estimado
    [pscustomobject]@{ codigo = '3'; horasEstimadas = '1,00'; horasRealizadas = '2,00'; homologacoes = @('30/09/2026') },                             # homologacao fora do periodo
    [pscustomobject]@{ codigo = '4'; horasEstimadas = '1,00'; horasRealizadas = '2,00'; homologacoes = @() },                                         # nunca foi para homologacao
    [pscustomobject]@{ codigo = '5'; horasEstimadas = '0,00'; horasRealizadas = '9,73'; homologacoes = @('31/10/2026') })                             # sem estimativa: entra
$l = @(Get-LinhasRelatorio $pags $det ([datetime]'2026-10-01') ([datetime]'2026-10-31'))
Confere ((@($l | ForEach-Object { $_.codigo }) -join ',') -eq '1,5') "relatorio: so realizado acima do estimado com homologacao no periodo (vieram $(@($l | ForEach-Object { $_.codigo }) -join ','))"
Confere ([Math]::Round($l[0].acima, 2) -eq 1.18 -and $l[0].homologacao -eq '08/10/2026' -and $l[0].titulo -eq 'T1' -and $l[0].url -eq ($ChamadoUrl + '1')) 'relatorio: diferenca, dias do periodo sem repetir, detalhes'
Confere ($l[1].estimadas -eq 0 -and $l[1].realizadas -eq 9.73 -and $l[1].titulo -eq '') 'relatorio: chamado sem estimativa entra; sem detalhes fica em branco'
$l[0].titulo = 'Erro em <b>A & B</b><script>alert(1)</script>'
$html = Get-HtmlRelatorio $l ([datetime]'2026-10-01') ([datetime]'2026-10-31') 'Faltou o 9.'
Confere ($html -match 'de 01/10/2026 a 31/10/2026' -and $html -match '<b>2</b>chamados' -and $html -match '<b>1,50 h</b>estimadas' -and $html -match '<b class="x">\+10,91 h</b>') 'relatorio em HTML: periodo e totais'
Confere ($html -match ('<a href="' + [regex]::Escape($ChamadoUrl) + '1">1</a>') -and $html -match '<td class="n x">\+1,18</td>' -and $html -match 'Faltou o 9\.') 'relatorio em HTML: link do chamado, horas com virgula, aviso'
Confere ($html -match 'Erro em &lt;b&gt;A &amp; B&lt;/b&gt;&lt;script&gt;' -and $html -notmatch '<script') 'relatorio em HTML: titulo do chamado entra como texto, nunca como HTML ou script'
$RelDir = Join-Path $env:TEMP 'FilaChamadosTeste-relatorios'
[void](New-Item -ItemType Directory -Force -Path $RelDir)
Set-Content -LiteralPath (Join-Path $RelDir 'relatorio-horas-2026-10-01-a-2026-10-31.pdf') -Value 'pdf'
Set-Content -LiteralPath (Join-Path $RelDir 'outro.txt') -Value 'fica'
Remove-RelatoriosPdf
Confere (-not (Test-Path -LiteralPath (Join-Path $RelDir 'relatorio-horas-2026-10-01-a-2026-10-31.pdf')) -and (Test-Path -LiteralPath (Join-Path $RelDir 'outro.txt'))) 'relatorio: o PDF temporario e apagado, e so ele'
Remove-Item -LiteralPath $RelDir -Recurse -Force

# ---------------------------------------------------------------- calendario
$fila = @([pscustomobject]@{ codigo = '1'; criacao = '23/06/2026' }, [pscustomobject]@{ codigo = '2'; criacao = '' })
$hist = @([pscustomobject]@{ codigo = '3'; criacao = '08/09/2026'; saiuEm = '2026-09-24T17:55:00.000Z' })
$ev = @(Get-EventosCalendario $fila $hist)
Confere ($ev.Count -eq 3) "calendario: 3 eventos (veio $($ev.Count))"
Confere (@($ev | Where-Object { $_.evento -eq 'chegou' -and $_.dia -eq [datetime]'2026-06-23' }).Count -eq 1) 'calendario: chegou pela criacao'
Confere (@($ev | Where-Object { $_.evento -eq 'chegou' -and $_.item.codigo -eq '3' }).Count -eq 1) 'calendario: historico tambem chegou'
Confere (@($ev | Where-Object { $_.evento -eq 'entregue' -and $_.dia -eq (Get-DataLocal '2026-09-24T17:55:00.000Z').Date }).Count -eq 1) 'calendario: entregue pela saida da fila'
$ev = @(Get-EventosCalendario @([pscustomobject]@{ codigo = '4'; criacao = '01/10/2026'; chegou = '07/10/2026' }) @())
Confere ($ev.Count -eq 1 -and $ev[0].dia -eq [datetime]'2026-10-07') 'calendario: chegou pelo dia em que veio para voce, quando a pagina do chamado informou'
$pgc = [pscustomobject]@{ transferencias = @(
        [pscustomobject]@{ dia = '06/10/2026'; para = 'BRUNO COSTA RAMOS' },
        [pscustomobject]@{ dia = '07/10/2026'; para = 'CARLOS EDUARDO PEREIRA DOS SANTOS' },
        [pscustomobject]@{ dia = '08/10/2026'; para = 'DANIELA DO CARMO SANTOS' },
        [pscustomobject]@{ dia = '09/10/2026'; para = 'CARLOS EDUARDO PEREIRA DOS SANTOS' },
        [pscustomobject]@{ dia = '10/10/2026'; para = 'FILA DE TESTE' }) }
Confere ((Get-Chegou $pgc 'CARLOS EDUARDO (EQUIPE 1)') -eq '07/10/2026') 'chegou: primeira vez que o chamado veio para voce (nome curto da lista contra o nome inteiro da interacao)'
Confere ((Get-Chegou $pgc 'carlos santos') -eq '07/10/2026' -and (Get-Chegou $pgc 'DANIELA SANTOS') -eq '08/10/2026') 'chegou: palavras do nome na ordem, sem ligar para maiusculas'
Confere ((Get-Chegou $pgc 'CARLOS EDU') -eq '' -and (Get-Chegou $pgc 'EDUARDO PEREIRA') -eq '' -and (Get-Chegou $pgc '(EQUIPE 1)') -eq '' -and (Get-Chegou ([pscustomobject]@{ codigo = '1' }) 'CARLOS') -eq '') 'chegou: outro nome, nome vazio ou pagina sem trocas de responsavel ficam sem data'
$ExecFile = Join-Path $env:TEMP 'FilaChamadosTeste-execucao.json'
Remove-Item -LiteralPath $ExecFile -ErrorAction SilentlyContinue
$lida = @([pscustomobject]@{ codigo = '1'; emExecucao = $true }, [pscustomobject]@{ codigo = '2'; emExecucao = $false })
foreach ($d in @('2026-09-08 10:00', '2026-09-09 16:00', '2026-09-09 17:00', '2026-09-11 09:00')) { Update-Execucao $lida ([datetime]$d) }
Update-Execucao @([pscustomobject]@{ codigo = '3'; emExecucao = $true }) ([datetime]'2026-09-24 11:00')
$exec = Read-Json $ExecFile
Confere ((@($exec.'1') -join ',') -eq '08/09/2026,09/09/2026,11/09/2026' -and $null -eq $exec.'2') "execucao: um dia por leitura em execucao, sem repetir (veio $(@($exec.'1') -join ','))"
Set-Content -LiteralPath $ExecFile -Value '{ quebrado' -Encoding UTF8
Update-Execucao $lida ([datetime]'2026-09-12 10:00')
Confere ((Get-Content -LiteralPath $ExecFile -Raw) -match 'quebrado') 'execucao: arquivo ilegivel nao e sobrescrito'
Remove-Item -LiteralPath $ExecFile
$per = @(Get-PeriodosGantt $fila $hist $exec)
Confere ($per.Count -eq 2) "gantt: so chamados com dia em execucao (veio $($per.Count))"
Confere ($per[0].item.codigo -eq '1' -and $per[0].trechos.Count -eq 2 -and $per[0].trechos[0].ini -eq [datetime]'2026-09-08' -and $per[0].trechos[0].fim -eq [datetime]'2026-09-09' -and $per[0].trechos[1].ini -eq [datetime]'2026-09-11') 'gantt: dias seguidos viram um trecho'
Confere ($per[1].item.codigo -eq '3' -and $per[1].trechos[0].ini -eq [datetime]'2026-09-24') 'gantt: chamado do historico tambem entra'

# ---------------------------------------------------------------- extranet
# Copia simplificada da lista "Minha responsabilidade" (mesmos ids da pagina real)
# O fetch de mentira faz as vezes da pagina de cada chamado (Chamado.aspx): a do 100002 nao abre (cai no login);
# 100003 saiu da fila, 100004 esta no historico sem horas. A interacao 0 e um comentario que so cita a troca de status
$pagina = @'
<html><head><script>
window.filaAnteriores = [100001, 100003];
window.filaHistorico = [100004];
const paginaChamado = c => `<input id="ContentPlaceHolder1_hddCodigo" value="${c}">
  <input id="ContentPlaceHolder1_txtHorasCodes" value=" 1,50 "><input id="ContentPlaceHolder1_txtHorasCodesRealizada" value="${c === '100003' ? '4,00' : '2,68'}">
  <blockquote id="ContentPlaceHolder1_rptInteracoes_trInteracao_0"><h4>Favor conferir. Alterou Status para: Aguardando Homologação</h4></blockquote>
  <input id="ContentPlaceHolder1_rptInteracoes_hddData_0" value="01/10/2026 09:00:00">
  <blockquote id="ContentPlaceHolder1_rptInteracoes_trInteracao_1"><h4><i class="glyphicon glyphicon-pencil"></i> Alterou Status para: Aguardando Homologação</h4></blockquote>
  <input id="ContentPlaceHolder1_rptInteracoes_hddData_1" value="08/10/2026 14:20:53">
  <blockquote id="ContentPlaceHolder1_rptInteracoes_trInteracao_2"><h4><i class="glyphicon glyphicon-pencil"></i> Alterou Status para: Homologado</h4></blockquote>
  <input id="ContentPlaceHolder1_rptInteracoes_hddData_2" value="08/10/2026 16:38:38">
  <blockquote id="ContentPlaceHolder1_rptInteracoes_trInteracao_3"><div class="int-header"><div class="inline">09/10/2026 08:00:00 - <b class="active">ANA LIMA SOUZA </b></div></div>
    <h4><i class="glyphicon glyphicon-retweet"></i> Resgatou chamado para sua responsabilidade!</h4></blockquote>
  <input id="ContentPlaceHolder1_rptInteracoes_hddData_3" value="09/10/2026 08:00:00">
  <blockquote id="ContentPlaceHolder1_rptInteracoes_trInteracao_4"><div class="int-header"><div class="inline">10/10/2026 09:30:00 - <b class="active">ANA LIMA SOUZA </b><br><span><b>2h </b> úteis em atividade!</span></div></div>
    <h4><i class="glyphicon glyphicon-user"></i> Chamado delegado para: FILA DE TESTE</h4></blockquote>
  <input id="ContentPlaceHolder1_rptInteracoes_hddData_4" value="10/10/2026 09:30:00">`;
window.fetch = async u => { const c = u.split('=').pop(); return { ok: true, text: async () => (c === '100002' ? '<p>Login</p>' : paginaChamado(c)) }; };
</script></head><body><div><b>&nbsp; Lista de Chamados
  (2)</b></div>
<table class="table table-bordered datatable"><tbody><tr><td><a>Código</a></td><td><a>Cliente</a></td><td><a>Status</a></td><td><a>Versão</a></td><td><a>Descritivo</a></td></tr></tbody><tbody>
<tr id="ContentPlaceHolder1_rptChamados_divChamado_0">
 <td id="ContentPlaceHolder1_rptChamados_tdCodigo_0"><input type="hidden" id="ContentPlaceHolder1_rptChamados_hddCodigo_0" value="100001"><h4>100001</h4></td>
 <td id="ContentPlaceHolder1_rptChamados_td1_0"><span>CLIENTE EXEMPLO A - UNIDADE 1</span></td>
 <td id="ContentPlaceHolder1_rptChamados_td2_0"><span>Em Desenvolvimento</span><br> (<a>0</a>) </td>
 <td id="ContentPlaceHolder1_rptChamados_td3_0"><b>9.9.9</b><br> <a>30/07/2026</a></td>
 <td><div class="panel-heading"><input type="hidden" id="ContentPlaceHolder1_rptChamados_hddPrioridade_0" value="1">
  <h5 style="text-transform: uppercase"><div class="tooltip2"><a href="Chamado.aspx?CODIGO=100001"><span>Incidência na abertura</span></a></div></h5>
  <p class="Desktop">Em Desenvolvimento</p> <p class="Desktop">CLIENTE EXEMPLO A</p> APP (PRODUTO - INCIDÊNCIA)<br> Responsável: <b>X</b> - Criação: <b>23/06/2026 13:28:41</b> (<a>94</a>) Tipo Cliente: <b></b>
  <div><input type="hidden" id="ContentPlaceHolder1_rptChamados_HddEmExecucao_0" value="S"></div></div></td>
</tr>
<tr id="ContentPlaceHolder1_rptChamados_divChamado_1">
 <td id="ContentPlaceHolder1_rptChamados_tdCodigo_1"><input type="hidden" id="ContentPlaceHolder1_rptChamados_hddCodigo_1" value="100002"><h4>100002</h4></td>
 <td id="ContentPlaceHolder1_rptChamados_td1_1"><span>MUNICÍPIO EXEMPLO - CLIENTE EXEMPLO B</span></td>
 <td id="ContentPlaceHolder1_rptChamados_td2_1"><span>Novo</span><br> (<a>24</a>) </td>
 <td id="ContentPlaceHolder1_rptChamados_td3_1"><b></b><br> <a></a></td>
 <td><div class="panel-heading"><input type="hidden" id="ContentPlaceHolder1_rptChamados_hddPrioridade_1" value="1">
  <h5><div class="tooltip2"><a href="Chamado.aspx?CODIGO=100002"><span>desenvolvimento front-end</span></a></div></h5>
  <p class="Desktop">Novo</p> <p class="Desktop">CLIENTE EXEMPLO B</p> SISTEMA WEB - Consultoria<br> Responsável: <b>X</b> - Criação: <b>24/08/2026 16:30:07</b> (<a>32</a>)
  <div><input type="hidden" id="ContentPlaceHolder1_rptChamados_HddEmExecucao_1" value="N"></div></div></td>
</tr></tbody></table></body></html>
'@
Add-Type -AssemblyName PresentationFramework
$wv2 = Join-Path $PSScriptRoot 'webview2'
Add-Type -Path (Join-Path $wv2 'Microsoft.Web.WebView2.Core.dll')
Add-Type -Path (Join-Path $wv2 'Microsoft.Web.WebView2.Wpf.dll')
[Microsoft.Web.WebView2.Core.CoreWebView2Environment]::SetLoaderDllFolderPath($wv2)
$script:js = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] -and $args[0].Left.Extent.Text -eq '$ScriptFila' }, $true).Right.Expression.Value

# roda $ScriptFila na pagina dentro de um WebView2 fora da tela e devolve o recado (JSON) que ele mandar;
# $pedido = JavaScript que roda antes (pede o relatorio de horas: ai sao varios recados - a fila, o andamento e, no fim, as paginas)
function Ler-Pagina([string]$h, [string]$pedido = '') {
    $script:html = $h; $script:resp = @(); $script:progressos = 0; $script:pedido = $pedido
    $script:jan = New-Object System.Windows.Window
    $script:jan.Width = 300; $script:jan.Height = 200; $script:jan.Left = -5000; $script:jan.ShowInTaskbar = $false; $script:jan.ShowActivated = $false
    $script:wv = New-Object Microsoft.Web.WebView2.Wpf.WebView2
    $cp = New-Object Microsoft.Web.WebView2.Wpf.CoreWebView2CreationProperties
    $cp.UserDataFolder = Join-Path $env:TEMP 'FilaChamadosTeste'
    $script:wv.CreationProperties = $cp
    $script:wv.Add_CoreWebView2InitializationCompleted({ $script:wv.CoreWebView2.NavigateToString($script:html) })
    $script:wv.Add_NavigationCompleted({ [void]$script:wv.ExecuteScriptAsync("$($script:pedido)`n$($script:js)") })
    $script:wv.Add_WebMessageReceived({
            param($s, $e)
            $m = $e.TryGetWebMessageAsString() | ConvertFrom-Json
            if ($m.progresso) { $script:progressos++; return }   # andamento da leitura em varias paginas: espera a resposta final
            $script:resp += $m
            if (-not $m.ok -or -not $script:pedido -or $m.PSObject.Properties.Name -contains 'relatorio') { $script:jan.Close() }
        })
    $script:jan.Content = $script:wv
    $script:jan.Add_ContentRendered({ [void]$script:wv.EnsureCoreWebView2Async($null) })
    $t = New-Object System.Windows.Threading.DispatcherTimer
    $t.Interval = [timespan]::FromSeconds(30); $t.Add_Tick({ $script:jan.Close() }); $t.Start()
    [void]$script:jan.ShowDialog()
    $t.Stop()
    Confere $script:resp 'o leitor respondeu (WebView2)'
    return $script:resp
}

$r = @(Ler-Pagina $pagina)[0]
Confere ($r.ok -and @($r.chamados).Count -eq 2) "leitor: 2 chamados (erro: $($r.erro))"
$a = $r.chamados[0]; $b = $r.chamados[1]
Confere ($a.codigo -eq '100001' -and $a.status -eq 'Em Desenvolvimento' -and $a.diasStatus -eq 0 -and $a.cliente -eq 'CLIENTE EXEMPLO A - UNIDADE 1') 'leitor: codigo, status, cliente'
Confere ($a.versao -eq '9.9.9' -and $a.dataPrevista -eq '30/07/2026' -and $a.criacao -eq '23/06/2026' -and $a.prioridade -eq '1') 'leitor: versao, prevista, criacao, prioridade'
Confere ($a.titulo -eq 'Incidência na abertura' -and $a.categoria -eq 'APP (PRODUTO - INCIDÊNCIA)' -and $a.emExecucao) "leitor: titulo sem maiusculas do CSS, categoria, em execucao (veio '$($a.titulo)' / '$($a.categoria)')"
Confere ($b.codigo -eq '100002' -and $b.diasStatus -eq 24 -and $b.versao -eq '' -and $b.dataPrevista -eq '' -and -not $b.emExecucao -and $b.categoria -eq 'SISTEMA WEB - Consultoria') 'leitor: chamado sem versao/prevista'
$pg = @($r.detalhes)
Confere ((@($pg | ForEach-Object { $_.codigo }) -join ',') -eq '100001,100003,100004') "leitor: pagina da fila, de quem saiu e do historico sem horas; a que nao abriu fica de fora (vieram $(@($pg | ForEach-Object { $_.codigo }) -join ','))"
Confere ($pg[0].horasEstimadas -eq '1,50' -and $pg[0].horasRealizadas -eq '2,68' -and $pg[1].horasRealizadas -eq '4,00') "leitor: horas da pagina do chamado (veio '$($pg[0].horasEstimadas)' / '$($pg[0].horasRealizadas)')"
Confere ((@($pg[0].homologacoes) -join ',') -eq '08/10/2026') "leitor: dia do Aguardando Homologacao, sem contar o comentario que cita a frase (veio '$(@($pg[0].homologacoes) -join ',')')"
$tr = @($pg[0].transferencias)
Confere ($tr.Count -eq 2 -and $tr[0].dia -eq '09/10/2026' -and $tr[0].para -eq 'ANA LIMA SOUZA' -and $tr[1].dia -eq '10/10/2026' -and $tr[1].para -eq 'FILA DE TESTE') "leitor: trocas de responsavel - quem resgatou e para quem foi delegado (vieram $(@($tr | ForEach-Object { "$($_.dia) $($_.para)" }) -join '; '))"
Confere ($a.responsavel -eq 'X') "leitor: responsavel da linha da lista (veio '$($a.responsavel)')"
$m = @(Ler-Pagina $pagina 'window.filaRelatorio = [100004, 100005];')
Confere ($m.Count -eq 4 -and @($m[0].chamados).Count -eq 2 -and $m[1].andamento -eq '1 de 2' -and $m[2].andamento -eq '2 de 2') "leitor: com relatorio pedido manda a fila, o andamento e as paginas (vieram $($m.Count) recados)"
Confere ((@($m[3].relatorio | ForEach-Object { $_.codigo }) -join ',') -eq '100001,100003,100004,100005') "leitor: relatorio com a pagina de todos os que abriram (vieram $(@($m[3].relatorio | ForEach-Object { $_.codigo }) -join ','))"
$r = @(Ler-Pagina ($pagina.Replace('(2)', '(3)')))[0]
Confere (-not $r.ok -and $r.erro -match '2 de 3' -and $r.html) 'leitor: lista incompleta (paginada) e recusada e manda a pagina'

# ---------------------------------------------------------------- lista em varias paginas
# A Extranet pagina de 20 em 20 e troca de pagina por AJAX (UpdatePanel). O stub imita o PageRequestManager do ASP.NET:
# o __doPostBack do "proxima" troca as linhas pela pagina 2 (chamado 100003), desabilita o "proxima" e dispara o endRequest.
# $modo: 'ok' troca de pagina; 'repete' devolve as mesmas linhas; 'erro' devolve erro no AJAX
function Pagina-Paginada([string]$modo) {
    $ajax = @'
<a id="ContentPlaceHolder1_lnkNext" href="javascript:__doPostBack('ctl00$ContentPlaceHolder1$lnkNext','')">»</a>
<script>
  let hs = [];
  const prm = { add_endRequest: h => hs.push(h), remove_endRequest: h => { hs = hs.filter(x => x !== h); } };
  window.Sys = { WebForms: { PageRequestManager: { getInstance: () => prm } } };
  window.__doPostBack = (alvo) => setTimeout(() => {
    if (alvo !== 'ctl00$ContentPlaceHolder1$lnkNext') return;
    let erro = null;
    if ('MODO' === 'erro') erro = { message: 'timeout do servidor' };
    else if ('MODO' === 'ok') {
      const linhas = [...document.querySelectorAll('tr[id^="ContentPlaceHolder1_rptChamados_divChamado_"]')];
      const n = linhas[1].cloneNode(true);
      n.querySelector('input[id^="ContentPlaceHolder1_rptChamados_hddCodigo_"]').value = '100003';
      const corpo = linhas[0].parentNode;
      linhas.forEach(x => x.remove());
      corpo.appendChild(n);
      const prox = document.getElementById('ContentPlaceHolder1_lnkNext');
      prox.removeAttribute('href'); prox.classList.add('aspNetDisabled');
    }
    hs.slice().forEach(h => h(prm, { get_error: () => erro, set_errorHandled: () => {} }));
  }, 20);
</script></body>
'@
    return $pagina.Replace('(2)', '(3)').Replace('</body>', $ajax.Replace('MODO', $modo))
}
$r = Ler-Pagina (Pagina-Paginada 'ok')
Confere ($r.ok -and $r.paginas -eq 2 -and (@($r.chamados | ForEach-Object { $_.codigo }) -join ',') -eq '100001,100002,100003') "paginas: junta as 2 paginas via AJAX (erro: $($r.erro))"
Confere ($script:progressos -eq 1) "paginas: avisa o andamento entre as paginas (veio $($script:progressos))"
$r = Ler-Pagina (Pagina-Paginada 'repete')
Confere (-not $r.ok -and $r.erro -match 'repetiu' -and $r.html) "paginas: pagina repetida e recusada, sem ficar girando (erro: $($r.erro))"
$r = Ler-Pagina (Pagina-Paginada 'erro')
Confere (-not $r.ok -and $r.erro -match 'timeout do servidor') "paginas: erro no AJAX e recusado (erro: $($r.erro))"
$r = Ler-Pagina ($pagina.Replace('(2)', '(3)').Replace('</body>', '<a id="ContentPlaceHolder1_lnkNext" href="javascript:__doPostBack(''ctl00$ContentPlaceHolder1$lnkNext'','''')">»</a></body>'))
Confere (-not $r.ok -and $r.erro -match 'trocar de página' -and $r.html) "paginas: sem AJAX da Extranet e recusado (erro: $($r.erro))"
$r = Ler-Pagina ($pagina.Replace('(2)', '(0)') -replace '(?s)<tr id="ContentPlaceHolder1_rptChamados_divChamado_.*?</tr>', '')
Confere ($r.ok -and @($r.chamados).Count -eq 0) "fila vazia: aceita (erro: $($r.erro))"

# ---------------------------------------------------------------- filtro por status
$fila = @('Em Análise', 'Novo', 'Em Análise', 'Aguardando Cliente', 'Novo', 'Em Análise') | ForEach-Object { [pscustomobject]@{ status = $_ } }
Confere ((@(Get-GruposStatus $fila | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ',') -eq 'Em Análise=3,Novo=2,Aguardando Cliente=1') 'filtro: status do mais numeroso para o menos'
'OK'
