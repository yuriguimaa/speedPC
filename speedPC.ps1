#Requires -Version 5.1
<#
.SYNOPSIS
    speedPC - Ferramenta gratuita de diagnóstico e manutenção de computadores Windows para empresas.

.DESCRIPTION
    PowerShell puro, sem programas de terceiros. Lê o estado da máquina e o histórico do próprio
    Windows para apontar o problema, a evidência e o que fazer; oferece correções automáticas
    (o técnico escolhe quais aplicar); faz limpeza segura, otimização, reparo e atualizações;
    gera relatório HTML por máquina e pode rodar agendado toda semana.

    Segurança: antes de alterar qualquer coisa cria um ponto de restauração e grava um backup de
    cada configuração alterada em %ProgramData%\speedPC\Backups (menu "Desfazer alterações").

.PARAMETER Auto
    Modo automático (sem perguntas): diagnóstico + limpeza segura + relatório. Não aplica correções.

.PARAMETER PastaRelatorios
    Pasta (local ou de rede, ex.: \\servidor\manutencao$) onde os relatórios serão salvos.

.PARAMETER SomenteRelatorio
    Apenas diagnostica e gera o relatório, sem alterar nada e sem pedir administrador.

.PARAMETER SemLimpeza
    No modo automático, não executa a limpeza.

.PARAMETER Agendar
    Cria a tarefa agendada semanal (roda como SYSTEM) com o modo automático.

.PARAMETER DiaSemana
    Dia da semana do agendamento (Monday..Sunday). Padrão: Wednesday.

.PARAMETER Hora
    Horário do agendamento (HH:mm). Padrão: 12:00.

.PARAMETER RemoverAgendamento
    Remove a tarefa agendada.

.EXAMPLE
    .\speedPC.ps1
    Abre o menu interativo.

.EXAMPLE
    .\speedPC.ps1 -Agendar -PastaRelatorios \\servidor\manutencao$ -DiaSemana Friday -Hora 12:30

.EXAMPLE
    .\speedPC.ps1 -SomenteRelatorio -PastaRelatorios C:\Temp
#>
[CmdletBinding()]
param(
    [switch]$Auto,
    [string]$PastaRelatorios,
    [switch]$SomenteRelatorio,
    [switch]$SemLimpeza,
    [switch]$Agendar,
    [ValidateSet('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')]
    [string]$DiaSemana = 'Wednesday',
    [ValidatePattern('^\d{1,2}:\d{2}$')]
    [string]$Hora = '12:00',
    [switch]$RemoverAgendamento,
    [switch]$Console,
    [switch]$Motor,
    [string]$CapturarTelas
)

Set-StrictMode -Off
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'

#region ===================== Configuração e estado =====================

$Script:Versao = '1.0.0'
$Script:Base = Join-Path $env:ProgramData 'speedPC'
$Script:PastaBackups = Join-Path $Script:Base 'Backups'
$Script:PastaRelLocal = Join-Path $Script:Base 'Relatorios'
$Script:PastaLogs = Join-Path $Script:Base 'Logs'
$Script:NomeTarefa = 'speedPC - Manutencao semanal'
$Script:ModoAuto = [bool]$Auto
$Script:ModoGui = $false          # true quando o motor roda por trás da janela gráfica
$Script:Ponte = $null             # hashtable sincronizada para perguntas motor -> janela
$Script:CaminhoScript = $PSCommandPath

$Script:EhAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)

$Script:Achados = New-Object System.Collections.Generic.List[object]
$Script:Correcoes = [ordered]@{}
$Script:Acoes = New-Object System.Collections.Generic.List[object]
$Script:Alteracoes = New-Object System.Collections.Generic.List[object]
$Script:ArquivoBackup = $null
$Script:PontoRestauracaoFeito = $false
$Script:DiagnosticoFeito = $false
$Script:Dados = @{ Resumo = $null; Tabelas = [ordered]@{}; AtrasoBoot = @{} }

$Script:Sev = @{
    Critico = @{ Ordem = 0; Rotulo = 'CRÍTICO'; Cor = 'Red'; Css = 'critico' }
    Atencao = @{ Ordem = 1; Rotulo = 'ATENÇÃO'; Cor = 'Yellow'; Css = 'atencao' }
    Info    = @{ Ordem = 2; Rotulo = 'INFO'; Cor = 'Cyan'; Css = 'info' }
    OK      = @{ Ordem = 3; Rotulo = 'OK'; Cor = 'Green'; Css = 'ok' }
}
$Script:OrdemCategorias = @('Lentidão', 'Travamentos', 'Câmera e microfone', 'Rede', 'Impressão', 'Sistema', 'Ferramenta')

$Script:GuidPlanos = @{
    Economia    = 'a1841308-3541-4fab-bc81-f71556f20b4a'
    Equilibrado = '381b4222-f694-41f0-9685-ff5bb260df2e'
    Alto        = '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
}

# Códigos de tela azul mais comuns: Nome, causa provável, tipo (define a correção sugerida)
$Script:BugChecks = @{
    'A'        = @('IRQL_NOT_LESS_OR_EQUAL', 'Um driver ou serviço acessou um endereço de memória inválido.', 'Driver')
    'D1'       = @('DRIVER_IRQL_NOT_LESS_OR_EQUAL', 'Um driver (muitas vezes de rede, Wi-Fi ou VPN) acessou memória inválida.', 'Driver')
    '1E'       = @('KMODE_EXCEPTION_NOT_HANDLED', 'Um driver gerou um erro que o Windows não conseguiu tratar.', 'Driver')
    '3B'       = @('SYSTEM_SERVICE_EXCEPTION', 'Erro em driver ou componente do sistema (comum com vídeo e antivírus).', 'Driver')
    '7E'       = @('SYSTEM_THREAD_EXCEPTION_NOT_HANDLED', 'Um driver falhou durante a execução.', 'Driver')
    'C2'       = @('BAD_POOL_CALLER', 'Um driver fez uso inválido da memória do sistema.', 'Driver')
    '139'      = @('KERNEL_SECURITY_CHECK_FAILURE', 'Estrutura do kernel corrompida: driver incompatível ou memória RAM.', 'Driver')
    '13A'      = @('KERNEL_MODE_HEAP_CORRUPTION', 'Um driver corrompeu a memória do kernel.', 'Driver')
    '9F'       = @('DRIVER_POWER_STATE_FAILURE', 'Um driver não respondeu ao suspender, hibernar ou desligar (comum: Wi-Fi, vídeo, dock USB-C).', 'Driver')
    'A0'       = @('INTERNAL_POWER_ERROR', 'Erro no gerenciamento de energia (suspensão/hibernação).', 'Driver')
    '116'      = @('VIDEO_TDR_FAILURE', 'O driver de vídeo travou e não conseguiu se recuperar.', 'Driver')
    '117'      = @('VIDEO_TDR_TIMEOUT_DETECTED', 'O driver de vídeo parou de responder.', 'Driver')
    '119'      = @('VIDEO_SCHEDULER_INTERNAL_ERROR', 'Erro interno do driver/placa de vídeo.', 'Driver')
    '10E'      = @('VIDEO_MEMORY_MANAGEMENT_INTERNAL', 'Erro de memória do driver de vídeo.', 'Driver')
    '113'      = @('VIDEO_DXGKRNL_FATAL_ERROR', 'Erro fatal no subsistema gráfico (driver de vídeo).', 'Driver')
    'EA'       = @('THREAD_STUCK_IN_DEVICE_DRIVER', 'Um driver (normalmente de vídeo) ficou preso em loop.', 'Driver')
    '133'      = @('DPC_WATCHDOG_VIOLATION', 'Um driver demorou demais para responder (comum: SSD ou controladora de armazenamento com firmware/driver antigo).', 'Driver')
    'BE'       = @('ATTEMPTED_WRITE_TO_READONLY_MEMORY', 'Um driver tentou gravar em memória somente leitura.', 'Driver')
    'C4'       = @('DRIVER_VERIFIER_DETECTED_VIOLATION', 'O Verificador de Drivers detectou um driver com erro (desative-o com "verifier /reset" depois).', 'Driver')
    'D5'       = @('DRIVER_PAGE_FAULT_IN_FREED_SPECIAL_POOL', 'Um driver usou memória que já havia sido liberada.', 'Driver')
    'FC'       = @('ATTEMPTED_EXECUTE_OF_NOEXECUTE_MEMORY', 'Tentativa de executar memória protegida (driver ou RAM).', 'Driver')
    '144'      = @('BUGCODE_USB3_DRIVER', 'Erro no driver/controladora USB 3 (docks, adaptadores, HDs externos).', 'Driver')
    '50'       = @('PAGE_FAULT_IN_NONPAGED_AREA', 'Referência a memória inválida: RAM com defeito, driver ou antivírus.', 'Memoria')
    '1A'       = @('MEMORY_MANAGEMENT', 'Erro grave no gerenciamento de memória: geralmente memória RAM com defeito.', 'Memoria')
    '4E'       = @('PFN_LIST_CORRUPT', 'Tabela de páginas de memória corrompida: RAM ou driver.', 'Memoria')
    '19'       = @('BAD_POOL_HEADER', 'Memória do kernel corrompida por driver ou RAM.', 'Memoria')
    '109'      = @('CRITICAL_STRUCTURE_CORRUPTION', 'Estrutura crítica do kernel alterada: RAM, driver ou software de virtualização.', 'Memoria')
    '7A'       = @('KERNEL_DATA_INPAGE_ERROR', 'Falha ao ler dados do disco para a memória: disco com defeito, cabo ou RAM.', 'Disco')
    '77'       = @('KERNEL_STACK_INPAGE_ERROR', 'Falha ao ler o arquivo de paginação: disco ou cabo.', 'Disco')
    'F4'       = @('CRITICAL_OBJECT_TERMINATION', 'Um processo essencial foi encerrado (frequentemente falha de disco).', 'Disco')
    '24'       = @('NTFS_FILE_SYSTEM', 'Erro no sistema de arquivos NTFS.', 'Disco')
    '154'      = @('UNEXPECTED_STORE_EXCEPTION', 'Erro no armazenamento (SSD/HD) ou conflito com antivírus.', 'Disco')
    '7B'       = @('INACCESSIBLE_BOOT_DEVICE', 'O Windows não conseguiu acessar o disco de inicialização.', 'Disco')
    'EF'       = @('CRITICAL_PROCESS_DIED', 'Um processo crítico do Windows parou: arquivos do sistema corrompidos ou disco.', 'Sistema')
    'C000021A' = @('STATUS_SYSTEM_PROCESS_TERMINATED', 'Um processo de sistema falhou (arquivos do sistema corrompidos).', 'Sistema')
    '1C8'      = @('MANUALLY_INITIATED_POWER_BUTTON_HOLD', 'O usuário segurou o botão de energia porque o sistema congelou.', 'Sistema')
    'E2'       = @('MANUALLY_INITIATED_CRASH', 'Travamento forçado manualmente (teclado ou depurador).', 'Sistema')
    '124'      = @('WHEA_UNCORRECTABLE_ERROR', 'Erro de hardware irrecuperável: processador, memória, superaquecimento ou fonte.', 'Hardware')
    '9C'       = @('MACHINE_CHECK_EXCEPTION', 'Erro de hardware detectado pelo processador.', 'Hardware')
    '101'      = @('CLOCK_WATCHDOG_TIMEOUT', 'Um núcleo do processador parou de responder (hardware, BIOS ou superaquecimento).', 'Hardware')
    '7F'       = @('UNEXPECTED_KERNEL_MODE_TRAP', 'Falha de hardware (RAM, superaquecimento) ou driver.', 'Hardware')
}
$Script:AcaoPorTipo = @{
    Driver   = 'Atualize os drivers (vídeo, rede/Wi-Fi, chipset, armazenamento) e o BIOS pelo site do fabricante; desinstale programas instalados pouco antes dos travamentos.'
    Memoria  = 'Rode o Diagnóstico de Memória do Windows; se acusar erro, troque o pente de RAM. Atualize também os drivers.'
    Disco    = 'Verifique o disco (CHKDSK) e a saúde SMART e faça backup. Troque o cabo SATA ou o disco se os erros se repetirem.'
    Sistema  = 'Execute o reparo do Windows (DISM + SFC) e instale as atualizações pendentes.'
    Hardware = 'Verifique temperatura (poeira, ventoinhas, pasta térmica), fonte/carregador e memória; atualize o BIOS. Persistindo, encaminhe para assistência.'
}
$Script:CodigosDispositivo = @{
    1 = 'não configurado corretamente'; 3 = 'driver corrompido'; 10 = 'o dispositivo não pode ser iniciado'
    12 = 'conflito de recursos'; 14 = 'requer reinício'; 18 = 'drivers precisam ser reinstalados'; 19 = 'configuração corrompida no registro'
    21 = 'sendo removido'; 22 = 'desativado'; 24 = 'não presente ou sem driver'; 28 = 'driver não instalado'
    31 = 'o driver não carregou'; 32 = 'driver desativado'; 37 = 'falha ao inicializar o driver'; 39 = 'driver corrompido ou ausente'
    41 = 'driver carregado, mas o dispositivo não foi encontrado'; 43 = 'o Windows parou o dispositivo porque ele relatou problemas'
    48 = 'driver bloqueado por incompatibilidade'; 52 = 'driver sem assinatura digital válida'
}

# Catálogo de aplicativos instaláveis pelo winget (repositório oficial da Microsoft).
# Para adicionar um programa, procure o Id com:  winget search <nome>
$Script:CatalogoApps = @(
    @{ Cat = 'Navegadores'; Apps = @(
            @{ Id = 'Google.Chrome'; Site = 'https://www.google.com/chrome/'; Icone = 'https://www.google.com/chrome/static/images/favicons/favicon-96x96.png'; Nome = 'Google Chrome'; Desc = 'Navegador do Google' }
            @{ Id = 'Mozilla.Firefox'; Site = 'https://www.mozilla.org/firefox/'; Nome = 'Mozilla Firefox'; Desc = 'Navegador da Mozilla' }
            @{ Id = 'Microsoft.Edge'; Site = 'https://www.microsoft.com/edge'; Nome = 'Microsoft Edge'; Desc = 'Navegador da Microsoft' }
            @{ Id = 'Brave.Brave'; Site = 'https://brave.com'; Nome = 'Brave'; Desc = 'Navegador com bloqueio de anúncios' }
            @{ Id = 'Opera.Opera'; Site = 'https://www.opera.com'; Nome = 'Opera'; Desc = 'Navegador com VPN embutida' }
        ) }
    @{ Cat = 'Comunicação e reuniões'; Apps = @(
            @{ Id = 'Microsoft.Teams'; Site = 'https://teams.microsoft.com'; Nome = 'Microsoft Teams'; Desc = 'Chat e reuniões do Microsoft 365' }
            @{ Id = 'Zoom.Zoom'; Site = 'https://zoom.us'; Nome = 'Zoom'; Desc = 'Videoconferência' }
            @{ Id = 'Cisco.Webex'; Site = 'https://www.webex.com'; Nome = 'Webex'; Desc = 'Videoconferência da Cisco' }
            @{ Id = 'SlackTechnologies.Slack'; Site = 'https://slack.com'; Nome = 'Slack'; Desc = 'Chat de equipes' }
            @{ Id = 'Telegram.TelegramDesktop'; Site = 'https://telegram.org'; Nome = 'Telegram'; Desc = 'Mensagens' }
            @{ Id = '9NKSQGP7F2NH'; Site = 'https://www.whatsapp.com'; Nome = 'WhatsApp'; Desc = 'Mensagens (Microsoft Store)'; Fonte = 'msstore' }
        ) }
    @{ Cat = 'Escritório e PDF'; Apps = @(
            @{ Id = 'Adobe.Acrobat.Reader.64-bit'; Site = 'https://acrobat.adobe.com'; Nome = 'Adobe Acrobat Reader'; Desc = 'Leitor de PDF' }
            @{ Id = 'Foxit.FoxitReader'; Site = 'https://www.foxit.com'; Nome = 'Foxit PDF Reader'; Desc = 'Leitor de PDF leve' }
            @{ Id = 'geeksoftwareGmbH.PDF24Creator'; Site = 'https://tools.pdf24.org'; Nome = 'PDF24'; Desc = 'Criar, juntar e comprimir PDFs' }
            @{ Id = 'TheDocumentFoundation.LibreOffice'; Site = ''; Nome = 'LibreOffice'; Desc = 'Pacote de escritório gratuito' }
            @{ Id = 'ONLYOFFICE.DesktopEditors'; Site = 'https://www.onlyoffice.com'; Nome = 'ONLYOFFICE'; Desc = 'Editor compatível com Office' }
            @{ Id = 'Notepad++.Notepad++'; Site = 'https://notepad-plus-plus.org'; Nome = 'Notepad++'; Desc = 'Editor de texto' }
        ) }
    @{ Cat = 'Acesso remoto'; Apps = @(
            @{ Id = 'AnyDesk.AnyDesk'; Site = 'https://anydesk.com'; Nome = 'AnyDesk'; Desc = 'Suporte remoto' }
            @{ Id = 'TeamViewer.TeamViewer'; Site = 'https://www.teamviewer.com'; Nome = 'TeamViewer'; Desc = 'Suporte remoto' }
        ) }
    @{ Cat = 'Utilitários'; Apps = @(
            @{ Id = '7zip.7zip'; Site = 'https://www.7-zip.org'; Nome = '7-Zip'; Desc = 'Compactador de arquivos' }
            @{ Id = 'RARLab.WinRAR'; Site = 'https://www.win-rar.com'; Nome = 'WinRAR'; Desc = 'Compactador de arquivos' }
            @{ Id = 'voidtools.Everything'; Site = 'https://www.voidtools.com'; Nome = 'Everything'; Desc = 'Busca instantânea de arquivos' }
            @{ Id = 'Microsoft.PowerToys'; Site = 'https://learn.microsoft.com/windows/powertoys/'; Nome = 'PowerToys'; Desc = 'Utilitários da Microsoft' }
            @{ Id = 'ShareX.ShareX'; Site = 'https://getsharex.com'; Nome = 'ShareX'; Desc = 'Captura de tela e gravação' }
            @{ Id = 'Greenshot.Greenshot'; Site = 'https://getgreenshot.org'; Nome = 'Greenshot'; Desc = 'Captura de tela' }
            @{ Id = 'VideoLAN.VLC'; Site = 'https://www.videolan.org'; Nome = 'VLC'; Desc = 'Player de vídeo e áudio' }
        ) }
    @{ Cat = 'Diagnóstico de hardware'; Apps = @(
            @{ Id = 'CrystalDewWorld.CrystalDiskInfo'; Site = 'https://crystalmark.info'; Nome = 'CrystalDiskInfo'; Desc = 'Saúde SMART de HD e SSD' }
            @{ Id = 'REALiX.HWiNFO'; Site = 'https://www.hwinfo.com'; Nome = 'HWiNFO'; Desc = 'Sensores, temperaturas e hardware' }
            @{ Id = 'CPUID.CPU-Z'; Site = 'https://www.cpuid.com'; Nome = 'CPU-Z'; Desc = 'Processador, placa-mãe e memória' }
            @{ Id = 'AntibodySoftware.WizTree'; Site = 'https://diskanalyzer.com'; Nome = 'WizTree'; Desc = 'O que está ocupando o disco' }
        ) }
    @{ Cat = 'Rede e TI'; Apps = @(
            @{ Id = 'Famatech.AdvancedIPScanner'; Site = 'https://www.advanced-ip-scanner.com'; Nome = 'Advanced IP Scanner'; Desc = 'Varredura da rede local' }
            @{ Id = 'PuTTY.PuTTY'; Site = 'https://www.chiark.greenend.org.uk/~sgtatham/putty/'; Nome = 'PuTTY'; Desc = 'SSH e Telnet' }
            @{ Id = 'WinSCP.WinSCP'; Site = 'https://winscp.net'; Nome = 'WinSCP'; Desc = 'SFTP e FTP' }
            @{ Id = 'WiresharkFoundation.Wireshark'; Site = 'https://www.wireshark.org'; Nome = 'Wireshark'; Desc = 'Análise de tráfego de rede' }
            @{ Id = 'Microsoft.VisualStudioCode'; Site = 'https://code.visualstudio.com'; Nome = 'VS Code'; Desc = 'Editor de código' }
        ) }
    @{ Cat = 'Segurança e senhas'; Apps = @(
            @{ Id = 'Malwarebytes.Malwarebytes'; Site = 'https://www.malwarebytes.com'; Nome = 'Malwarebytes'; Desc = 'Remoção de malware' }
            @{ Id = 'Bitwarden.Bitwarden'; Site = 'https://bitwarden.com'; Nome = 'Bitwarden'; Desc = 'Gerenciador de senhas' }
            @{ Id = 'KeePassXCTeam.KeePassXC'; Site = 'https://keepassxc.org'; Nome = 'KeePassXC'; Desc = 'Gerenciador de senhas offline' }
        ) }
    @{ Cat = 'Componentes do sistema'; Apps = @(
            @{ Id = 'Microsoft.VCRedist.2015+.x64'; Site = 'https://learn.microsoft.com/cpp/'; Nome = 'Visual C++ 2015-2022 (x64)'; Desc = 'Exigido por muitos programas' }
            @{ Id = 'Microsoft.VCRedist.2015+.x86'; Site = 'https://learn.microsoft.com/cpp/'; Nome = 'Visual C++ 2015-2022 (x86)'; Desc = 'Exigido por programas 32 bits' }
            @{ Id = 'Microsoft.DotNet.DesktopRuntime.8'; Site = 'https://dotnet.microsoft.com'; Nome = '.NET Desktop Runtime 8'; Desc = 'Exigido por programas .NET' }
            @{ Id = 'Oracle.JavaRuntimeEnvironment'; Site = 'https://www.java.com'; Nome = 'Java (JRE)'; Desc = 'Exigido por sistemas legados' }
        ) }
)

#endregion

#region ===================== Utilitários =====================

function Write-Log {
    param([string]$Mensagem)
    try {
        if (-not (Test-Path -LiteralPath $Script:PastaLogs)) { New-Item -ItemType Directory -Path $Script:PastaLogs -Force | Out-Null }
        $arq = Join-Path $Script:PastaLogs ('speedPC_{0:yyyy-MM}.log' -f (Get-Date))
        Add-Content -LiteralPath $arq -Value ('{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $Mensagem) -Encoding UTF8
    } catch { }
}

function Write-Titulo {
    param([string]$Texto)
    if ($Script:ModoGui) { Write-Host ''; Write-Host $Texto -ForegroundColor Cyan; return }
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
    Write-Host "  $Texto" -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
}

function Write-Etapa { param([string]$Texto) Write-Host "  > $Texto" -ForegroundColor DarkGray }

function Format-Bytes {
    param([double]$Bytes)
    if ($Bytes -ge 1TB) { return '{0:N1} TB' -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return '{0:N1} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N0} MB' -f ($Bytes / 1MB) }
    return '{0:N0} KB' -f ($Bytes / 1KB)
}

function Format-Duracao {
    param([TimeSpan]$Tempo)
    if ($Tempo.TotalDays -ge 1) { return '{0} dia(s) e {1} h' -f [math]::Floor($Tempo.TotalDays), $Tempo.Hours }
    if ($Tempo.TotalHours -ge 1) { return '{0} h {1} min' -f [math]::Floor($Tempo.TotalHours), $Tempo.Minutes }
    return '{0} min' -f [math]::Max(0, [math]::Floor($Tempo.TotalMinutes))
}

function Get-Barra {
    param([double]$Pct, [int]$Largura = 24)
    $n = [int][math]::Round([math]::Min(100, [math]::Max(0, $Pct)) / 100 * $Largura)
    return '[' + ('#' * $n) + ('-' * ($Largura - $n)) + ']'
}

function Confirm-SN {
    param([string]$Pergunta, [ValidateSet('S', 'N')][string]$Padrao = 'N')
    if ($Script:ModoAuto) { return ($Padrao -eq 'S') }
    if ($Script:ModoGui) { return [bool](Request-Gui -Tipo 'Confirmar' -Dados @{ Texto = $Pergunta; Padrao = $Padrao }) }
    $sufixo =if ($Padrao -eq 'S') { '(S/n)' } else { '(s/N)' }
    $r = Read-Host "  $Pergunta $sufixo"
    if ([string]::IsNullOrWhiteSpace($r)) { return ($Padrao -eq 'S') }
    return ($r.Trim() -match '^(s|sim|y|yes)$')
}

function Select-Indices {
    param([int]$Total, [string]$Prompt = 'Escolha (ex.: 1,3 ou 1-4; T = todas; Enter = nenhuma)')
    $r = Read-Host "  $Prompt"
    if ([string]::IsNullOrWhiteSpace($r)) { return @() }
    if ($r.Trim() -match '^(t|todas|todos)$') { return @(0..($Total - 1)) }
    $set = New-Object System.Collections.Generic.List[int]
    foreach ($parte in ($r -split '[,; ]+')) {
        $ini = 0; $fim = 0
        if ($parte -match '^(\d+)-(\d+)$') { $ini = [int]$Matches[1]; $fim = [int]$Matches[2] }
        elseif ($parte -match '^\d+$') { $ini = [int]$parte; $fim = $ini }
        else { continue }
        for ($i = $ini; $i -le $fim; $i++) {
            if ($i -ge 1 -and $i -le $Total -and -not $set.Contains($i - 1)) { $set.Add($i - 1) }
        }
    }
    return @($set)
}

function Request-Gui {
    # O motor roda num runspace separado; pede à janela que mostre um diálogo e espera a resposta.
    param([string]$Tipo, [hashtable]$Dados)
    $Script:Ponte.Resposta = $null
    $Script:Ponte.Pedido = @{ Tipo = $Tipo; Dados = $Dados }
    while ($null -eq $Script:Ponte.Resposta) { Start-Sleep -Milliseconds 120 }
    $r = $Script:Ponte.Resposta
    $Script:Ponte.Resposta = $null
    return $r
}

function Select-Opcoes {
    # Lista de escolhas múltiplas. Itens: objetos com Nome e Detalhe. Retorna os índices escolhidos.
    param([string]$Titulo, [object[]]$Itens, [string]$Aviso = '')
    $Itens = @($Itens)
    if (-not $Itens) { return @() }
    if ($Script:ModoGui) { return @(Request-Gui -Tipo 'Selecao' -Dados @{ Titulo = $Titulo; Itens = $Itens; Aviso = $Aviso }) }
    for ($i = 0; $i -lt $Itens.Count; $i++) {
        Write-Host ("  [{0}] {1}" -f ($i + 1), $Itens[$i].Nome) -ForegroundColor White
        if ($Itens[$i].Detalhe) { Write-Host ("      {0}" -f $Itens[$i].Detalhe) -ForegroundColor Gray }
    }
    Write-Host ''
    if ($Aviso) { Write-Host "  $Aviso" -ForegroundColor Yellow }
    return @(Select-Indices -Total $Itens.Count)
}

function Start-Ferramenta {
    # DISM, SFC, CHKDSK, winget: no console mostram o progresso na própria janela;
    # na interface gráfica abrem uma janela própria para o técnico acompanhar.
    param([string]$Arquivo, [string[]]$Argumentos)
    if ($Script:ModoGui) { return Start-Process -FilePath $Arquivo -ArgumentList $Argumentos -Wait -PassThru }
    return Start-Process -FilePath $Arquivo -ArgumentList $Argumentos -NoNewWindow -Wait -PassThru
}

function Get-Eventos {
    param([string]$Log, [int[]]$Ids, [int]$Dias = 30, [string]$Provider, [int]$Max = 500, [int]$NivelMax = 0)
    $f = @{ LogName = $Log; StartTime = (Get-Date).AddDays(-$Dias) }
    if ($Ids) { $f.Id = $Ids }
    if ($Provider) { $f.ProviderName = $Provider }
    if ($NivelMax -gt 0) { $f.Level = 1..$NivelMax }
    try { return @(Get-WinEvent -FilterHashtable $f -MaxEvents $Max -ErrorAction Stop) } catch { return @() }
}

function Get-DadosEvento {
    param($Evento)
    $h = @{}
    try {
        $x = [xml]$Evento.ToXml()
        foreach ($d in $x.Event.EventData.Data) { if ($d.Name) { $h[$d.Name] = $d.'#text' } }
    } catch { }
    return $h
}

function Get-PrimeiraLinha {
    param([string]$Texto, [int]$Max = 180)
    if (-not $Texto) { return '' }
    $l = ($Texto -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -First 1)
    if ($l.Length -gt $Max) { $l = $l.Substring(0, $Max) + '...' }
    return $l.Trim()
}

function Get-HivesUsuarios {
    # Usuários com perfil carregado (logados) - locais, de domínio ou Azure AD
    $lista = @()
    foreach ($k in (Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue)) {
        $sid = $k.PSChildName
        if ($sid -notmatch '^S-1-(5-21|12-1)-[\d-]+$') { continue }
        $nome = $sid
        try { $nome = (New-Object Security.Principal.SecurityIdentifier $sid).Translate([Security.Principal.NTAccount]).Value } catch { }
        $perfil = (Get-ItemProperty "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -ErrorAction SilentlyContinue).ProfileImagePath
        $lista += [pscustomobject]@{ Sid = $sid; Usuario = $nome; Perfil = $perfil }
    }
    return $lista
}

function Get-PerfisUsuarios {
    if (-not $Script:EhAdmin) { return @([pscustomobject]@{ Caminho = $env:USERPROFILE }) }
    $perfis = Get-CimInstance Win32_UserProfile -ErrorAction SilentlyContinue |
        Where-Object { -not $_.Special -and $_.LocalPath -and (Test-Path -LiteralPath $_.LocalPath) }
    return @($perfis | ForEach-Object { [pscustomobject]@{ Caminho = $_.LocalPath } })
}

function Test-PortaTcp {
    param([string]$Destino, [int]$Porta, [int]$TimeoutMs = 3000)
    $c = New-Object Net.Sockets.TcpClient
    try {
        $iar = $c.BeginConnect($Destino, $Porta, $null, $null)
        $ok = $iar.AsyncWaitHandle.WaitOne($TimeoutMs)
        return ($ok -and $c.Connected)
    } catch { return $false } finally { $c.Close() }
}

function Get-PlanoEnergiaAtivo {
    $s = (powercfg /getactivescheme) 2>$null | Out-String
    $m = [regex]::Match($s, '[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}')
    $nome = ''
    if ($s -match '\(([^)]+)\)') { $nome = $Matches[1] }
    return [pscustomobject]@{ Guid = $m.Value.ToLower(); Nome = $nome }
}

function Test-ReinicioPendente {
    $chaves = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    )
    foreach ($c in $chaves) { if (Test-Path $c) { return $true } }
    return $false
}

function Add-Achado {
    param(
        [string]$Categoria,
        [ValidateSet('Critico', 'Atencao', 'Info', 'OK')][string]$Severidade,
        [string]$Titulo,
        [string]$Evidencia = '',
        [string]$Recomendacao = '',
        [string[]]$Correcao = @()
    )
    $Script:Achados.Add([pscustomobject]@{
            Categoria    = $Categoria
            Severidade   = $Severidade
            Ordem        = $Script:Sev[$Severidade].Ordem
            Titulo       = $Titulo
            Evidencia    = $Evidencia
            Recomendacao = $Recomendacao
            Correcoes    = @($Correcao | Where-Object { $_ })
        })
}

function Add-Acao {
    param([string]$Acao, [string]$Resultado)
    $Script:Acoes.Add([pscustomobject]@{ Hora = Get-Date; Acao = $Acao; Resultado = $Resultado })
    Write-Log "$Acao -> $Resultado"
}

function Add-Tabela {
    param([string]$Titulo, $Linhas)
    $Script:Dados.Tabelas[$Titulo] = @($Linhas)
}

#endregion

#region ===================== Proteção: backup, ponto de restauração e desfazer =====================

function New-PontoRestauracao {
    if ($Script:PontoRestauracaoFeito -or $Script:ModoAuto -or -not $Script:EhAdmin) { return }
    $Script:PontoRestauracaoFeito = $true
    if (-not (Get-Command Checkpoint-Computer -ErrorAction SilentlyContinue)) {
        Write-Host '  Ponto de restauração indisponível nesta edição do Windows. O backup das configurações continua ativo.' -ForegroundColor Yellow
        return
    }
    Write-Host '  Criando ponto de restauração do sistema...' -ForegroundColor Gray
    $k = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore'
    $freqAnt = (Get-ItemProperty $k -Name SystemRestorePointCreationFrequency -ErrorAction SilentlyContinue).SystemRestorePointCreationFrequency
    try {
        # O Windows só permite 1 ponto a cada 24 h; liberamos temporariamente
        Set-ItemProperty $k -Name SystemRestorePointCreationFrequency -Value 0 -Type DWord -ErrorAction SilentlyContinue
        $tentar = $true
        while ($tentar) {
            $tentar = $false
            try {
                Checkpoint-Computer -Description ("speedPC {0:dd/MM/yyyy HH:mm}" -f (Get-Date)) -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
                Write-Host '  Ponto de restauração criado.' -ForegroundColor Green
                Add-Acao 'Ponto de restauração' 'Criado'
            } catch {
                Write-Host "  Não foi possível criar o ponto de restauração: $($_.Exception.Message)" -ForegroundColor Yellow
                if (Confirm-SN 'A Proteção do Sistema pode estar desativada. Ativar na unidade do Windows e tentar de novo?' 'S') {
                    try { Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction Stop; $tentar = $true; Add-Acao 'Proteção do Sistema' "Ativada em $env:SystemDrive" }
                    catch { Write-Host "  Falha ao ativar: $($_.Exception.Message)" -ForegroundColor Yellow }
                }
                if (-not $tentar) { Add-Acao 'Ponto de restauração' 'Não criado' }
            }
        }
    } finally {
        if ($null -eq $freqAnt) { Remove-ItemProperty $k -Name SystemRestorePointCreationFrequency -ErrorAction SilentlyContinue }
        else { Set-ItemProperty $k -Name SystemRestorePointCreationFrequency -Value $freqAnt -Type DWord -ErrorAction SilentlyContinue }
    }
}

function Confirm-Protecao {
    if (-not $Script:ArquivoBackup) {
        $Script:ArquivoBackup = Join-Path $Script:PastaBackups ('{0:yyyy-MM-dd_HHmmss}.json' -f (Get-Date))
    }
    New-PontoRestauracao
}

function Save-Backup {
    if (-not $Script:ArquivoBackup) { return }
    if (-not (Test-Path -LiteralPath $Script:PastaBackups)) { New-Item -ItemType Directory -Path $Script:PastaBackups -Force | Out-Null }
    $obj = [pscustomobject]@{
        Ferramenta  = "speedPC $Script:Versao"
        Computador  = $env:COMPUTERNAME
        Data        = (Get-Date).ToString('s')
        Executor    = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        Alteracoes  = @($Script:Alteracoes.ToArray())
    }
    $json = ConvertTo-Json -InputObject $obj -Depth 6
    [IO.File]::WriteAllText($Script:ArquivoBackup, $json, (New-Object Text.UTF8Encoding($true)))
}

function Add-Alteracao {
    param($Registro)
    $Registro | Add-Member -NotePropertyName Hora -NotePropertyValue ((Get-Date).ToString('s')) -Force
    $Script:Alteracoes.Add($Registro)
    Save-Backup
}

function Set-RegistroSeguro {
    param(
        [string]$Caminho, [string]$Nome, $Valor,
        [ValidateSet('String', 'ExpandString', 'DWord', 'QWord', 'Binary', 'MultiString')][string]$Tipo = 'DWord',
        [string]$Descricao = ''
    )
    Confirm-Protecao
    $existia = $false; $anterior = $null; $tipoAnt = $null
    if (Test-Path -LiteralPath $Caminho) {
        $k = Get-Item -LiteralPath $Caminho
        if ($k.GetValueNames() -contains $Nome) {
            $existia = $true
            $tipoAnt = $k.GetValueKind($Nome).ToString()
            $anterior = $k.GetValue($Nome, $null, 'DoNotExpandEnvironmentNames')
            if ($anterior -is [byte[]]) { $anterior = [Convert]::ToBase64String($anterior) }
        }
    } else {
        New-Item -Path $Caminho -Force | Out-Null
    }
    Add-Alteracao ([pscustomobject]@{
            Tipo = 'Registro'; Caminho = $Caminho; Nome = $Nome; Existia = $existia
            ValorAnterior = $anterior; TipoAnterior = $tipoAnt; Descricao = $Descricao
        })
    New-ItemProperty -LiteralPath $Caminho -Name $Nome -Value $Valor -PropertyType $Tipo -Force | Out-Null
}

function Set-ServicoSeguro {
    param([string]$Nome, [ValidateSet('', 'Automatic', 'Manual', 'Disabled')][string]$Inicio = '', [switch]$Iniciar)
    Confirm-Protecao
    $w = Get-CimInstance Win32_Service -Filter "Name='$Nome'" -ErrorAction SilentlyContinue
    if (-not $w) { throw "Serviço $Nome não encontrado." }
    Add-Alteracao ([pscustomobject]@{ Tipo = 'Servico'; Nome = $Nome; InicioAnterior = $w.StartMode; EstadoAnterior = $w.State; Descricao = "Serviço $($w.DisplayName)" })
    if ($Inicio) { Set-Service -Name $Nome -StartupType $Inicio -ErrorAction Stop }
    if ($Iniciar) { Start-Service -Name $Nome -ErrorAction Stop }
}

function Enable-DispositivoSeguro {
    param([string]$InstanceId, [string]$Nome)
    Confirm-Protecao
    Add-Alteracao ([pscustomobject]@{ Tipo = 'Dispositivo'; InstanceId = $InstanceId; Nome = $Nome; Descricao = "Dispositivo ativado: $Nome" })
    Enable-PnpDevice -InstanceId $InstanceId -Confirm:$false -ErrorAction Stop
}

function Set-PlanoEnergiaSeguro {
    param([string]$Guid, [string]$Nome)
    Confirm-Protecao
    $atual = Get-PlanoEnergiaAtivo
    Add-Alteracao ([pscustomobject]@{ Tipo = 'PlanoEnergia'; GuidAnterior = $atual.Guid; NomeAnterior = $atual.Nome; Descricao = "Plano de energia: $($atual.Nome) -> $Nome" })
    powercfg /setactive $Guid | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "powercfg retornou $LASTEXITCODE" }
}

function Get-SessoesBackup {
    $sessoes = @()
    $arqs = @(Get-ChildItem -LiteralPath $Script:PastaBackups -Filter '*.json' -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notlike '*.desfeito.json' } | Sort-Object Name -Descending)
    foreach ($a in $arqs) {
        try {
            $j = Get-Content -LiteralPath $a.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            $sessoes += [pscustomobject]@{
                Arquivo = $a.FullName; Data = [datetime]$j.Data; Executor = $j.Executor
                Qtd = @($j.Alteracoes).Count; Itens = @($j.Alteracoes | ForEach-Object { [string]$_.Descricao })
            }
        } catch { }
    }
    return $sessoes
}

function Invoke-Desfazer {
    Write-Titulo 'Desfazer alterações'
    $sessoes = @(Get-SessoesBackup)
    if (-not $sessoes) { Write-Host '  Nenhum backup de alterações encontrado.' -ForegroundColor Yellow; return }
    for ($i = 0; $i -lt $sessoes.Count; $i++) {
        $s = $sessoes[$i]
        Write-Host ("  [{0}] {1}  -  {2} alteração(ões)  -  por {3}" -f ($i + 1), $s.Data.ToString('dd/MM/yyyy HH:mm'), $s.Qtd, $s.Executor)
    }
    Write-Host ''
    Write-Host '  Dica: para voltar o sistema inteiro, use também o ponto de restauração (rstrui.exe).' -ForegroundColor DarkGray
    $r = Read-Host '  Número da sessão a desfazer (Enter = cancelar)'
    if ($r -notmatch '^\d+$' -or [int]$r -lt 1 -or [int]$r -gt $sessoes.Count) { return }
    $s = $sessoes[[int]$r - 1]
    Write-Host ''
    foreach ($it in $s.Itens) { Write-Host "   - $it" -ForegroundColor Gray }
    if (-not (Confirm-SN "Reverter estas $($s.Qtd) alterações?" 'N')) { return }
    Undo-SessaoBackup -Arquivo $s.Arquivo
}

function Undo-SessaoBackup {
    param([string]$Arquivo)
    Write-Titulo 'Desfazendo alterações'
    $j = Get-Content -LiteralPath $Arquivo -Raw -Encoding UTF8 | ConvertFrom-Json
    $itens = @($j.Alteracoes)
    [array]::Reverse($itens)
    $falhas = 0
    foreach ($a in $itens) {
        try {
            switch ($a.Tipo) {
                'Registro' {
                    if ($a.Existia) {
                        $v = $a.ValorAnterior
                        switch ($a.TipoAnterior) {
                            'Binary' { $v = [Convert]::FromBase64String([string]$v) }
                            'DWord' { $v = [int]$v }
                            'QWord' { $v = [int64]$v }
                            'MultiString' { $v = [string[]]@($v) }
                        }
                        New-ItemProperty -LiteralPath $a.Caminho -Name $a.Nome -Value $v -PropertyType $a.TipoAnterior -Force -ErrorAction Stop | Out-Null
                    } else {
                        Remove-ItemProperty -LiteralPath $a.Caminho -Name $a.Nome -Force -ErrorAction SilentlyContinue
                    }
                }
                'Servico' {
                    $mapa = @{ Auto = 'Automatic'; Manual = 'Manual'; Disabled = 'Disabled' }
                    if ($mapa.ContainsKey([string]$a.InicioAnterior)) { Set-Service -Name $a.Nome -StartupType $mapa[[string]$a.InicioAnterior] -ErrorAction Stop }
                    if ($a.EstadoAnterior -eq 'Stopped') { Stop-Service -Name $a.Nome -Force -ErrorAction SilentlyContinue }
                }
                'Dispositivo' { Disable-PnpDevice -InstanceId $a.InstanceId -Confirm:$false -ErrorAction Stop }
                'Adaptador' { Disable-NetAdapter -Name $a.Nome -Confirm:$false -ErrorAction Stop }
                'PlanoEnergia' { powercfg /setactive $a.GuidAnterior | Out-Null }
            }
            Write-Host "  [desfeito] $($a.Descricao)" -ForegroundColor Green
        } catch {
            $falhas++
            Write-Host "  [falhou]   $($a.Descricao): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    Rename-Item -LiteralPath $Arquivo -NewName ([IO.Path]::GetFileNameWithoutExtension($Arquivo) + '.desfeito.json') -ErrorAction SilentlyContinue
    Add-Acao 'Desfazer alterações' "$($itens.Count - $falhas) revertida(s), $falhas falha(s)"
    Write-Host '  Algumas alterações só valem após reiniciar ou fazer logoff.' -ForegroundColor DarkGray
}

#endregion

#region ===================== Resumo inicial =====================

function Get-CausaKernelPower {
    param($Evento)
    $d = Get-DadosEvento $Evento
    $bc = [int64]0; [void][int64]::TryParse([string]$d['BugcheckCode'], [ref]$bc)
    $pb = [int64]0; [void][int64]::TryParse([string]$d['PowerButtonTimestamp'], [ref]$pb)
    if ($bc -ne 0) {
        $cod = ConvertTo-CodigoBugCheck ('{0:X}' -f $bc)
        $nome = ''
        if ($Script:BugChecks.ContainsKey($cod)) { $nome = ' ' + $Script:BugChecks[$cod][0] }
        return [pscustomobject]@{ Texto = "tela azul (0x$cod$nome)"; Codigo = $cod }
    }
    if ($pb -ne 0) { return [pscustomobject]@{ Texto = 'botão de energia pressionado (desligamento forçado)'; Codigo = $null } }
    return [pscustomobject]@{ Texto = 'falta de energia, bateria esgotada ou congelamento total'; Codigo = $null }
}

function ConvertTo-CodigoBugCheck {
    param([string]$Hex)
    try { $v = [Convert]::ToInt64($Hex, 16) } catch { return $Hex.ToUpper() }
    if ($v -ge 0x10000000 -and $v -lt 0x20000000) { $v = $v - 0x10000000 }   # variantes 0x1000007E etc.
    return ('{0:X}' -f $v)
}

function Get-BootsRecentes {
    if (-not $Script:Dados.Boots) {
        $Script:Dados.Boots = @(Get-Eventos -Log 'System' -Ids 12 -Provider 'Microsoft-Windows-Kernel-General' -Dias 95 -Max 1000 | Sort-Object TimeCreated)
    }
    return $Script:Dados.Boots
}

function Get-ContextoTravamento {
    # O evento 41 é gravado no boot SEGUINTE ao travamento. Aqui olhamos o fim da sessão anterior
    # para descobrir o que o Windows estava fazendo quando travou: desligando, suspendendo ou em uso.
    param([datetime]$Momento)
    $boots = @(Get-BootsRecentes)
    $atual = $boots | Where-Object { $_.TimeCreated -le $Momento.AddMinutes(1) } | Select-Object -Last 1
    if (-not $atual) { return $null }
    $ant = $boots | Where-Object { $_.TimeCreated -lt $atual.TimeCreated.AddSeconds(-5) } | Select-Object -Last 1
    if (-not $ant) { return $null }
    $fim = $atual.TimeCreated.AddSeconds(-5)
    $ult = $null
    try { $ult = Get-WinEvent -FilterHashtable @{ LogName = 'System'; StartTime = $ant.TimeCreated; EndTime = $fim } -MaxEvents 1 -ErrorAction Stop } catch { }
    if (-not $ult) { return $null }
    $marca = $null
    try {
        $marca = Get-WinEvent -FilterHashtable @{ LogName = 'System'; StartTime = $ant.TimeCreated; EndTime = $fim
            ProviderName = 'User32', 'Microsoft-Windows-Kernel-Power', 'Microsoft-Windows-Kernel-General'; Id = 1074, 109, 13, 42 } -MaxEvents 1 -ErrorAction Stop
    } catch { }
    $momentoTxt = 'Em uso'; $pedido = $null
    if ($marca -and ($ult.TimeCreated - $marca.TimeCreated).TotalMinutes -le 15) {
        $pedido = $marca.TimeCreated
        if ($marca.Id -eq 42) { $momentoTxt = 'Suspender' } else { $momentoTxt = 'Desligar' }
    }
    return [pscustomobject]@{ InicioSessao = $ant.TimeCreated; UltimoRegistro = $ult.TimeCreated; Momento = $momentoTxt; Pedido = $pedido }
}

function Get-UltimoDesligamento {
    param([datetime]$Boot)
    $ev = @()
    $ev += Get-Eventos -Log 'System' -Ids 1074 -Dias 90 -Max 10
    $ev += Get-Eventos -Log 'System' -Ids 41 -Provider 'Microsoft-Windows-Kernel-Power' -Dias 90 -Max 10
    $ev += Get-Eventos -Log 'System' -Ids 6008 -Provider 'EventLog' -Dias 90 -Max 10
    $ev = @($ev | Where-Object { $_ } | Sort-Object TimeCreated -Descending)
    if (-not $ev) {
        return [pscustomobject]@{ Tipo = 'Desconhecido'; Texto = 'Sem registros de desligamento nos últimos 90 dias'; Data = $null; Detalhe = ''; Observacao = '' }
    }
    $u = $ev[0]
    if ($u.Id -eq 1074) {
        $p = $u.Properties
        $proc = ''
        if ($p.Count -gt 0) { $proc = ([string]$p[0].Value) -replace '\s*\(.*$', ''; try { $proc = Split-Path $proc -Leaf } catch { } }
        $tipo = 'desligamento'; if ($p.Count -gt 4 -and $p[4].Value) { $tipo = [string]$p[4].Value }
        $usr = ''; if ($p.Count -gt 6) { $usr = [string]$p[6].Value }
        $motivo = ''; if ($p.Count -gt 2) { $motivo = [string]$p[2].Value }
        $obs = ''
        if ($Boot -and $u.TimeCreated -gt $Boot.AddMinutes(2)) {
            $obs = 'Foi um "Desligar" com Inicialização Rápida: o Windows não reiniciou de fato (o tempo ligado continua contando).'
        }
        return [pscustomobject]@{
            Tipo = 'Normal'; Data = $u.TimeCreated; Detalhe = $motivo; Observacao = $obs
            Texto = ('Normal - {0} por {1} via {2} em {3:dd/MM/yyyy HH:mm}' -f $tipo, $usr, $proc, $u.TimeCreated)
        }
    }
    $k41 = $ev | Where-Object { $_.Id -eq 41 } | Select-Object -First 1
    $causa = 'registrado pelo Windows ao ligar novamente'
    if ($k41 -and [math]::Abs(($k41.TimeCreated - $u.TimeCreated).TotalMinutes) -lt 15) { $causa = (Get-CausaKernelPower $k41).Texto }
    $obs = ''
    $ctx = $null; if ($k41) { $ctx = Get-ContextoTravamento $k41.TimeCreated }
    if ($ctx -and $ctx.Momento -eq 'Desligar') {
        $causa += ' ao desligar'
        $obs = ('O desligamento pedido em {0:dd/MM HH:mm} não terminou: o Windows travou enquanto desligava.' -f $ctx.Pedido)
    } elseif ($ctx -and $ctx.Momento -eq 'Suspender') {
        $causa += ' ao suspender'
        $obs = ('A suspensão pedida em {0:dd/MM HH:mm} não terminou: o Windows travou ao entrar em suspensão.' -f $ctx.Pedido)
    } elseif ($ctx) {
        $obs = ('Travou durante o uso (último registro em {0:dd/MM HH:mm}).' -f $ctx.UltimoRegistro)
    }
    return [pscustomobject]@{
        Tipo = 'Inesperado'; Data = $u.TimeCreated; Detalhe = ''; Observacao = $obs
        Texto = ('INESPERADO - {0} (detectado em {1:dd/MM/yyyy HH:mm})' -f $causa, $u.TimeCreated)
    }
}

function Get-TopMemoria {
    param([int]$Quantidade = 8)
    Get-Process -ErrorAction SilentlyContinue | Group-Object ProcessName | ForEach-Object {
        $soma = ($_.Group | Measure-Object WorkingSet64 -Sum).Sum
        $desc = ($_.Group | Where-Object { $_.Description } | Select-Object -First 1).Description
        if (-not $desc) { $desc = $_.Name }
        [pscustomobject]@{ Processo = $_.Name; Nome = $desc; Bytes = [int64]$soma; Instancias = $_.Count }
    } | Sort-Object Bytes -Descending | Select-Object -First $Quantidade
}

function Get-Resumo {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
    $cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
    $boot = $os.LastBootUpTime
    $total = [double]$os.TotalVisibleMemorySize * 1KB
    $livre = [double]$os.FreePhysicalMemory * 1KB
    $pct = 0; if ($total -gt 0) { $pct = [math]::Round(($total - $livre) / $total * 100) }
    $versao = $cv.DisplayVersion; if (-not $versao) { $versao = $cv.ReleaseId }
    $ips = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' } | ForEach-Object IPAddress)
    $sistema = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'" -ErrorAction SilentlyContinue

    $Script:Dados.Resumo = [pscustomobject]@{
        Computador     = $env:COMPUTERNAME
        Boot           = $boot
        TempoLigado    = (Get-Date) - $boot
        Desligamento   = Get-UltimoDesligamento -Boot $boot
        RamTotal       = $total
        RamUsada       = $total - $livre
        RamPct         = $pct
        TopMemoria     = @(Get-TopMemoria 8)
        Fabricante     = $cs.Manufacturer
        Modelo         = $cs.Model
        Serie          = $bios.SerialNumber
        Bios           = ('{0} ({1})' -f $bios.SMBIOSBIOSVersion, $(if ($bios.ReleaseDate) { $bios.ReleaseDate.ToString('dd/MM/yyyy') } else { '?' }))
        Windows        = ('{0} {1} (build {2}.{3})' -f $os.Caption, $versao, $cv.CurrentBuild, $cv.UBR)
        Processador    = ($cpu.Name -replace '\s+', ' ').Trim()
        Nucleos        = $cs.NumberOfLogicalProcessors
        Dominio        = $(if ($cs.PartOfDomain) { $cs.Domain } else { "Grupo de trabalho ($($cs.Workgroup))" })
        NoDominio      = [bool]$cs.PartOfDomain
        UsuarioLogado  = $cs.UserName
        IPs            = ($ips -join ', ')
        DiscoSistema   = $sistema
        Notebook       = [bool](Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue)
        InicioRapido   = ((Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' -ErrorAction SilentlyContinue).HiberbootEnabled -eq 1)
    }
    return $Script:Dados.Resumo
}

function Show-Resumo {
    $r = $Script:Dados.Resumo
    if (-not $r) { $r = Get-Resumo }
    Clear-Host
    Write-Host ''
    Write-Host ("  speedPC {0}  |  {1}  |  {2}" -f $Script:Versao, $r.Computador, (Get-Date).ToString('dd/MM/yyyy HH:mm')) -ForegroundColor Cyan
    Write-Host ("  {0}  |  {1} {2}" -f $r.Windows, $r.Fabricante, $r.Modelo) -ForegroundColor DarkGray
    if ($r.UsuarioLogado) { Write-Host "  Usuário logado: $($r.UsuarioLogado)" -ForegroundColor DarkGray }
    Write-Host ('  ' + ('-' * 76)) -ForegroundColor DarkGray

    $corT = 'Green'; if ($r.TempoLigado.TotalDays -ge 7) { $corT = 'Red' } elseif ($r.TempoLigado.TotalDays -ge 3) { $corT = 'Yellow' }
    Write-Host '  Ligado há            ' -NoNewline
    Write-Host (Format-Duracao $r.TempoLigado) -ForegroundColor $corT -NoNewline
    Write-Host ("  (desde {0:dd/MM/yyyy HH:mm})" -f $r.Boot) -ForegroundColor DarkGray

    $corD = 'Green'; if ($r.Desligamento.Tipo -eq 'Inesperado') { $corD = 'Red' } elseif ($r.Desligamento.Tipo -eq 'Desconhecido') { $corD = 'Gray' }
    Write-Host '  Último desligamento  ' -NoNewline
    Write-Host $r.Desligamento.Texto -ForegroundColor $corD
    if ($r.Desligamento.Observacao) { Write-Host "                       $($r.Desligamento.Observacao)" -ForegroundColor Yellow }

    $corR = 'Green'; if ($r.RamPct -ge 90) { $corR = 'Red' } elseif ($r.RamPct -ge 80) { $corR = 'Yellow' }
    Write-Host '  Memória RAM          ' -NoNewline
    Write-Host ("{0} {1} de {2} ({3}%)" -f (Get-Barra $r.RamPct), (Format-Bytes $r.RamUsada), (Format-Bytes $r.RamTotal), $r.RamPct) -ForegroundColor $corR
    $tops = ($r.TopMemoria | Select-Object -First 4 | ForEach-Object {
            $n = $_.Nome; if ($n.Length -gt 22) { $n = $n.Substring(0, 22).Trim() + '…' }
            '{0} {1}' -f $n, (Format-Bytes $_.Bytes) }) -join '  |  '
    Write-Host "                       Maiores: $tops" -ForegroundColor DarkGray

    if ($r.DiscoSistema -and $r.DiscoSistema.Size) {
        $pl = [math]::Round($r.DiscoSistema.FreeSpace / $r.DiscoSistema.Size * 100)
        $corC = 'Green'; if ($pl -lt 10) { $corC = 'Red' } elseif ($pl -lt 15) { $corC = 'Yellow' }
        Write-Host ("  Disco {0}             " -f $env:SystemDrive) -NoNewline
        Write-Host ("{0} livre de {1} ({2}% livre)" -f (Format-Bytes $r.DiscoSistema.FreeSpace), (Format-Bytes $r.DiscoSistema.Size), $pl) -ForegroundColor $corC
    }
    if (-not $Script:EhAdmin) {
        Write-Host ''
        Write-Host '  Executando SEM administrador: várias verificações ficam limitadas e as correções não estão disponíveis.' -ForegroundColor Yellow
    }
}

#endregion

#region ===================== Correções (catálogo) =====================

function New-Correcao {
    param(
        [string]$Chave, [string]$Nome, [string]$Descricao,
        [ValidateSet('Reversivel', 'SemEfeitoPermanente', 'NaoReversivel')][string]$Reversao = 'Reversivel',
        [scriptblock]$Acao, [hashtable]$Parametros = @{}, [switch]$Interativa
    )
    if (-not $Script:Correcoes.Contains($Chave)) {
        $Script:Correcoes[$Chave] = [pscustomobject]@{
            Chave = $Chave; Nome = $Nome; Descricao = $Descricao; Reversao = $Reversao
            Acao = $Acao; Parametros = $Parametros; Interativa = [bool]$Interativa; Aplicada = $false
        }
    }
    return $Chave
}

function Get-CorrecaoPadrao {
    param([string]$Chave)
    switch ($Chave) {
        'Limpeza' { return New-Correcao $Chave 'Limpeza segura' 'Remove temporários e caches de todos os usuários e do Windows (sem tocar em senhas, histórico ou arquivos pessoais).' 'SemEfeitoPermanente' { Invoke-LimpezaSegura -Silencioso } }
        'InicioRapido' { return New-Correcao $Chave 'Desativar Inicialização Rápida' 'Faz o "Desligar" reiniciar o Windows de verdade, zerando o tempo ligado e liberando memória.' 'Reversivel' { Disable-InicioRapido } }
        'PlanoEquilibrado' { return New-Correcao $Chave 'Plano de energia Equilibrado' 'Sai do modo Economia de energia, que limita o processador.' 'Reversivel' { Set-PlanoEnergiaSeguro -Guid $Script:GuidPlanos.Equilibrado -Nome 'Equilibrado' } }
        'Inicializacao' { return New-Correcao $Chave 'Revisar programas de inicialização' 'Lista os programas que iniciam com o Windows para você escolher quais desativar (igual ao Gerenciador de Tarefas).' 'Reversivel' { Invoke-RevisaoInicializacao } -Interativa }
        'ReparoWindows' { return New-Correcao $Chave 'Reparar o Windows (DISM + SFC)' 'Verifica e corrige arquivos do sistema corrompidos. Leva de 10 a 40 minutos.' 'SemEfeitoPermanente' { Invoke-ReparoSistema } }
        'VerificarDisco' { return New-Correcao $Chave 'Verificar o disco (CHKDSK online)' 'Analisa o sistema de arquivos sem precisar reiniciar.' 'SemEfeitoPermanente' { Invoke-VerificacaoDisco } }
        'AtualizarDrivers' { return New-Correcao $Chave 'Atualizar drivers (Windows Update)' 'Procura e instala drivers publicados pela Microsoft/fabricante no Windows Update.' 'NaoReversivel' { Install-AtualizacoesWU -Tipo Driver } -Interativa }
        'AtualizarWindows' { return New-Correcao $Chave 'Instalar atualizações do Windows' 'Procura e instala atualizações de segurança e qualidade.' 'NaoReversivel' { Install-AtualizacoesWU -Tipo Software } -Interativa }
        'TesteMemoria' { return New-Correcao $Chave 'Agendar teste de memória RAM' 'Abre o Diagnóstico de Memória do Windows (o teste roda no próximo reinício).' 'SemEfeitoPermanente' { Start-Process "$env:SystemRoot\System32\mdsched.exe" } -Interativa }
        'RenovarRede' { return New-Correcao $Chave 'Renovar IP e limpar cache DNS' 'ipconfig /renew, /flushdns e /registerdns. Em acesso remoto a conexão pode oscilar por alguns segundos.' 'SemEfeitoPermanente' { Invoke-RenovarRede } }
        'ResetRede' { return New-Correcao $Chave 'Redefinir a pilha de rede (Winsock/TCP-IP)' 'Corrige rede corrompida por VPN/antivírus antigos. Exige reiniciar o computador.' 'NaoReversivel' { Invoke-ResetRede } }
        'LimparFila' { return New-Correcao $Chave 'Destravar fila de impressão' 'Para o spooler, remove os trabalhos presos e inicia de novo. Os documentos na fila precisarão ser reimpressos.' 'NaoReversivel' { Clear-FilaImpressao } }
        'IniciarSpooler' { return New-Correcao $Chave 'Iniciar o serviço de impressão (Spooler)' 'Coloca o Spooler em automático e o inicia.' 'Reversivel' { Set-ServicoSeguro -Nome Spooler -Inicio Automatic -Iniciar } }
        'ReparoWU' { return New-Correcao $Chave 'Reparar o Windows Update' 'Para os serviços, renomeia SoftwareDistribution e catroot2 (ficam guardadas como .bak) e reinicia os serviços.' 'SemEfeitoPermanente' { Invoke-ReparoWindowsUpdate } }
    }
    return $null
}

function Invoke-Correcao {
    param($Correcao)
    Write-Host ''
    Write-Host ">> $($Correcao.Nome)" -ForegroundColor Cyan
    try {
        if ($Correcao.Reversao -ne 'SemEfeitoPermanente') { Confirm-Protecao }
        & $Correcao.Acao $Correcao.Parametros
        $Correcao.Aplicada = $true
        Add-Acao $Correcao.Nome 'Concluído'
        Write-Host '   Concluído.' -ForegroundColor Green
    } catch {
        Add-Acao $Correcao.Nome "Falhou: $($_.Exception.Message)"
        Write-Host "   Falhou: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Invoke-MenuCorrecoes {
    $lista = @($Script:Correcoes.Values | Where-Object { -not $_.Aplicada })
    if (-not $lista) { Write-Host '  Nenhuma correção automática disponível para os problemas encontrados.' -ForegroundColor Green; return }
    if (-not $Script:EhAdmin) { Write-Host '  As correções exigem executar como administrador.' -ForegroundColor Yellow; return }
    Write-Titulo 'Correções automáticas disponíveis'
    $rev = @{ Reversivel = 'pode ser desfeita'; SemEfeitoPermanente = 'sem alteração de configuração'; NaoReversivel = 'não reversível' }
    $itens = @($lista | ForEach-Object {
            $extra = ''; if ($_.Interativa) { $extra = ' · pede confirmação durante a execução' }
            [pscustomobject]@{ Nome = $_.Nome; Detalhe = "$($_.Descricao)  (Reversão: $($rev[$_.Reversao])$extra)" } })
    $sel = @(Select-Opcoes -Titulo 'Correções automáticas disponíveis' -Itens $itens -Aviso 'Antes da primeira alteração é criado um ponto de restauração e um backup de cada configuração.')
    foreach ($i in $sel) { Invoke-Correcao $lista[$i] }
}

#endregion

#region ===================== Diagnóstico: Lentidão =====================

function Get-ItensInicializacao {
    $itens = New-Object System.Collections.Generic.List[object]
    $specs = @(
        @{ Raiz = 'Registry::HKEY_LOCAL_MACHINE'; Run = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Run'; Aprov = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run'; Escopo = 'Todos os usuários' },
        @{ Raiz = 'Registry::HKEY_LOCAL_MACHINE'; Run = 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'; Aprov = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32'; Escopo = 'Todos os usuários' }
    )
    $hives = @(Get-HivesUsuarios)
    foreach ($h in $hives) {
        $specs += @{ Raiz = "Registry::HKEY_USERS\$($h.Sid)"; Run = 'Software\Microsoft\Windows\CurrentVersion\Run'; Aprov = 'Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run'; Escopo = $h.Usuario }
    }
    foreach ($s in $specs) {
        $k = "$($s.Raiz)\$($s.Run)"
        if (-not (Test-Path -LiteralPath $k)) { continue }
        $chave = Get-Item -LiteralPath $k
        foreach ($n in $chave.GetValueNames()) {
            if (-not $n) { continue }
            $aprov = "$($s.Raiz)\$($s.Aprov)"
            $itens.Add([pscustomobject]@{
                    Nome = $n; Comando = [string]$chave.GetValue($n); Escopo = $s.Escopo; Origem = 'Registro'
                    Ativo = (Test-ItemAprovado $aprov $n); ChaveAprovacao = $aprov; NomeAprovacao = $n
                })
        }
    }
    $pastas = @(@{ Pasta = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"; Aprov = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder'; Escopo = 'Todos os usuários' })
    foreach ($h in $hives) {
        if ($h.Perfil) {
            $pastas += @{ Pasta = (Join-Path $h.Perfil 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup'); Aprov = "Registry::HKEY_USERS\$($h.Sid)\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder"; Escopo = $h.Usuario }
        }
    }
    foreach ($p in $pastas) {
        if (-not (Test-Path -LiteralPath $p.Pasta)) { continue }
        foreach ($f in (Get-ChildItem -LiteralPath $p.Pasta -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'desktop.ini' })) {
            $itens.Add([pscustomobject]@{
                    Nome = $f.BaseName; Comando = $f.FullName; Escopo = $p.Escopo; Origem = 'Pasta Inicializar'
                    Ativo = (Test-ItemAprovado $p.Aprov $f.Name); ChaveAprovacao = $p.Aprov; NomeAprovacao = $f.Name
                })
        }
    }
    foreach ($i in $itens) {
        $exe = ''
        $m = [regex]::Matches($i.Comando, '([^\\/"]+\.exe)', 'IgnoreCase')
        if ($m.Count -gt 0) { $exe = $m[0].Groups[1].Value.ToLower() }
        $atraso = 0
        if ($exe -and $Script:Dados.AtrasoBoot.ContainsKey($exe)) { $atraso = $Script:Dados.AtrasoBoot[$exe] }
        $i | Add-Member -NotePropertyName Executavel -NotePropertyValue $exe -Force
        $i | Add-Member -NotePropertyName AtrasoMs -NotePropertyValue $atraso -Force
    }
    return $itens
}

function Test-ItemAprovado {
    param([string]$Chave, [string]$Nome)
    try {
        $v = (Get-Item -LiteralPath $Chave -ErrorAction Stop).GetValue($Nome)
        if ($v -is [byte[]] -and $v.Length -gt 0) { return (($v[0] -band 1) -eq 0) }
    } catch { }
    return $true
}

function Invoke-RevisaoInicializacao {
    $itens = @(Get-ItensInicializacao | Where-Object { $_.Ativo } | Sort-Object AtrasoMs -Descending)
    if (-not $itens) { Write-Host '  Nenhum programa ativo na inicialização.' -ForegroundColor Green; return }
    $opcoes = @($itens | ForEach-Object {
            $atraso = ''; if ($_.AtrasoMs -gt 0) { $atraso = ' · atrasou o boot em {0:N1} s' -f ($_.AtrasoMs / 1000) }
            $cmd = $_.Comando; if ($cmd.Length -gt 110) { $cmd = $cmd.Substring(0, 110) + '...' }
            [pscustomobject]@{ Nome = "$($_.Nome)  ($($_.Escopo)$atraso)"; Detalhe = $cmd } })
    $sel = @(Select-Opcoes -Titulo 'Quais programas DESATIVAR da inicialização?' -Itens $opcoes -Aviso 'NÃO desative: antivírus, VPN, drivers (áudio, touchpad, vídeo), agentes de TI/inventário e backup.')
    foreach ($i in $sel) {
        $it = $itens[$i]
        try {
            Set-RegistroSeguro -Caminho $it.ChaveAprovacao -Nome $it.NomeAprovacao -Valor ([byte[]](3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)) -Tipo Binary -Descricao "Inicialização desativada: $($it.Nome) ($($it.Escopo))"
            Write-Host "   Desativado: $($it.Nome)" -ForegroundColor Green
            Add-Acao "Desativar da inicialização: $($it.Nome)" 'Concluído'
        } catch { Write-Host "   Falha em $($it.Nome): $($_.Exception.Message)" -ForegroundColor Red }
    }
}

function Disable-InicioRapido {
    Set-RegistroSeguro -Caminho 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power' -Nome 'HiberbootEnabled' -Valor 0 -Tipo DWord -Descricao 'Inicialização Rápida desativada'
}

function Test-Lentidao {
    Write-Etapa 'Lentidão: tempo ligado, memória, processador, temperatura, disco e inicialização'
    $cat = 'Lentidão'
    $r = $Script:Dados.Resumo

    # --- Tempo ligado
    $dias = $r.TempoLigado.TotalDays
    if ($dias -ge 7) {
        $sev = 'Atencao'; if ($dias -ge 21) { $sev = 'Critico' }
        $ev = "Ligado há $(Format-Duracao $r.TempoLigado), desde $($r.Boot.ToString('dd/MM/yyyy HH:mm'))."
        $corr = @()
        if ($r.InicioRapido) {
            $ev += ' A Inicialização Rápida está ativa: "Desligar" não reinicia o Windows; só "Reiniciar" zera esse tempo.'
            $corr += Get-CorrecaoPadrao 'InicioRapido'
        }
        Add-Achado $cat $sev 'PC sem reiniciar há vários dias' $ev 'Reinicie o computador (use "Reiniciar", não "Desligar"). Vazamentos de memória e atualizações pendentes se acumulam com o tempo.' $corr
    } else {
        Add-Achado $cat 'OK' "Reiniciado recentemente ($(Format-Duracao $r.TempoLigado))"
    }

    # --- Memória RAM
    $top = ($r.TopMemoria | Select-Object -First 5 | ForEach-Object {
            $q = ''; if ($_.Instancias -gt 1) { $q = " ($($_.Instancias) processos)" }
            '{0} {1}{2}' -f $_.Nome, (Format-Bytes $_.Bytes), $q }) -join '; '
    Add-Tabela 'Maiores consumidores de memória RAM' ($r.TopMemoria | ForEach-Object {
            [pscustomobject]@{ Programa = $_.Nome; Processo = "$($_.Processo).exe"; 'Instâncias' = $_.Instancias; 'Memória' = (Format-Bytes $_.Bytes) } })
    $ramGB = [math]::Round($r.RamTotal / 1GB)
    if ($r.RamPct -ge 80) {
        $sev = 'Atencao'; if ($r.RamPct -ge 90) { $sev = 'Critico' }
        $rec = 'Feche programas e abas que não estão em uso; reinicie se o PC está ligado há dias.'
        if ($ramGB -lt 8) { $rec += " Com apenas $ramGB GB, recomenda-se ampliar para 8 GB ou 16 GB." }
        elseif ($ramGB -lt 16) { $rec += ' Se o uso alto for frequente, considere 16 GB.' }
        Add-Achado $cat $sev "Memória RAM quase esgotada ($($r.RamPct)%)" ("Em uso {0} de {1}. Maiores consumidores: {2}." -f (Format-Bytes $r.RamUsada), (Format-Bytes $r.RamTotal), $top) $rec
    } else {
        Add-Achado $cat 'OK' "Memória RAM em $($r.RamPct)% de uso"
        if ($ramGB -lt 8) { Add-Achado $cat 'Info' "Apenas $ramGB GB de RAM" 'Abaixo do recomendado para Windows 10/11 com navegador, Office e Teams.' 'Planeje ampliar para 8 GB ou mais.' }
    }

    # --- Processador
    $cargas = @()
    for ($i = 0; $i -lt 3; $i++) {
        $c = (Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Measure-Object LoadPercentage -Average).Average
        if ($null -ne $c) { $cargas += $c }
        Start-Sleep -Milliseconds 600
    }
    $carga = 0; if ($cargas) { $carga = [math]::Round(($cargas | Measure-Object -Average).Average) }
    $nucleos = [math]::Max(1, [int]$r.Nucleos)
    $topCpu = @(Get-CimInstance Win32_PerfFormattedData_PerfProc_Process -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notin '_Total', 'Idle' -and $_.PercentProcessorTime -gt 0 } |
            Sort-Object PercentProcessorTime -Descending | Select-Object -First 5)
    $txtCpu = ($topCpu | ForEach-Object { '{0} {1:N0}%' -f ($_.Name -replace '#\d+$', ''), ($_.PercentProcessorTime / $nucleos) }) -join '; '
    if ($carga -ge 80) {
        Add-Achado $cat 'Atencao' "Processador sobrecarregado agora ($carga%)" "Maiores consumidores neste momento: $txtCpu." 'Verifique se é uma tarefa pontual (antivírus, atualização, indexação) ou um programa travado consumindo CPU.'
    } else {
        Add-Achado $cat 'OK' "Uso do processador em $carga%"
    }

    # --- Temperatura e limitação térmica
    $temps = @()
    try {
        $tz = Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop
        $temps = @($tz | ForEach-Object { [math]::Round($_.CurrentTemperature / 10 - 273.15, 1) } | Where-Object { $_ -gt 5 -and $_ -lt 125 })
    } catch { }
    $limitacoes = @(Get-Eventos -Log 'System' -Ids 37 -Provider 'Microsoft-Windows-Kernel-Processor-Power' -Dias 7 -Max 200)
    $tMax = $null; if ($temps) { $tMax = ($temps | Measure-Object -Maximum).Maximum }
    $recTemp = 'Limpe poeira das saídas de ar e ventoinhas, verifique se a ventoinha gira, troque a pasta térmica em máquinas com mais de 3 anos e use o carregador original (em notebooks, carregador mais fraco também limita o processador).'
    if ($tMax -ge 90) {
        Add-Achado $cat 'Critico' "Processador superaquecendo ($tMax °C)" "Sensor térmico ACPI em $tMax °C. Eventos de limitação de velocidade nos últimos 7 dias: $($limitacoes.Count)." $recTemp
    } elseif ($tMax -ge 80 -or $limitacoes.Count -gt 0) {
        $ev = @()
        if ($tMax) { $ev += "Sensor térmico ACPI em $tMax °C." }
        if ($limitacoes.Count -gt 0) { $ev += "O Windows registrou $($limitacoes.Count) vez(es) em 7 dias que o processador teve a velocidade limitada pelo firmware (último: $($limitacoes[0].TimeCreated.ToString('dd/MM HH:mm')))." }
        Add-Achado $cat 'Atencao' 'Processador esquentando ou com velocidade limitada' ($ev -join ' ') $recTemp
    } elseif ($tMax) {
        Add-Achado $cat 'OK' "Temperatura em $tMax °C (sensor ACPI)"
    } else {
        Add-Achado $cat 'OK' 'Sem registros de limitação térmica do processador (7 dias)'
    }

    # --- Discos: espaço
    $linhasDisco = @()
    foreach ($v in (Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction SilentlyContinue)) {
        if (-not $v.Size) { continue }
        $pct = [math]::Round($v.FreeSpace / $v.Size * 100, 1)
        $sist = ($v.DeviceID -eq $env:SystemDrive)
        $linhasDisco += [pscustomobject]@{ Unidade = $v.DeviceID; 'Rótulo' = $v.VolumeName; Tamanho = (Format-Bytes $v.Size); Livre = (Format-Bytes $v.FreeSpace); '% livre' = "$pct%" }
        $txt = "{0} livre de {1} ({2}%)" -f (Format-Bytes $v.FreeSpace), (Format-Bytes $v.Size), $pct
        if ($pct -lt 10 -or $v.FreeSpace -lt 5GB) {
            Add-Achado $cat 'Critico' "Disco $($v.DeviceID) quase cheio" $txt 'Rode a limpeza segura; depois verifique Downloads, Área de Trabalho e arquivos grandes com o usuário. Abaixo de 10% livre o Windows fica lento e atualizações falham.' @(Get-CorrecaoPadrao 'Limpeza')
        } elseif ($pct -lt 15 -or ($sist -and $v.FreeSpace -lt 20GB)) {
            Add-Achado $cat 'Atencao' "Pouco espaço no disco $($v.DeviceID)" $txt 'Rode a limpeza segura.' @(Get-CorrecaoPadrao 'Limpeza')
        } else {
            Add-Achado $cat 'OK' "Espaço livre no disco $($v.DeviceID): $pct%"
        }
    }
    Add-Tabela 'Unidades de disco' $linhasDisco

    # --- Discos físicos: tipo, saúde, idade
    $linhasFis = @()
    try {
        $numSist = $null
        try { $numSist = (Get-Partition -DriveLetter $env:SystemDrive.Substring(0, 1) -ErrorAction Stop).DiskNumber } catch { }
        foreach ($pd in (Get-PhysicalDisk -ErrorAction Stop)) {
            $rel = $null
            try { $rel = $pd | Get-StorageReliabilityCounter -ErrorAction Stop } catch { }
            $tipo = [string]$pd.MediaType
            if ($tipo -eq 'Unspecified' -and [string]$pd.BusType -eq 'NVMe') { $tipo = 'SSD' }
            $ehSist = ($null -ne $numSist -and [string]$pd.DeviceId -eq [string]$numSist)
            $horas = $null; if ($rel -and $rel.PowerOnHours) { $horas = [int]$rel.PowerOnHours }
            $nome = "$($pd.FriendlyName) ($tipo, $(Format-Bytes $pd.Size))"
            $linhasFis += [pscustomobject]@{
                Disco = $pd.FriendlyName; Tipo = $tipo; Barramento = [string]$pd.BusType; Tamanho = (Format-Bytes $pd.Size)
                'Saúde' = [string]$pd.HealthStatus; 'Horas ligado' = $(if ($horas) { '{0:N0} h' -f $horas } else { '-' })
                'Desgaste' = $(if ($rel -and $null -ne $rel.Wear) { "$($rel.Wear)%" } else { '-' })
                'Temperatura' = $(if ($rel -and $rel.Temperature) { "$($rel.Temperature) °C" } else { '-' })
            }
            if ([string]$pd.HealthStatus -ne 'Healthy' -and $pd.HealthStatus) {
                Add-Achado $cat 'Critico' "Disco com saúde '$($pd.HealthStatus)': $($pd.FriendlyName)" "O próprio disco reporta problema (status operacional: $($pd.OperationalStatus))." 'Faça backup imediatamente e programe a troca do disco.'
            }
            if ($rel -and $rel.ReadErrorsUncorrected -gt 0) {
                Add-Achado $cat 'Critico' "Erros de leitura não corrigidos: $($pd.FriendlyName)" "$($rel.ReadErrorsUncorrected) erro(s) de leitura não corrigidos registrados pelo disco." 'Faça backup e programe a troca do disco.' @(Get-CorrecaoPadrao 'VerificarDisco')
            }
            if ($rel -and $null -ne $rel.Wear -and $rel.Wear -ge 80) {
                $s = 'Atencao'; if ($rel.Wear -ge 90) { $s = 'Critico' }
                Add-Achado $cat $s "SSD com $($rel.Wear)% de desgaste: $($pd.FriendlyName)" 'O SSD está perto do fim da vida útil estimada.' 'Programe a troca do SSD e mantenha backup em dia.'
            }
            if ($tipo -eq 'HDD' -and $ehSist) {
                $ev = "Windows instalado em HD mecânico: $nome."
                $titulo = 'Windows rodando em HD mecânico (lento)'
                if ($horas -and $horas -gt 26000) { $ev += (' O disco já tem {0:N0} horas ligado (~{1:N1} anos de uso contínuo).' -f $horas, ($horas / 8760)); $titulo = 'HD mecânico antigo com o Windows' }
                Add-Achado $cat 'Atencao' $titulo $ev 'Trocar o HD por um SSD é a melhoria de desempenho mais efetiva e barata para esta máquina.'
            } elseif ($tipo -eq 'HDD' -and $horas -and $horas -gt 35000) {
                Add-Achado $cat 'Info' "HD antigo: $($pd.FriendlyName)" ('{0:N0} horas ligado.' -f $horas) 'Mantenha backup; HDs com muitos anos de uso têm maior risco de falha.'
            }
        }
    } catch { }
    try {
        $fp = @(Get-CimInstance -Namespace root/wmi -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction Stop | Where-Object { $_.PredictFailure })
        if ($fp) { Add-Achado $cat 'Critico' 'SMART prevê falha do disco' "O disco ($($fp[0].InstanceName)) informou falha iminente pelo SMART." 'Faça backup imediatamente e troque o disco.' }
    } catch { }
    if ($linhasFis) { Add-Tabela 'Discos físicos' $linhasFis }

    # --- Inicialização (medida pelo próprio Windows)
    $logPerf = 'Microsoft-Windows-Diagnostics-Performance/Operational'
    $deg = @(Get-Eventos -Log $logPerf -Ids 101, 102, 103, 106, 109 -Dias 60 -Max 400)
    $grupos = @{}
    foreach ($e in $deg) {
        $d = Get-DadosEvento $e
        $nome = $d['FriendlyName']; if (-not $nome) { $nome = $d['Name'] }
        $arq = [string]$d['Name']; if (-not $arq) { $arq = [string]$d['FileName'] }
        $ms = 0; [void][int]::TryParse([string]$d['DegradationTime'], [ref]$ms)
        if (-not $nome) { continue }
        $tipoItem = switch ($e.Id) { 101 { 'Programa' } 102 { 'Driver' } 103 { 'Serviço' } 106 { 'Tarefa em segundo plano' } 109 { 'Dispositivo' } default { 'Outro' } }
        if (-not $grupos.ContainsKey($nome)) { $grupos[$nome] = [pscustomobject]@{ Nome = $nome; Tipo = $tipoItem; Arquivo = $arq; Ms = 0; Vezes = 0 } }
        $grupos[$nome].Ms += $ms; $grupos[$nome].Vezes++
        if ($arq -and $arq -match '\.exe$') {
            $k = ([IO.Path]::GetFileName($arq)).ToLower()
            if (-not $Script:Dados.AtrasoBoot.ContainsKey($k)) { $Script:Dados.AtrasoBoot[$k] = 0 }
            $Script:Dados.AtrasoBoot[$k] += $ms
        }
    }
    $piores = @($grupos.Values | Sort-Object Ms -Descending | Select-Object -First 8)
    if ($piores) {
        Add-Tabela 'O que mais atrasou a inicialização (60 dias)' ($piores | ForEach-Object {
                [pscustomobject]@{ Item = $_.Nome; Tipo = $_.Tipo; Arquivo = $_.Arquivo; 'Vezes' = $_.Vezes; 'Atraso total' = ('{0:N1} s' -f ($_.Ms / 1000)) } })
    }
    $boots = @(Get-Eventos -Log $logPerf -Ids 100 -Dias 30 -Max 5)
    $txtPiores = ($piores | Select-Object -First 4 | ForEach-Object { '{0} (+{1:N1} s)' -f $_.Nome, ($_.Ms / 1000) }) -join '; '
    if ($boots) {
        $d = Get-DadosEvento $boots[0]
        $bt = [double]$d['BootTime'] / 1000; $main = [double]$d['MainPathBootTime'] / 1000; $post = [double]$d['BootPostBootTime'] / 1000
        $media = (($boots | ForEach-Object { [double](Get-DadosEvento $_)['BootTime'] / 1000 }) | Measure-Object -Average).Average
        $ev = ('Último boot: {0:N0} s (até a área de trabalho: {1:N0} s; carregamento depois disso: {2:N0} s). Média dos últimos {3}: {4:N0} s.' -f $bt, $main, $post, $boots.Count, $media)
        if ($txtPiores) { $ev += " O que mais atrasou: $txtPiores." }
        if ($media -ge 120) {
            $sev = 'Atencao'; if ($media -ge 240) { $sev = 'Critico' }
            Add-Achado $cat $sev ('Inicialização lenta ({0:N0} s em média)' -f $media) $ev 'Desative da inicialização programas que não precisam abrir com o Windows; em HD mecânico, troque por SSD.' @(Get-CorrecaoPadrao 'Inicializacao')
        } else {
            Add-Achado $cat 'OK' ('Inicialização em {0:N0} s (média)' -f $media)
        }
    } elseif (-not $Script:EhAdmin) {
        Add-Achado $cat 'Info' 'Tempo de inicialização não medido' 'Ler o histórico de desempenho do Windows exige administrador.'
    }

    # --- Programas de inicialização
    $itens = @(Get-ItensInicializacao)
    $ativos = @($itens | Where-Object { $_.Ativo })
    Add-Tabela 'Programas de inicialização' ($itens | Sort-Object @{ e = { -not $_.Ativo } }, @{ e = { -$_.AtrasoMs } }, Nome | ForEach-Object {
            $cmd = $_.Comando; if ($cmd.Length -gt 110) { $cmd = $cmd.Substring(0, 110) + '...' }
            [pscustomobject]@{ Programa = $_.Nome; Status = $(if ($_.Ativo) { 'Ativo' } else { 'Desativado' }); Escopo = $_.Escopo
                'Atraso medido' = $(if ($_.AtrasoMs) { '{0:N1} s' -f ($_.AtrasoMs / 1000) } else { '-' }); Comando = $cmd }
        })
    if ($ativos.Count -ge 10) {
        $nomes = ($ativos | Sort-Object AtrasoMs -Descending | Select-Object -First 8 | ForEach-Object { $_.Nome }) -join ', '
        Add-Achado $cat 'Atencao' "$($ativos.Count) programas iniciam com o Windows" "Ativos: $nomes..." 'Desative os que não são essenciais (atualizadores, mensageiros, launchers de jogos, utilitários de fabricante).' @(Get-CorrecaoPadrao 'Inicializacao')
    } else {
        Add-Achado $cat 'OK' "$($ativos.Count) programa(s) na inicialização"
    }

    # --- Plano de energia
    $plano = Get-PlanoEnergiaAtivo
    if ($plano.Guid -eq $Script:GuidPlanos.Economia) {
        Add-Achado $cat 'Atencao' 'Plano de energia em "Economia de energia"' 'Esse plano limita a velocidade do processador o tempo todo.' 'Use o plano Equilibrado.' @(Get-CorrecaoPadrao 'PlanoEquilibrado')
    } elseif ($plano.Nome) {
        Add-Achado $cat 'OK' "Plano de energia: $($plano.Nome)"
    }
}

#endregion

#region ===================== Diagnóstico: Travamentos =====================

function Test-Travamentos {
    Write-Etapa 'Travamentos: telas azuis, desligamentos inesperados, erros de hardware e de disco (30 dias)'
    $cat = 'Travamentos'
    $dias = 30

    $dumps = @(Get-ChildItem "$env:SystemRoot\Minidump\*.dmp" -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -gt (Get-Date).AddDays(-$dias) })
    $crashes = New-Object System.Collections.Generic.List[object]
    foreach ($e in (Get-Eventos -Log 'System' -Ids 1001 -Provider 'Microsoft-Windows-WER-SystemErrorReporting' -Dias $dias -Max 50)) {
        $txt = ''; if ($e.Properties.Count -gt 0) { $txt = [string]$e.Properties[0].Value }
        if (-not $txt) { $txt = [string]$e.Message }
        if ($txt -match '0x([0-9a-fA-F]{1,16})') { $crashes.Add([pscustomobject]@{ Data = $e.TimeCreated; Codigo = (ConvertTo-CodigoBugCheck $Matches[1]) }) }
    }

    # Kernel-Power 41: reinício sem desligamento limpo
    $inesperados = New-Object System.Collections.Generic.List[object]
    foreach ($e in (Get-Eventos -Log 'System' -Ids 41 -Provider 'Microsoft-Windows-Kernel-Power' -Dias $dias -Max 100)) {
        $c = Get-CausaKernelPower $e
        if ($c.Codigo) {
            $cod = $c.Codigo
            $dup = $crashes | Where-Object { $_.Codigo -eq $cod -and [math]::Abs(($_.Data - $e.TimeCreated).TotalMinutes) -lt 30 }
            if (-not $dup) { $crashes.Add([pscustomobject]@{ Data = $e.TimeCreated; Codigo = $cod }) }
        } else {
            # Às vezes o evento 41 vem sem o código, mas a tela azul foi registrada no mesmo boot (evento 1001)
            $perto = $crashes | Where-Object { [math]::Abs(($_.Data - $e.TimeCreated).TotalMinutes) -lt 15 } | Select-Object -First 1
            if ($perto) {
                $nome = ''; if ($Script:BugChecks.ContainsKey($perto.Codigo)) { $nome = ' ' + $Script:BugChecks[$perto.Codigo][0] }
                $c = [pscustomobject]@{ Texto = "tela azul (0x$($perto.Codigo)$nome)"; Codigo = $perto.Codigo }
            }
        }
        $inesperados.Add([pscustomobject]@{ Data = $e.TimeCreated; Causa = $c.Texto; Contexto = (Get-ContextoTravamento $e.TimeCreated) })
    }

    # Telas azuis agrupadas por código
    if ($crashes.Count -gt 0) {
        foreach ($g in ($crashes | Group-Object Codigo | Sort-Object Count -Descending)) {
            $info = $Script:BugChecks[$g.Name]
            if (-not $info) { $info = @('CÓDIGO NÃO CATALOGADO', 'Código não está na tabela da ferramenta; pesquise "bug check 0x' + $g.Name + '" na documentação da Microsoft.', 'Driver') }
            $datas = ($g.Group | Sort-Object Data -Descending | Select-Object -First 6 | ForEach-Object { $_.Data.ToString('dd/MM HH:mm') }) -join ', '
            $ev = "Ocorreu $($g.Count) vez(es) em $dias dias ($datas). Causa provável: $($info[1])"
            if ($dumps) { $ev += " Arquivos de despejo para análise: $env:SystemRoot\Minidump ($($dumps.Count) recentes)." }
            $corr = switch ($info[2]) {
                'Driver' { @(Get-CorrecaoPadrao 'AtualizarDrivers') }
                'Memoria' { @(Get-CorrecaoPadrao 'TesteMemoria'), @(Get-CorrecaoPadrao 'AtualizarDrivers') }
                'Disco' { @(Get-CorrecaoPadrao 'VerificarDisco') }
                'Sistema' { @(Get-CorrecaoPadrao 'ReparoWindows') }
                'Hardware' { @(Get-CorrecaoPadrao 'TesteMemoria') }
            }
            Add-Achado $cat 'Critico' "Tela azul: $($info[0]) (0x$($g.Name))" $ev $Script:AcaoPorTipo[$info[2]] @($corr)
        }
    } else {
        Add-Achado $cat 'OK' "Nenhuma tela azul em $dias dias"
    }

    # Desligamentos inesperados (sem tela azul)
    $semBsod = @($inesperados | Where-Object { $_.Causa -notlike 'tela azul*' })
    if ($inesperados.Count -gt 0) {
        Add-Tabela 'Desligamentos inesperados (30 dias)' ($inesperados | ForEach-Object {
                $q = '-'
                if ($_.Contexto) {
                    switch ($_.Contexto.Momento) {
                        'Desligar' { $q = 'Ao desligar (pedido às {0:dd/MM HH:mm})' -f $_.Contexto.Pedido }
                        'Suspender' { $q = 'Ao suspender (pedido às {0:dd/MM HH:mm})' -f $_.Contexto.Pedido }
                        default { $q = 'Em uso (último registro {0:dd/MM HH:mm})' -f $_.Contexto.UltimoRegistro }
                    }
                }
                [pscustomobject]@{ 'Detectado em' = $_.Data.ToString('dd/MM/yyyy HH:mm'); 'Quando travou' = $q; 'Causa provável' = $_.Causa } })
    }
    if ($semBsod.Count -gt 0) {
        $sev = 'Atencao'; if ($semBsod.Count -ge 3) { $sev = 'Critico' }
        $causas = ($semBsod | Group-Object Causa | ForEach-Object { "$($_.Count)x $($_.Name)" }) -join '; '
        $rec = 'Se foi o botão de energia, investigue o congelamento (drivers, disco, temperatura). Se foi falta de energia, verifique tomada, nobreak, fonte e, em notebooks, a bateria e o carregador.'
        Add-Achado $cat $sev "$($semBsod.Count) desligamento(s) inesperado(s) em $dias dias" "Causas: $causas. Último em $($semBsod[0].Data.ToString('dd/MM/yyyy HH:mm'))." $rec
    } elseif ($inesperados.Count -eq 0) {
        Add-Achado $cat 'OK' "Nenhum desligamento inesperado em $dias dias"
    }

    # Erros de hardware (WHEA)
    $whea = @(Get-Eventos -Log 'System' -Provider 'Microsoft-Windows-WHEA-Logger' -Dias $dias -Max 200)
    if ($whea) {
        $graves = @($whea | Where-Object { $_.Level -le 2 })
        $sev = 'Atencao'; if ($graves) { $sev = 'Critico' }
        $ex = Get-PrimeiraLinha $whea[0].Message 200
        Add-Achado $cat $sev "Erros de hardware registrados ($($whea.Count) em $dias dias)" "O Windows registrou erros de hardware (WHEA): $($graves.Count) grave(s). Último: $ex" $Script:AcaoPorTipo['Hardware'] @(Get-CorrecaoPadrao 'TesteMemoria')
    } else {
        Add-Achado $cat 'OK' 'Nenhum erro de hardware (WHEA) registrado'
    }

    # Erros de disco e sistema de arquivos
    $descr = @{ 7 = 'setor defeituoso (bad block)'; 11 = 'erro da controladora de disco'; 51 = 'erro ao gravar/ler a paginação'; 153 = 'operações de leitura/gravação repetidas (disco lento ou falhando)'; 129 = 'reset do dispositivo de armazenamento por tempo esgotado'; 55 = 'estrutura do sistema de arquivos corrompida'; 98 = 'volume precisa de verificação'; 140 = 'falha ao gravar no volume' }
    $provs = 'disk', 'Ntfs', 'Microsoft-Windows-Ntfs', 'stornvme', 'storahci', 'iaStorA', 'iaStorAC', 'iaStorAVC', 'iaStorV'
    $errDisco = @()
    foreach ($p in $provs) { $errDisco += Get-Eventos -Log 'System' -Provider $p -Ids 7, 11, 51, 153, 129, 55, 98, 140 -Dias $dias -Max 200 -NivelMax 3 }
    if ($errDisco) {
        $graves = @($errDisco | Where-Object { $_.Id -in 7, 11, 51, 55 })
        $sev = 'Atencao'; if ($graves) { $sev = 'Critico' }
        $resumo = ($errDisco | Group-Object Id | ForEach-Object { "{0}x {1}" -f $_.Count, $descr[[int]$_.Name] }) -join '; '
        Add-Achado $cat $sev "Erros de disco registrados ($($errDisco.Count) em $dias dias)" "$resumo. Último em $(($errDisco | Sort-Object TimeCreated -Descending)[0].TimeCreated.ToString('dd/MM/yyyy HH:mm'))." 'Faça backup, verifique o disco (CHKDSK) e a saúde SMART. Em desktops, troque o cabo SATA; se persistir, troque o disco.' @(Get-CorrecaoPadrao 'VerificarDisco')
    } else {
        Add-Achado $cat 'OK' 'Nenhum erro de disco registrado'
    }

    # Padrão: o PC trava ao desligar ou ao suspender (a tela azul só aparece no boot seguinte)
    $comCtx = @($inesperados | Where-Object { $_.Contexto })
    $aoDesl = @($comCtx | Where-Object { $_.Contexto.Momento -eq 'Desligar' })
    $aoSusp = @($comCtx | Where-Object { $_.Contexto.Momento -eq 'Suspender' })
    $exemplos = { param($l) ($l | Sort-Object Data -Descending | Select-Object -First 4 | ForEach-Object { '{0:dd/MM HH:mm}' -f $_.Contexto.Pedido }) -join ', ' }
    if ($aoDesl.Count -ge 2) {
        $pct = [math]::Round($aoDesl.Count / $comCtx.Count * 100)
        $ev = "$($aoDesl.Count) de $($comCtx.Count) travamentos ($pct%) aconteceram com o Windows já desligando (pedidos de desligamento em $(& $exemplos $aoDesl)). A tela azul só aparece no registro quando o PC é ligado de novo, por isso parece que 'deu problema ao ligar'."
        $rec = ''
        $corr = @()
        if ($Script:Dados.Resumo.InicioRapido) {
            $ev += ' A Inicialização Rápida está ativa: ao "Desligar", o Windows grava a memória do sistema no disco (como numa hibernação), justamente no momento em que trava.'
            $rec = 'Desative a Inicialização Rápida e acompanhe por uma semana: se as telas azuis pararem, a causa está confirmada. '
            $corr += Get-CorrecaoPadrao 'InicioRapido'
        }
        if ($errDisco) {
            $ev += " O disco também registrou $($errDisco.Count) erro(s) no período, o que aponta para o armazenamento (SSD, cabo ou controladora)."
            $corr += Get-CorrecaoPadrao 'VerificarDisco'
        }
        $rec += 'Atualize os drivers de chipset, armazenamento e rede (dispositivos que precisam desligar corretamente). Se continuar, troque o cabo SATA ou o disco.'
        $corr += Get-CorrecaoPadrao 'AtualizarDrivers'
        Add-Achado $cat 'Critico' 'O PC trava ao desligar' $ev $rec $corr
    }
    if ($aoSusp.Count -ge 2) {
        Add-Achado $cat 'Critico' 'O PC trava ao entrar em suspensão' "$($aoSusp.Count) travamento(s) aconteceram ao suspender (pedidos em $(& $exemplos $aoSusp))." 'Atualize os drivers de vídeo, rede, Wi-Fi e USB/dock. Enquanto isso, mude o plano de energia para não suspender automaticamente.' @(Get-CorrecaoPadrao 'AtualizarDrivers')
    }

    # Programas que travam ou fecham sozinhos (14 dias)
    $app = @()
    $app += Get-Eventos -Log 'Application' -Ids 1000 -Provider 'Application Error' -Dias 14 -Max 300
    $app += Get-Eventos -Log 'Application' -Ids 1002 -Provider 'Application Hang' -Dias 14 -Max 300
    $porApp = $app | Where-Object { $_.Properties.Count -gt 0 } | Group-Object { ([string]$_.Properties[0].Value).ToLower() } | Where-Object { $_.Count -ge 3 } | Sort-Object Count -Descending | Select-Object -First 5
    foreach ($g in $porApp) {
        $falhas = @($g.Group | Where-Object { $_.Id -eq 1000 }).Count
        $congel = @($g.Group | Where-Object { $_.Id -eq 1002 }).Count
        $modulo = ($g.Group | Where-Object { $_.Id -eq 1000 -and $_.Properties.Count -gt 3 } | Group-Object { [string]$_.Properties[3].Value } | Sort-Object Count -Descending | Select-Object -First 1).Name
        $ev = "$falhas fechamento(s) por erro e $congel congelamento(s) em 14 dias."
        if ($modulo) { $ev += " Módulo mais envolvido: $modulo." }
        $corr = @(); $rec = 'Atualize ou repare o programa (Aplicativos > Modificar/Reparar). Se o módulo for um plug-in ou driver, atualize-o.'
        if ($g.Name -in 'explorer.exe', 'svchost.exe', 'dwm.exe', 'searchhost.exe', 'shellexperiencehost.exe', 'startmenuexperiencehost.exe') { $corr = @(Get-CorrecaoPadrao 'ReparoWindows'); $rec = 'Componente do Windows instável: rode o reparo do Windows e instale as atualizações.' }
        Add-Achado $cat 'Atencao' "$($g.Name) travou/fechou $($g.Count) vezes" $ev $rec $corr
    }
}

#endregion

#region ===================== Diagnóstico: Câmera e microfone =====================

function Get-UsoDispositivo {
    param([ValidateSet('webcam', 'microphone')][string]$Recurso)
    $res = @()
    $raizes = @(@{ Raiz = 'Registry::HKEY_LOCAL_MACHINE'; Usuario = 'Sistema' })
    foreach ($h in (Get-HivesUsuarios)) { $raizes += @{ Raiz = "Registry::HKEY_USERS\$($h.Sid)"; Usuario = $h.Usuario } }
    foreach ($r in $raizes) {
        $base = "$($r.Raiz)\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\$Recurso"
        if (-not (Test-Path -LiteralPath $base)) { continue }
        $chaves = @(Get-ChildItem -LiteralPath $base -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -ne 'NonPackaged' })
        $np = "$base\NonPackaged"
        if (Test-Path -LiteralPath $np) { $chaves += @(Get-ChildItem -LiteralPath $np -ErrorAction SilentlyContinue) }
        foreach ($k in $chaves) {
            $ini = $k.GetValue('LastUsedTimeStart'); $fim = $k.GetValue('LastUsedTimeStop')
            if ($null -eq $ini -or [int64]$ini -le 0) { continue }
            $empacotado = ($k.PSParentPath -notmatch 'NonPackaged$')
            $app = $k.PSChildName; if (-not $empacotado) { $app = $app -replace '#', '\' }
            $dtIni = $null; try { $dtIni = [DateTime]::FromFileTime([int64]$ini) } catch { }
            $dtFim = $null; if ($fim -and [int64]$fim -gt 0) { try { $dtFim = [DateTime]::FromFileTime([int64]$fim) } catch { } }
            $res += [pscustomobject]@{ App = $app; Usuario = $r.Usuario; Empacotado = $empacotado; EmUso = ($null -eq $dtFim); Inicio = $dtIni; Fim = $dtFim }
        }
    }
    return $res
}

function Test-Privacidade {
    param([string]$Recurso, [string]$Nome, [string]$Politica)
    $cat = 'Câmera e microfone'
    $problemas = 0
    $pol = (Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy' -ErrorAction SilentlyContinue).$Politica
    if ($pol -eq 2) {
        $problemas++
        Add-Achado $cat 'Critico' "$Nome bloqueado(a) por política da empresa" "Política $Politica = 2 (negar) em HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy." 'Essa configuração vem de GPO/Intune: ajuste a política no servidor. Alterar localmente seria desfeito na próxima atualização de políticas.'
    }
    if ($Recurso -eq 'webcam' -and (Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Camera' -ErrorAction SilentlyContinue).AllowCamera -eq 0) {
        $problemas++
        Add-Achado $cat 'Critico' 'Câmera desativada por política (AllowCamera = 0)' 'HKLM\SOFTWARE\Policies\Microsoft\Camera\AllowCamera = 0.' 'Ajuste a política "Permitir uso da câmera" na GPO/Intune.'
    }
    $hklm = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\$Recurso"
    if ((Get-ItemProperty -LiteralPath $hklm -ErrorAction SilentlyContinue).Value -eq 'Deny') {
        $problemas++
        $ch = New-Correcao "Priv:$hklm" "Liberar $Nome neste dispositivo" "Liga 'Acesso a $Nome neste dispositivo' nas configurações de privacidade." 'Reversivel' {
            param($p) Set-RegistroSeguro -Caminho $p.Caminho -Nome 'Value' -Valor 'Allow' -Tipo String -Descricao $p.Desc } @{ Caminho = $hklm; Desc = "Privacidade: $Nome liberado(a) no dispositivo" }
        Add-Achado $cat 'Critico' "$Nome desligado(a) na privacidade para todo o computador" "Configurações > Privacidade > $Nome > 'Acesso a $Nome neste dispositivo' está desativado." "Ative o acesso à $Nome." @($ch)
    }
    foreach ($h in (Get-HivesUsuarios)) {
        $base = "Registry::HKEY_USERS\$($h.Sid)\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\$Recurso"
        if ((Get-ItemProperty -LiteralPath $base -ErrorAction SilentlyContinue).Value -eq 'Deny') {
            $problemas++
            $ch = New-Correcao "Priv:$base" "Liberar $Nome para os aplicativos ($($h.Usuario))" "Liga 'Permitir que aplicativos acessem $Nome' para o usuário." 'Reversivel' {
                param($p) Set-RegistroSeguro -Caminho $p.Caminho -Nome 'Value' -Valor 'Allow' -Tipo String -Descricao $p.Desc } @{ Caminho = $base; Desc = "Privacidade: $Nome liberado(a) para apps ($($h.Usuario))" }
            Add-Achado $cat 'Critico' "$Nome bloqueado(a) para os aplicativos de $($h.Usuario)" "'Permitir que aplicativos acessem $Nome' está desativado para este usuário." "Libere o acesso à $Nome." @($ch)
        }
        $np = "$base\NonPackaged"
        if ((Get-ItemProperty -LiteralPath $np -ErrorAction SilentlyContinue).Value -eq 'Deny') {
            $problemas++
            $ch = New-Correcao "Priv:$np" "Liberar $Nome para programas da área de trabalho ($($h.Usuario))" "Liga 'Permitir que aplicativos da área de trabalho acessem $Nome' (Zoom, Webex, navegadores, Teams clássico)." 'Reversivel' {
                param($p) Set-RegistroSeguro -Caminho $p.Caminho -Nome 'Value' -Valor 'Allow' -Tipo String -Descricao $p.Desc } @{ Caminho = $np; Desc = "Privacidade: $Nome liberado(a) para programas da área de trabalho ($($h.Usuario))" }
            Add-Achado $cat 'Critico' "$Nome bloqueado(a) para programas da área de trabalho ($($h.Usuario))" "Zoom, Webex, navegadores e Teams clássico não conseguem usar $Nome." "Libere 'Permitir que aplicativos da área de trabalho acessem $Nome'." @($ch)
        }
        foreach ($k in (Get-ChildItem -LiteralPath $base -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match 'Teams|Zoom|Skype|Webex|WindowsCamera|Whatsapp' })) {
            if ($k.GetValue('Value') -eq 'Deny') {
                $problemas++
                $app = ($k.PSChildName -split '_')[0]
                $ch = New-Correcao "Priv:$($k.Name)" "Liberar $Nome para $app ($($h.Usuario))" "Liga a permissão individual do aplicativo $app." 'Reversivel' {
                    param($p) Set-RegistroSeguro -Caminho $p.Caminho -Nome 'Value' -Valor 'Allow' -Tipo String -Descricao $p.Desc } @{ Caminho = "Registry::$($k.Name)"; Desc = "Privacidade: $Nome liberado(a) para $app ($($h.Usuario))" }
                Add-Achado $cat 'Atencao' "$app sem permissão para usar $Nome ($($h.Usuario))" "A permissão individual do aplicativo $app está desligada." "Libere $Nome para o $app." @($ch)
            }
        }
    }
    return $problemas
}

function Test-CameraMicrofone {
    Write-Etapa 'Câmera e microfone: dispositivos, privacidade, políticas e programas usando agora'
    $cat = 'Câmera e microfone'

    # --- Câmeras
    $filtro = { $_.FriendlyName -notmatch 'scan|print|impress|fax|mfp|wia' }
    $presentes = @(Get-PnpDevice -Class Camera, Image -PresentOnly -ErrorAction SilentlyContinue | Where-Object $filtro)
    $todas = @(Get-PnpDevice -Class Camera, Image -ErrorAction SilentlyContinue | Where-Object $filtro)
    $camOk = $false
    foreach ($c in $presentes) {
        $cod = [int]$c.ConfigManagerErrorCode
        if ($cod -eq 22) {
            $ch = New-Correcao "Dev:$($c.InstanceId)" "Ativar câmera: $($c.FriendlyName)" 'Reativa o dispositivo no Gerenciador de Dispositivos.' 'Reversivel' {
                param($p) Enable-DispositivoSeguro -InstanceId $p.Id -Nome $p.Nome } @{ Id = $c.InstanceId; Nome = $c.FriendlyName }
            Add-Achado $cat 'Critico' "Câmera desativada: $($c.FriendlyName)" 'O dispositivo está desativado no Gerenciador de Dispositivos.' 'Ative o dispositivo.' @($ch)
        } elseif ($cod -ne 0) {
            Add-Achado $cat 'Critico' "Câmera com erro: $($c.FriendlyName)" "Código $cod ($($Script:CodigosDispositivo[$cod]))." 'Reinstale/atualize o driver da câmera; em notebooks, use o driver do fabricante.' @(Get-CorrecaoPadrao 'AtualizarDrivers')
        } else { $camOk = $true }
    }
    if ($camOk) { Add-Achado $cat 'OK' ("Câmera funcionando: " + (($presentes | Where-Object { $_.ConfigManagerErrorCode -eq 0 } | ForEach-Object FriendlyName) -join ', ')) }
    if (-not $presentes) {
        $antigas = @($todas | Where-Object { $_.InstanceId -notin $presentes.InstanceId })
        if ($antigas) {
            Add-Achado $cat 'Atencao' 'Câmera não detectada agora' "A câmera já foi instalada nesta máquina ($(($antigas | Select-Object -First 2 | ForEach-Object FriendlyName) -join ', ')), mas não está presente agora." 'Verifique a tecla de privacidade da câmera (ex.: Fn+F8/F10), o obturador físico, o cabo/porta USB ou se ela foi desativada no BIOS.'
        } else {
            Add-Achado $cat 'Info' 'Nenhuma câmera instalada' 'Nenhum dispositivo de câmera foi encontrado.' 'Se deveria haver uma câmera, confira cabo/USB, BIOS e a tecla de privacidade.'
        }
    }

    # --- Microfones (endpoints de captura de áudio)
    $mics = @()
    $capRoot = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture'
    foreach ($k in (Get-ChildItem $capRoot -ErrorAction SilentlyContinue)) {
        $estado = [int](Get-ItemProperty $k.PSPath -ErrorAction SilentlyContinue).DeviceState -band 0xF
        $props = Get-ItemProperty (Join-Path $k.PSPath 'Properties') -ErrorAction SilentlyContinue
        $nome = $props.'{a45c254e-df1c-4efd-8020-67d146a850e0},2'
        $placa = $props.'{b3f8fa53-0004-438e-9003-51a46e139bfc},6'
        if (-not $nome) { $nome = 'Microfone' }
        if ($placa) { $nome = "$nome ($placa)" }
        $mics += [pscustomobject]@{ Guid = $k.PSChildName; Nome = $nome; Estado = $estado }
    }
    $ativos = @($mics | Where-Object { $_.Estado -eq 1 })
    $desativados = @($mics | Where-Object { $_.Estado -eq 2 })
    foreach ($m in $desativados) {
        $pnp = Get-PnpDevice -Class AudioEndpoint -ErrorAction SilentlyContinue | Where-Object { $_.InstanceId -like "*$($m.Guid)*" } | Select-Object -First 1
        $corr = @()
        if ($pnp) {
            $corr += New-Correcao "Dev:$($pnp.InstanceId)" "Ativar microfone: $($m.Nome)" 'Reativa o microfone desativado.' 'Reversivel' {
                param($p) Enable-DispositivoSeguro -InstanceId $p.Id -Nome $p.Nome } @{ Id = $pnp.InstanceId; Nome = $m.Nome }
        }
        $sev = 'Info'; if (-not $ativos) { $sev = 'Critico' }
        Add-Achado $cat $sev "Microfone desativado: $($m.Nome)" 'Desativado em Som > Gravação.' 'Ative em Configurações > Sistema > Som > Todos os dispositivos de som, ou pela correção automática.' $corr
    }
    if ($ativos) {
        Add-Achado $cat 'OK' ("Microfone ativo: " + (($ativos | ForEach-Object Nome) -join ', '))
    } elseif (-not $desativados) {
        $desc = @($mics | Where-Object { $_.Estado -eq 8 } | ForEach-Object Nome | Select-Object -Unique)
        if ($desc) {
            Add-Achado $cat 'Info' 'Nenhum microfone conectado' "Entradas de áudio sem nada plugado: $($desc -join ', ')." 'Se o usuário usa headset/microfone, conecte-o (ou teste outra porta) e rode a análise de novo.'
        } else {
            Add-Achado $cat 'Atencao' 'Nenhum microfone disponível' 'Nenhuma entrada de áudio foi encontrada.' 'Conecte o headset/microfone e verifique o driver de áudio.'
        }
    }

    # --- Serviços de áudio e câmera
    foreach ($s in 'AudioEndpointBuilder', 'Audiosrv') {
        $svc = Get-Service -Name $s -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -ne 'Running') {
            $ch = New-Correcao "Svc:$s" "Iniciar serviço $($svc.DisplayName)" 'Coloca o serviço em automático e o inicia.' 'Reversivel' {
                param($p) Set-ServicoSeguro -Nome $p.Nome -Inicio Automatic -Iniciar } @{ Nome = $s }
            Add-Achado $cat 'Critico' "Serviço de áudio parado: $($svc.DisplayName)" "Status: $($svc.Status); inicialização: $($svc.StartType)." 'Sem esse serviço não há som nem microfone.' @($ch)
        }
    }
    $fs = Get-Service -Name FrameServer -ErrorAction SilentlyContinue
    if ($fs -and $fs.StartType -eq 'Disabled') {
        $ch = New-Correcao 'Svc:FrameServer' 'Reativar o serviço Windows Camera Frame Server' 'Volta a inicialização do serviço para Manual (padrão do Windows).' 'Reversivel' { Set-ServicoSeguro -Nome FrameServer -Inicio Manual }
        Add-Achado $cat 'Atencao' 'Serviço de câmera do Windows desativado' 'Windows Camera Frame Server está desativado; vários aplicativos não conseguem abrir a câmera.' 'Reative o serviço (inicialização Manual).' @($ch)
    }

    # --- Privacidade e políticas
    $p1 = Test-Privacidade -Recurso 'webcam' -Nome 'câmera' -Politica 'LetAppsAccessCamera'
    $p2 = Test-Privacidade -Recurso 'microphone' -Nome 'microfone' -Politica 'LetAppsAccessMicrophone'
    if (($p1 + $p2) -eq 0) { Add-Achado $cat 'OK' 'Privacidade e políticas liberam câmera e microfone' }

    # --- Quem está usando a câmera/microfone agora
    $procs = @(Get-Process -ErrorAction SilentlyContinue)
    foreach ($rec in @(@{ R = 'webcam'; N = 'câmera' }, @{ R = 'microphone'; N = 'microfone' })) {
        $usos = @(Get-UsoDispositivo $rec.R)
        $emUso = @()
        foreach ($u in ($usos | Where-Object { $_.EmUso })) {
            if ($u.Empacotado) {
                if ($u.Inicio -and $u.Inicio -gt (Get-Date).AddHours(-12)) { $emUso += $u }
            } else {
                $proc = $procs | Where-Object { $_.Path -and $_.Path -ieq $u.App } | Select-Object -First 1
                if ($proc) { $emUso += $u }
            }
        }
        foreach ($u in $emUso) {
            $exe = $u.App; if (-not $u.Empacotado) { $exe = Split-Path $u.App -Leaf } else { $exe = ($u.App -split '_')[0] }
            $desde = ''; if ($u.Inicio) { $desde = " desde $($u.Inicio.ToString('dd/MM HH:mm'))" }
            if ($rec.R -eq 'webcam') {
                $corr = @()
                if (-not $u.Empacotado) {
                    $corr += New-Correcao "Fechar:$($u.App)" "Fechar $exe (está segurando a câmera)" "Encerra o $exe para liberar a câmera. Trabalho não salvo nele será perdido." 'NaoReversivel' {
                        param($p) Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Path -ieq $p.Caminho } | Stop-Process -Force -ErrorAction Stop } @{ Caminho = $u.App } -Interativa
                }
                Add-Achado $cat 'Atencao' "Câmera em uso agora por $exe" "$exe ($($u.Usuario)) está usando a câmera$desde. Só um programa por vez consegue usar a câmera: é o clássico 'a câmera não abre no Teams porque o Zoom está aberto'." "Feche o $exe (ou encerre a chamada nele) antes de abrir a câmera em outro programa." $corr
            } else {
                Add-Achado $cat 'Info' "Microfone em uso agora por $exe" "$exe ($($u.Usuario)) está usando o microfone$desde." 'Normalmente o microfone pode ser compartilhado; se outro programa não captar áudio, feche este.'
            }
        }
        $ultimo = $usos | Where-Object { $_.Fim } | Sort-Object Fim -Descending | Select-Object -First 1
        if ($ultimo -and -not $emUso) {
            $nomeApp = $ultimo.App; if (-not $ultimo.Empacotado) { $nomeApp = Split-Path $ultimo.App -Leaf }
            Add-Achado $cat 'OK' "$($rec.N.Substring(0,1).ToUpper() + $rec.N.Substring(1)) livre (último uso: $nomeApp em $($ultimo.Fim.ToString('dd/MM HH:mm')))"
        }
    }

    # --- Vários programas de reunião abertos
    $reuniao = @{ 'ms-teams' = 'Teams'; 'teams' = 'Teams (clássico)'; 'zoom' = 'Zoom'; 'ciscocollabhost' = 'Webex'; 'webexmta' = 'Webex'; 'atmgr' = 'Webex'; 'skype' = 'Skype'; 'lync' = 'Skype for Business'; 'discord' = 'Discord'; 'obs64' = 'OBS'; 'slack' = 'Slack'; 'g2mcomm' = 'GoTo Meeting'; 'gotomeeting' = 'GoTo Meeting' }
    $abertos = @($procs | ForEach-Object { $reuniao[$_.ProcessName.ToLower()] } | Where-Object { $_ } | Select-Object -Unique)
    if ($abertos.Count -ge 2) {
        Add-Achado $cat 'Info' "Vários programas de reunião abertos: $($abertos -join ', ')" 'Programas de reunião costumam reservar câmera e microfone mesmo minimizados na bandeja.' 'Se a câmera/microfone não funcionar em um deles, feche completamente os outros (inclusive pela bandeja do relógio).'
    }
}

#endregion

#region ===================== Diagnóstico: Rede =====================

function Get-InfoWifi {
    $saida = @(netsh wlan show interfaces 2>$null)
    $txt = $saida -join "`n"
    if ($txt -match 'location|localiza') { return [pscustomobject]@{ Erro = 'O Windows exige permissão de localização para ler o sinal do Wi-Fi.' } }
    $campos = [ordered]@{}
    foreach ($l in $saida) {
        if ($l -match '^\s*([^:]+?)\s+:\s(.*)$') { $k = $Matches[1].Trim(); if (-not $campos.Contains($k)) { $campos[$k] = $Matches[2].Trim() } }
    }
    $pega = { param($re) foreach ($k in $campos.Keys) { if ($k -match $re) { return $campos[$k] } }; return $null }
    $sinal = & $pega '^(Signal|Sinal)$'
    return [pscustomobject]@{
        Erro = $null
        SSID = & $pega '^SSID$'
        Sinal = $(if ($sinal -match '(\d+)') { [int]$Matches[1] } else { $null })
        Canal = & $pega '^(Channel|Canal)$'
        Radio = & $pega '(Radio type|Tipo de r)'
        Banda = & $pega '^(Band|Banda)$'
        Taxa = & $pega '(Receive rate|Taxa de recep)'
    }
}

function Test-Rede {
    Write-Etapa 'Rede: conexão, IP, DNS, internet, Wi-Fi e velocidade do cabo'
    $cat = 'Rede'
    $adapt = @(Get-NetAdapter -Physical -ErrorAction SilentlyContinue)
    $ativos = @($adapt | Where-Object { $_.Status -eq 'Up' })

    foreach ($d in ($adapt | Where-Object { $_.Status -eq 'Disabled' })) {
        $ch = New-Correcao "Adapt:$($d.Name)" "Ativar o adaptador de rede '$($d.Name)'" $d.InterfaceDescription 'Reversivel' {
            param($p)
            Confirm-Protecao
            Add-Alteracao ([pscustomobject]@{ Tipo = 'Adaptador'; Nome = $p.Nome; Descricao = "Adaptador de rede ativado: $($p.Nome)" })
            Enable-NetAdapter -Name $p.Nome -Confirm:$false -ErrorAction Stop } @{ Nome = $d.Name }
        $sev = 'Info'; if (-not $ativos) { $sev = 'Critico' }
        Add-Achado $cat $sev "Adaptador de rede desativado: $($d.Name)" "$($d.InterfaceDescription) está desativado." 'Ative-o se for o adaptador usado para conectar.' @($ch)
    }
    if (-not $ativos) {
        Add-Achado $cat 'Critico' 'Sem conexão de rede' 'Nenhum adaptador de rede está conectado (cabo desconectado, Wi-Fi desligado ou modo avião).' 'Verifique o cabo e a porta do switch, ou ligue o Wi-Fi e conecte à rede da empresa.'
        return
    }

    $linhas = @()
    foreach ($a in $ativos) {
        $ehWifi = ($a.PhysicalMediaType -match '802\.11|Wireless') -or ($a.InterfaceDescription -match 'Wi-?Fi|Wireless|WLAN|802\.11') -or ($a.NdisPhysicalMedium -eq 9)
        $ipv4 = @(Get-NetIPAddress -InterfaceIndex $a.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue)
        $valido = $ipv4 | Where-Object { $_.IPAddress -notlike '169.254.*' } | Select-Object -First 1
        $cfg = Get-NetIPConfiguration -InterfaceIndex $a.ifIndex -ErrorAction SilentlyContinue
        $gw = $null; if ($cfg -and $cfg.IPv4DefaultGateway) { $gw = @($cfg.IPv4DefaultGateway)[0].NextHop }
        $velMbps = [math]::Round([double]$a.ReceiveLinkSpeed / 1e6)
        $linhas += [pscustomobject]@{ Adaptador = $a.Name; Tipo = $(if ($ehWifi) { 'Wi-Fi' } else { 'Cabo' }); Velocidade = $a.LinkSpeed; IP = (($ipv4 | ForEach-Object IPAddress) -join ', '); Gateway = $gw; MAC = $a.MacAddress }

        if (-not $valido) {
            $ip = ($ipv4 | ForEach-Object IPAddress) -join ', '; if (-not $ip) { $ip = 'nenhum' }
            Add-Achado $cat 'Critico' "Sem IP válido em '$($a.Name)'" "Endereço IPv4: $ip. Um endereço 169.254.x.x significa que o servidor DHCP não respondeu." 'Renove o IP; se persistir, verifique o cabo/porta do switch, a VLAN e o servidor DHCP.' @(Get-CorrecaoPadrao 'RenovarRede')
        } elseif (-not $gw) {
            Add-Achado $cat 'Atencao' "Sem gateway padrão em '$($a.Name)'" "IP $($valido.IPAddress) sem gateway: o PC não consegue sair da rede local." 'Verifique a configuração de IP (fixo?) ou o DHCP.' @(Get-CorrecaoPadrao 'RenovarRede')
        } elseif (-not (Test-Connection -ComputerName $gw -Count 2 -Quiet -ErrorAction SilentlyContinue)) {
            Add-Achado $cat 'Atencao' "Gateway não responde ($gw)" "O roteador/gateway $gw não respondeu ao ping pelo adaptador '$($a.Name)'." 'Alguns roteadores bloqueiam ping; se a internet também falhar, verifique cabo, switch e roteador.'
        }

        if ($ehWifi) {
            $w = Get-InfoWifi
            if ($w.Erro) {
                Add-Achado $cat 'Info' 'Sinal do Wi-Fi não pôde ser lido' $w.Erro 'Ative "Serviços de localização" para ver o sinal, ou verifique pelo ícone de rede.'
            } elseif ($null -ne $w.Sinal) {
                $banda = ''
                if ($w.Banda) { $banda = $w.Banda } elseif ($w.Canal -match '^\d+$') { if ([int]$w.Canal -le 14) { $banda = '2,4 GHz' } else { $banda = '5 GHz' } }
                $ev = "Rede '$($w.SSID)': sinal $($w.Sinal)%, canal $($w.Canal) ($banda), $($w.Radio), taxa $($w.Taxa) Mbps."
                $rec = 'Aproxime o computador do ponto de acesso, evite paredes/metal no caminho ou use cabo. Verifique se o driver Wi-Fi está atualizado.'
                if ($banda -like '2*') { $rec += ' A rede está em 2,4 GHz (mais interferência e menos velocidade): se houver 5 GHz, prefira.' }
                if ($w.Sinal -lt 40) { Add-Achado $cat 'Critico' "Wi-Fi muito fraco ($($w.Sinal)%)" $ev $rec }
                elseif ($w.Sinal -lt 60) { Add-Achado $cat 'Atencao' "Wi-Fi fraco ($($w.Sinal)%)" $ev $rec }
                else { Add-Achado $cat 'OK' "Wi-Fi com bom sinal ($($w.Sinal)%, $banda)" }
            }
        } else {
            if ($velMbps -gt 0 -and $velMbps -le 10) {
                Add-Achado $cat 'Critico' "Cabo de rede a apenas $velMbps Mbps ($($a.Name))" "Placa: $($a.InterfaceDescription). Link negociado em $($a.LinkSpeed)." 'Troque o cabo (Cat5e/Cat6), confira o conector RJ45 e teste outra porta do switch.'
            } elseif ($velMbps -gt 0 -and $velMbps -le 100) {
                Add-Achado $cat 'Atencao' "Cabo de rede limitado a 100 Mbps ($($a.Name))" "Placa: $($a.InterfaceDescription). Link negociado em $($a.LinkSpeed); o normal hoje é 1 Gbps. Geralmente é cabo com par rompido, conector mal crimpado ou porta antiga de switch." 'Troque o cabo/conector ou teste outra porta do switch. Se o switch for de 100 Mbps, isso é esperado.'
            } else {
                Add-Achado $cat 'OK' "Cabo de rede em $($a.LinkSpeed) ($($a.Name))"
            }
            if ($a.FullDuplex -eq $false) {
                Add-Achado $cat 'Atencao' "Rede em half-duplex ($($a.Name))" 'O link está em half-duplex, o que causa lentidão e colisões.' 'Deixe placa e porta do switch em negociação automática; troque o cabo.'
            }
        }
    }
    Add-Tabela 'Adaptadores de rede conectados' $linhas

    # --- DNS e internet
    $servidores = (@(Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.InterfaceIndex -in $ativos.ifIndex } | ForEach-Object ServerAddresses) | Select-Object -Unique) -join ', '
    $dnsOk = $false
    try { Resolve-DnsName -Name 'www.msftconnecttest.com' -Type A -DnsOnly -QuickTimeout -ErrorAction Stop | Out-Null; $dnsOk = $true } catch { }
    $ipOk = (Test-PortaTcp '1.1.1.1' 443) -or (Test-PortaTcp '8.8.8.8' 443)
    $httpOk = $false
    if ($dnsOk) {
        try { $w = Invoke-WebRequest 'http://www.msftconnecttest.com/connecttest.txt' -UseBasicParsing -TimeoutSec 8 -ErrorAction Stop; $httpOk = ($w.Content -match 'Microsoft Connect Test') } catch { }
    }
    if (-not $dnsOk -and $ipOk) {
        Add-Achado $cat 'Critico' 'DNS falhando' "Servidores DNS configurados: $servidores. A internet responde por IP, mas os nomes não são resolvidos." 'Limpe o cache DNS e renove o IP; se persistir, verifique o servidor DNS (ou use o DNS da empresa/roteador).' @(Get-CorrecaoPadrao 'RenovarRede')
    } elseif (-not $dnsOk -and -not $ipOk) {
        Add-Achado $cat 'Critico' 'Sem acesso à internet' "Nem a resolução de nomes nem a conexão direta por IP funcionaram. DNS: $servidores." 'Verifique roteador/link de internet e firewall. Em redes que só saem por proxy, este teste pode falhar mesmo com o navegador funcionando.' @(Get-CorrecaoPadrao 'RenovarRede')
    } elseif (-not $httpOk) {
        Add-Achado $cat 'Atencao' 'Nomes resolvem, mas a internet não responde' 'O DNS funciona, porém o teste de conectividade da Microsoft não respondeu (proxy, firewall ou portal de login do Wi-Fi).' 'Abra o navegador para ver se há portal de login; verifique proxy e firewall.'
    } else {
        Add-Achado $cat 'OK' "Internet e DNS funcionando (DNS: $servidores)"
    }

    # --- Domínio
    $r = $Script:Dados.Resumo
    if ($r.NoDominio) {
        $dom = (Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).Domain
        $domOk = $false
        try { Resolve-DnsName -Name $dom -DnsOnly -QuickTimeout -ErrorAction Stop | Out-Null; $domOk = $true } catch { }
        if (-not $domOk) {
            Add-Achado $cat 'Critico' "DNS não encontra o domínio $dom" "Servidores DNS em uso: $servidores. Em PCs de domínio o DNS deve apontar para os controladores de domínio." 'Configure o DNS da empresa (normalmente via DHCP). Fora da empresa, conecte a VPN.'
        } elseif ($Script:EhAdmin) {
            $canal = $null
            try { $canal = Test-ComputerSecureChannel -ErrorAction Stop } catch { }
            if ($canal -eq $false) {
                Add-Achado $cat 'Critico' 'Relação de confiança com o domínio quebrada' "O computador não conseguiu validar o canal seguro com o domínio $dom." 'Execute "Test-ComputerSecureChannel -Repair -Credential DOMINIO\admin" ou reingresse a máquina no domínio.'
            } elseif ($canal) {
                Add-Achado $cat 'OK' "Conectado ao domínio $dom"
            }
        }
    }
}

function Invoke-RenovarRede {
    Write-Host '   ipconfig /renew ...' -ForegroundColor Gray
    ipconfig /renew | Out-Null
    ipconfig /flushdns | Out-Null
    ipconfig /registerdns | Out-Null
    Clear-DnsClientCache -ErrorAction SilentlyContinue
}

function Invoke-ResetRede {
    if (-not (Confirm-SN 'A redefinição da pilha de rede exige reiniciar e pode remover configurações de VPN antigas. Continuar?' 'N')) { throw 'Cancelado pelo técnico.' }
    netsh winsock reset | Out-Null
    netsh int ip reset | Out-Null
    ipconfig /flushdns | Out-Null
    Write-Host '   Pilha de rede redefinida. REINICIE o computador para concluir.' -ForegroundColor Yellow
}

#endregion

#region ===================== Diagnóstico: Impressão =====================

function Test-Impressao {
    Write-Etapa 'Impressão: serviço de spooler, filas travadas e impressoras com erro'
    $cat = 'Impressão'
    $sp = Get-Service -Name Spooler -ErrorAction SilentlyContinue
    if (-not $sp) { return }
    if ($sp.Status -ne 'Running') {
        Add-Achado $cat 'Critico' 'Serviço de impressão (Spooler) parado' "Status: $($sp.Status); inicialização: $($sp.StartType)." 'Inicie o serviço. Se ele parar de novo sozinho, um driver de impressora está travando o spooler.' @(Get-CorrecaoPadrao 'IniciarSpooler')
        return
    }
    $impressoras = @(Get-Printer -ErrorAction SilentlyContinue)
    $travadas = 0
    $linhas = @()
    foreach ($p in $impressoras) {
        $jobs = @(Get-PrintJob -PrinterName $p.Name -ErrorAction SilentlyContinue)
        $presos = @($jobs | Where-Object { $_.JobStatus -match 'Error|Blocked|Offline|PaperOut|UserIntervention|Deleting' -or ($_.SubmittedTime -and $_.SubmittedTime -lt (Get-Date).AddMinutes(-20)) })
        $linhas += [pscustomobject]@{ Impressora = $p.Name; Status = [string]$p.PrinterStatus; Porta = $p.PortName; Driver = $p.DriverName; 'Na fila' = $jobs.Count }
        if ($presos) {
            $travadas++
            $velho = ($presos | Sort-Object SubmittedTime | Select-Object -First 1)
            $min = [math]::Round(((Get-Date) - $velho.SubmittedTime).TotalMinutes)
            Add-Achado $cat 'Atencao' "Fila de impressão travada: $($p.Name)" "$($presos.Count) trabalho(s) parado(s); o mais antigo há $min min ('$($velho.DocumentName)', status: $($velho.JobStatus))." 'Destrave a fila (os documentos precisarão ser reimpressos) e verifique papel, conexão e erros no painel da impressora.' @(Get-CorrecaoPadrao 'LimparFila')
        }
    }
    if ($impressoras) { Add-Tabela 'Impressoras' $linhas }
    if ($impressoras -and $travadas -eq 0) { Add-Achado $cat 'OK' "Nenhuma fila de impressão travada ($($impressoras.Count) impressora(s))" }

    $padrao = Get-CimInstance Win32_Printer -Filter 'Default=TRUE' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($padrao -and ($padrao.WorkOffline -or $padrao.PrinterStatus -in 7)) {
        Add-Achado $cat 'Atencao' "Impressora padrão offline: $($padrao.Name)" 'A impressora padrão está marcada como offline.' 'Verifique se ela está ligada e na rede; desmarque "Usar impressora offline" na fila.'
    }

    $falhas = @(Get-Eventos -Log 'Application' -Ids 1000 -Provider 'Application Error' -Dias 14 -Max 200 | Where-Object { $_.Properties.Count -gt 0 -and ([string]$_.Properties[0].Value) -ieq 'spoolsv.exe' })
    if ($falhas) {
        $mod = ($falhas | Where-Object { $_.Properties.Count -gt 3 } | Group-Object { [string]$_.Properties[3].Value } | Sort-Object Count -Descending | Select-Object -First 1).Name
        Add-Achado $cat 'Atencao' "Spooler de impressão travou $($falhas.Count) vez(es) em 14 dias" "Módulo envolvido: $mod." 'Normalmente é driver de impressora com defeito: atualize ou remova o driver citado.'
    }
}

function Clear-FilaImpressao {
    Write-Host '   Parando o spooler e removendo trabalhos presos...' -ForegroundColor Gray
    Stop-Service -Name Spooler -Force -ErrorAction Stop
    $pasta = Join-Path $env:SystemRoot 'System32\spool\PRINTERS'
    Get-ChildItem -LiteralPath $pasta -File -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    Start-Service -Name Spooler -ErrorAction Stop
}

#endregion

#region ===================== Diagnóstico: Sistema =====================

function Test-Sistema {
    Write-Etapa 'Sistema: atualizações, reinício pendente, antivírus, drivers e bateria'
    $cat = 'Sistema'
    if (-not $Script:EhAdmin) { Add-Achado 'Ferramenta' 'Info' 'Executado sem administrador' 'Algumas verificações (tempo de boot, saúde do disco, temperatura, canal do domínio) foram limitadas.' 'Execute como administrador para o diagnóstico completo.' }

    if (Test-ReinicioPendente) {
        Add-Achado $cat 'Atencao' 'Reinício pendente' 'O Windows tem atualizações ou componentes aguardando reinício.' 'Reinicie o computador para concluir as instalações.'
    }

    $hf = @(Get-HotFix -ErrorAction SilentlyContinue | Where-Object { $_.InstalledOn } | Sort-Object InstalledOn -Descending)
    if ($hf) {
        $dias = [math]::Round(((Get-Date) - $hf[0].InstalledOn).TotalDays)
        if ($dias -ge 60) {
            $sev = 'Atencao'; if ($dias -ge 90) { $sev = 'Critico' }
            Add-Achado $cat $sev "Windows sem atualizações há $dias dias" "Última atualização instalada: $($hf[0].HotFixID) em $($hf[0].InstalledOn.ToString('dd/MM/yyyy'))." 'Instale as atualizações; se o Windows Update estiver falhando, use o reparo do Windows Update.' @(Get-CorrecaoPadrao 'AtualizarWindows'), @(Get-CorrecaoPadrao 'ReparoWU')
        } else {
            Add-Achado $cat 'OK' "Última atualização do Windows há $dias dia(s)"
        }
    }
    $wu = Get-Service wuauserv -ErrorAction SilentlyContinue
    if ($wu -and $wu.StartType -eq 'Disabled') {
        Add-Achado $cat 'Atencao' 'Serviço Windows Update desativado' 'O serviço wuauserv está desativado.' 'Se não for intencional (WSUS/política), reative o serviço.'
    }

    # Antivírus
    $avs = @(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction SilentlyContinue)
    if ($avs) {
        $ativos = @(); $desatualizados = @()
        foreach ($av in $avs) {
            $hex = '{0:X6}' -f [int]$av.productState
            if ($hex.Substring(2, 2) -in '10', '11') { $ativos += $av.displayName }
            if ($hex.Substring(4, 2) -ne '00') { $desatualizados += $av.displayName }
        }
        if (-not $ativos) {
            Add-Achado $cat 'Critico' 'Nenhum antivírus ativo' "Instalados: $(($avs | ForEach-Object displayName) -join ', '), mas nenhum com proteção em tempo real ligada." 'Ative a proteção em tempo real ou verifique o console do antivírus corporativo.'
        } else {
            Add-Achado $cat 'OK' "Antivírus ativo: $($ativos -join ', ')"
            $velhos = @($desatualizados | Where-Object { $_ -in $ativos })
            if ($velhos) { Add-Achado $cat 'Atencao' "Antivírus desatualizado: $($velhos -join ', ')" 'As definições de vírus não estão em dia.' 'Atualize as definições e verifique a comunicação com o servidor do antivírus.' }
        }
    } else {
        $mp = Get-MpComputerStatus -ErrorAction SilentlyContinue
        if ($mp) {
            if (-not $mp.RealTimeProtectionEnabled -and $mp.AMRunningMode -eq 'Normal') { Add-Achado $cat 'Critico' 'Proteção em tempo real do Defender desligada' 'O Microsoft Defender está sem proteção em tempo real.' 'Ative a proteção em tempo real.' }
            elseif ($mp.AntivirusSignatureAge -gt 7) { Add-Achado $cat 'Atencao' "Defender sem atualizar há $($mp.AntivirusSignatureAge) dias" 'Definições de vírus antigas.' 'Atualize as definições.' @(Get-CorrecaoPadrao 'AtualizarWindows') }
        }
    }

    # Dispositivos com problema
    $probs = @(Get-CimInstance Win32_PnPEntity -Filter 'ConfigManagerErrorCode <> 0' -ErrorAction SilentlyContinue |
            Where-Object { $_.ConfigManagerErrorCode -notin 22, 24, 45 -and $_.PNPClass -notin 'Camera', 'Image', 'AudioEndpoint' })
    if ($probs) {
        $lista = ($probs | Select-Object -First 6 | ForEach-Object {
                $n = $_.Name; if (-not $n) { $n = $_.DeviceID }
                $d = $Script:CodigosDispositivo[[int]$_.ConfigManagerErrorCode]; if (-not $d) { $d = "código $($_.ConfigManagerErrorCode)" }
                "$n ($d)" }) -join '; '
        Add-Achado $cat 'Atencao' "$($probs.Count) dispositivo(s) com problema ou sem driver" $lista 'Instale os drivers pelo Windows Update ou pelo site do fabricante do computador.' @(Get-CorrecaoPadrao 'AtualizarDrivers')
        Add-Tabela 'Dispositivos com problema' ($probs | ForEach-Object { [pscustomobject]@{ Dispositivo = $_.Name; Classe = $_.PNPClass; 'Código' = $_.ConfigManagerErrorCode; Significado = $Script:CodigosDispositivo[[int]$_.ConfigManagerErrorCode] } })
    } else {
        Add-Achado $cat 'OK' 'Todos os dispositivos com driver funcionando'
    }

    # Bateria
    if ($Script:Dados.Resumo.Notebook) {
        try {
            $full = (Get-CimInstance -Namespace root/wmi -ClassName BatteryFullChargedCapacity -ErrorAction Stop | Measure-Object FullChargedCapacity -Sum).Sum
            $proj = (Get-CimInstance -Namespace root/wmi -ClassName BatteryStaticData -ErrorAction Stop | Measure-Object DesignedCapacity -Sum).Sum
            if ($full -gt 0 -and $proj -gt 0) {
                $saude = [math]::Round($full / $proj * 100)
                if ($saude -lt 50) { Add-Achado $cat 'Critico' "Bateria muito desgastada ($saude% da capacidade original)" "Carga máxima atual: $full mWh de $proj mWh de fábrica." 'Programe a troca da bateria. Baterias muito gastas também causam desligamentos inesperados.' }
                elseif ($saude -lt 70) { Add-Achado $cat 'Atencao' "Bateria desgastada ($saude% da capacidade original)" "Carga máxima atual: $full mWh de $proj mWh de fábrica." 'Planeje a troca da bateria.' }
                else { Add-Achado $cat 'OK' "Bateria com $saude% da capacidade original" }
            }
        } catch { }
    }
}

#endregion

#region ===================== Diagnóstico: orquestração =====================

function Invoke-Diagnostico {
    param([switch]$Silencioso)
    $Script:Achados.Clear()
    $Script:Correcoes.Clear()
    $Script:Dados.Tabelas = [ordered]@{}
    $Script:Dados.AtrasoBoot = @{}
    Write-Titulo 'Análise de problemas'
    Get-Resumo | Out-Null
    foreach ($t in 'Test-Lentidao', 'Test-Travamentos', 'Test-CameraMicrofone', 'Test-Rede', 'Test-Impressao', 'Test-Sistema') {
        try { & $t } catch {
            Write-Log ("Erro em {0}: {1}" -f $t, $_.Exception.Message)
            Add-Achado 'Ferramenta' 'Info' "Falha ao executar a verificação $t" $_.Exception.Message
        }
    }
    $Script:DiagnosticoFeito = $true
    $Script:Dados.DataDiagnostico = Get-Date
    if (-not $Silencioso) { Show-ResultadoDiagnostico }
}

function Get-EstadoGui {
    # Visão "só dados" do estado do motor, para a janela gráfica desenhar (sem scriptblocks).
    $achados = @(Get-AchadosOrdenados | ForEach-Object {
            $nomes = @($_.Correcoes | ForEach-Object { $c = $Script:Correcoes[$_]; if ($c) { if ($c.Aplicada) { "$($c.Nome) (aplicada)" } else { $c.Nome } } })
            [pscustomobject]@{ Categoria = $_.Categoria; Severidade = $_.Severidade; Titulo = $_.Titulo; Evidencia = $_.Evidencia; Recomendacao = $_.Recomendacao; Correcoes = ($nomes -join '; ') } })
    [pscustomobject]@{
        Resumo           = $Script:Dados.Resumo
        Achados          = $achados
        Pendentes        = @($Script:Correcoes.Values | Where-Object { -not $_.Aplicada }).Count
        DiagnosticoFeito = $Script:DiagnosticoFeito
        EhAdmin          = $Script:EhAdmin
        PastaRelatorios  = $Script:PastaRelLocal
    }
}

function Get-StatusAgendamento {
    $t = Get-ScheduledTask -TaskPath '\speedPC\' -TaskName $Script:NomeTarefa -ErrorAction SilentlyContinue
    if (-not $t) { return [pscustomobject]@{ Existe = $false; Texto = 'Nenhum agendamento configurado neste computador.'; Pasta = ''; Dia = 'Wednesday'; Hora = '12:00' } }
    $info = $t | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue
    $arg = [string]$t.Actions[0].Arguments
    $pasta = ''; if ($arg -match '-PastaRelatorios "([^"]+)"') { $pasta = $Matches[1] }
    $g = $t.Triggers[0]
    $dia = 'Wednesday'
    try { $dia = ([string]$g.DaysOfWeek); $nomes = @{ 1 = 'Sunday'; 2 = 'Monday'; 4 = 'Tuesday'; 8 = 'Wednesday'; 16 = 'Thursday'; 32 = 'Friday'; 64 = 'Saturday' }; if ($nomes.ContainsKey([int]$g.DaysOfWeek)) { $dia = $nomes[[int]$g.DaysOfWeek] } } catch { }
    $hora = '12:00'; try { $hora = ([datetime]$g.StartBoundary).ToString('HH:mm') } catch { }
    $txt = "Agendado: toda semana às $hora."
    if ($info -and $info.LastRunTime -and $info.LastRunTime.Year -gt 2000) { $txt += " Última execução: $($info.LastRunTime.ToString('dd/MM/yyyy HH:mm'))." }
    if ($info -and $info.NextRunTime) { $txt += " Próxima: $($info.NextRunTime.ToString('dd/MM/yyyy HH:mm'))." }
    if ($pasta) { $txt += " Relatórios em $pasta." }
    return [pscustomobject]@{ Existe = $true; Texto = $txt; Pasta = $pasta; Dia = $dia; Hora = $hora }
}

function Test-PastaRede {
    param([string]$Pasta)
    if (-not $Pasta) { return 'Sem pasta de rede: os relatórios ficam só neste computador.' }
    try {
        $teste = Join-Path $Pasta ".speedpc_teste_$env:COMPUTERNAME"
        Set-Content -LiteralPath $teste -Value 'ok' -ErrorAction Stop
        Remove-Item -LiteralPath $teste -Force -ErrorAction SilentlyContinue
        return 'Gravação na pasta testada com o seu usuário: OK.'
    } catch { return "Aviso: o seu usuário não conseguiu gravar em $Pasta ($($_.Exception.Message))." }
}

function Get-AchadosOrdenados {
    param([switch]$SomenteProblemas)
    $lista = $Script:Achados
    if ($SomenteProblemas) { $lista = $lista | Where-Object { $_.Severidade -ne 'OK' } }
    $lista | Sort-Object @{ e = { $i = [array]::IndexOf($Script:OrdemCategorias, $_.Categoria); if ($i -lt 0) { 99 } else { $i } } }, Ordem
}

function Show-ResultadoDiagnostico {
    $prob = @(Get-AchadosOrdenados -SomenteProblemas)
    $ok = @($Script:Achados | Where-Object { $_.Severidade -eq 'OK' })
    $cont = @{}
    foreach ($s in 'Critico', 'Atencao', 'Info') { $cont[$s] = @($prob | Where-Object { $_.Severidade -eq $s }).Count }
    Write-Host ''
    $catAtual = ''
    foreach ($a in $prob) {
        if ($a.Categoria -ne $catAtual) { $catAtual = $a.Categoria; Write-Host ''; Write-Host "  --- $catAtual ---" -ForegroundColor White }
        $s = $Script:Sev[$a.Severidade]
        Write-Host ('  [{0,-7}] ' -f $s.Rotulo) -ForegroundColor $s.Cor -NoNewline
        Write-Host $a.Titulo -ForegroundColor White
        if ($a.Evidencia) { Write-Host "            Evidência:   $($a.Evidencia)" -ForegroundColor Gray }
        if ($a.Recomendacao) { Write-Host "            O que fazer: $($a.Recomendacao)" -ForegroundColor DarkGray }
        if ($a.Correcoes) { Write-Host "            Correção automática disponível." -ForegroundColor DarkCyan }
    }
    Write-Host ''
    Write-Host ('  ' + ('-' * 76)) -ForegroundColor DarkGray
    Write-Host '  Resultado: ' -NoNewline
    Write-Host "$($cont.Critico) crítico(s)" -ForegroundColor Red -NoNewline; Write-Host '  |  ' -NoNewline
    Write-Host "$($cont.Atencao) atenção" -ForegroundColor Yellow -NoNewline; Write-Host '  |  ' -NoNewline
    Write-Host "$($cont.Info) informativo(s)" -ForegroundColor Cyan -NoNewline; Write-Host '  |  ' -NoNewline
    Write-Host "$($ok.Count) verificação(ões) OK" -ForegroundColor Green
    if (-not $prob) { Write-Host '  Nenhum problema encontrado.' -ForegroundColor Green }
}

#endregion

#region ===================== Limpeza segura =====================

function Remove-ConteudoPasta {
    # Apaga arquivos dentro da pasta (nunca a pasta em si), sem seguir junções/links simbólicos.
    param([string]$Pasta, [int]$Dias = 0, [string]$Filtro = '*', [string[]]$Excluir = @())
    $bytes = [int64]0
    if (-not $Pasta -or -not (Test-Path -LiteralPath $Pasta -PathType Container)) { return $bytes }
    $raiz = New-Object IO.DirectoryInfo $Pasta
    if ($raiz.Attributes -band [IO.FileAttributes]::ReparsePoint) { return $bytes }
    $limite = (Get-Date).AddDays(-$Dias)
    $pilha = New-Object System.Collections.Stack
    $pilha.Push($raiz)
    $subpastas = New-Object System.Collections.Generic.List[string]
    while ($pilha.Count -gt 0) {
        $d = $pilha.Pop()
        $itens = @()
        try { $itens = @($d.GetFileSystemInfos()) } catch { continue }
        foreach ($i in $itens) {
            if ($i.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
            $pular = $false
            foreach ($x in $Excluir) { if ($i.FullName -like "*$x*") { $pular = $true; break } }
            if ($pular) { continue }
            if ($i -is [IO.DirectoryInfo]) { $pilha.Push($i); $subpastas.Add($i.FullName); continue }
            if ($i.Name -notlike $Filtro) { continue }
            if ($i.LastWriteTime -gt $limite) { continue }
            try {
                $tam = $i.Length
                if ($i.Attributes -band [IO.FileAttributes]::ReadOnly) { $i.Attributes = 'Normal' }
                $i.Delete()
                $bytes += $tam
            } catch { }
        }
    }
    foreach ($s in ($subpastas | Sort-Object Length -Descending)) {
        try { [IO.Directory]::Delete($s, $false) } catch { }   # só remove se estiver vazia
    }
    return $bytes
}

function Get-AlvosUsuario {
    $alvos = @(
        @{ Nome = 'Temporários dos usuários'; Rel = 'AppData\Local\Temp'; Dias = 2 }
        @{ Nome = 'Cache de internet (IE/Edge legado)'; Rel = 'AppData\Local\Microsoft\Windows\INetCache'; Excluir = @('Content.Outlook', 'Content.MSO') }
        @{ Nome = 'Relatórios de erro'; Rel = 'AppData\Local\Microsoft\Windows\WER' }
        @{ Nome = 'Despejos de falha de programas'; Rel = 'AppData\Local\CrashDumps' }
        @{ Nome = 'Cache de miniaturas'; Rel = 'AppData\Local\Microsoft\Windows\Explorer'; Filtro = 'thumbcache_*.db' }
        @{ Nome = 'Cache de sombreadores (DirectX)'; Rel = 'AppData\Local\D3DSCache' }
        @{ Nome = 'Cache da Área de Trabalho Remota'; Rel = 'AppData\Local\Microsoft\Terminal Server Client\Cache' }
        @{ Nome = 'Cache do Firefox'; Rel = 'AppData\Local\Mozilla\Firefox\Profiles\*\cache2' }
    )
    $navegadores = @(
        @{ Nome = 'Cache do Chrome'; Base = 'AppData\Local\Google\Chrome\User Data' }
        @{ Nome = 'Cache do Edge'; Base = 'AppData\Local\Microsoft\Edge\User Data' }
        @{ Nome = 'Cache do Brave'; Base = 'AppData\Local\BraveSoftware\Brave-Browser\User Data' }
    )
    foreach ($n in $navegadores) {
        foreach ($sub in 'Cache', 'Code Cache', 'GPUCache') { $alvos += @{ Nome = $n.Nome; Rel = "$($n.Base)\*\$sub" } }
        $alvos += @{ Nome = $n.Nome; Rel = "$($n.Base)\ShaderCache" }
        $alvos += @{ Nome = $n.Nome; Rel = "$($n.Base)\GrShaderCache" }
    }
    foreach ($sub in 'Cache', 'Code Cache', 'GPUCache') { $alvos += @{ Nome = 'Cache do Teams (clássico)'; Rel = "AppData\Roaming\Microsoft\Teams\$sub" } }
    return $alvos
}

function Get-TamanhoPasta {
    param([string]$Pasta)
    if (-not (Test-Path -LiteralPath $Pasta)) { return 0 }
    return [int64](Get-ChildItem -LiteralPath $Pasta -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
}

function Invoke-LimpezaSegura {
    param([switch]$Silencioso, [switch]$IncluirLixeira, [switch]$Profunda)
    Write-Titulo 'Limpeza segura'
    Write-Host '  NÃO são tocados: senhas, cookies, histórico, favoritos, Downloads, Documentos, Área de Trabalho, arquivos do Outlook e despejos de tela azul.' -ForegroundColor DarkGray
    if (-not $Script:EhAdmin) { Write-Host '  Sem administrador: apenas o perfil do usuário atual será limpo.' -ForegroundColor Yellow }
    if (-not $Silencioso -and -not $Script:ModoAuto) {
        $IncluirLixeira = Confirm-SN 'Esvaziar também a Lixeira de todos os usuários? (contém arquivos que os usuários apagaram)' 'N'
        if ($Script:EhAdmin) { $Profunda = Confirm-SN 'Incluir limpeza de componentes do Windows (WinSxS)? Pode levar mais de 10 minutos.' 'N' }
    }
    $disco = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'" -ErrorAction SilentlyContinue
    $livreAntes = [int64]$disco.FreeSpace
    $porItem = [ordered]@{}
    $soma = { param($nome, $b) if (-not $porItem.Contains($nome)) { $porItem[$nome] = [int64]0 }; $porItem[$nome] += [int64]$b }

    $perfis = @(Get-PerfisUsuarios)
    $alvos = Get-AlvosUsuario
    Write-Host "  Limpando $($perfis.Count) perfil(is) de usuário..." -ForegroundColor Gray
    foreach ($p in $perfis) {
        foreach ($a in $alvos) {
            $caminho = Join-Path $p.Caminho $a.Rel
            $pastas = @(Get-Item -Path $caminho -Force -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer })
            foreach ($pasta in $pastas) {
                $dias = 0; if ($a.Dias) { $dias = $a.Dias }
                $filtro = '*'; if ($a.Filtro) { $filtro = $a.Filtro }
                $excl = @(); if ($a.Excluir) { $excl = $a.Excluir }
                & $soma $a.Nome (Remove-ConteudoPasta -Pasta $pasta.FullName -Dias $dias -Filtro $filtro -Excluir $excl)
            }
        }
    }

    if ($Script:EhAdmin) {
        Write-Host '  Limpando pastas do sistema...' -ForegroundColor Gray
        & $soma 'Temporários do Windows' (Remove-ConteudoPasta -Pasta "$env:SystemRoot\Temp" -Dias 2)
        foreach ($w in 'ReportArchive', 'ReportQueue', 'Temp') {
            & $soma 'Relatórios de erro do sistema' (Remove-ConteudoPasta -Pasta "$env:ProgramData\Microsoft\Windows\WER\$w")
        }
        if (Test-ReinicioPendente) {
            Write-Host '  Cache do Windows Update preservado: há um reinício pendente.' -ForegroundColor DarkGray
        } else {
            $rodando = @(Get-Service wuauserv, bits -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Running' })
            try {
                $rodando | Stop-Service -Force -ErrorAction Stop
                & $soma 'Downloads antigos do Windows Update' (Remove-ConteudoPasta -Pasta "$env:SystemRoot\SoftwareDistribution\Download")
            } catch { Write-Host "  Não foi possível parar o Windows Update: $($_.Exception.Message)" -ForegroundColor Yellow }
            finally { $rodando | Start-Service -ErrorAction SilentlyContinue }
        }
        if (Get-Command Delete-DeliveryOptimizationCache -ErrorAction SilentlyContinue) {
            $doPasta = "$env:SystemRoot\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"
            $antes = Get-TamanhoPasta $doPasta
            try { Delete-DeliveryOptimizationCache -Force -ErrorAction Stop; & $soma 'Cache de Otimização de Entrega' ([math]::Max(0, $antes - (Get-TamanhoPasta $doPasta))) } catch { }
        }
        if ($IncluirLixeira) {
            foreach ($v in (Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction SilentlyContinue)) {
                & $soma 'Lixeira (todos os usuários)' (Remove-ConteudoPasta -Pasta (Join-Path "$($v.DeviceID)\" '$Recycle.Bin') -Excluir @('desktop.ini'))
            }
        }
        if ($Profunda) {
            Write-Host '  Limpando componentes do Windows (DISM /StartComponentCleanup)...' -ForegroundColor Gray
            $sys = Get-PastaSistema
            Start-Ferramenta "$sys\dism.exe" @('/Online', '/Cleanup-Image', '/StartComponentCleanup') | Out-Null
        }
    }

    $livreDepois = [int64](Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'" -ErrorAction SilentlyContinue).FreeSpace
    $total = [int64]0; foreach ($v in $porItem.Values) { $total += $v }
    Write-Host ''
    foreach ($k in $porItem.Keys) { if ($porItem[$k] -gt 0) { Write-Host ('  {0,-45} {1,12}' -f $k, (Format-Bytes $porItem[$k])) } }
    Write-Host ('  {0,-45} {1,12}' -f 'TOTAL (arquivos removidos)', (Format-Bytes $total)) -ForegroundColor Green
    if ($livreDepois -gt $livreAntes) { Write-Host ("  Espaço livre em {0}: {1} -> {2}" -f $env:SystemDrive, (Format-Bytes $livreAntes), (Format-Bytes $livreDepois)) -ForegroundColor Green }
    Add-Tabela 'Limpeza segura' (@($porItem.Keys | Where-Object { $porItem[$_] -gt 0 } | ForEach-Object { [pscustomobject]@{ Item = $_; Liberado = (Format-Bytes $porItem[$_]) } }))
    $extra = ''; if ($Profunda) { $extra = ' + limpeza de componentes' }; if ($IncluirLixeira) { $extra += ' + Lixeira' }
    Add-Acao 'Limpeza segura' ("{0} removidos em {1} perfil(is){2}" -f (Format-Bytes $total), $perfis.Count, $extra)
}

#endregion

#region ===================== Otimização =====================

function Invoke-Otimizacao {
    Write-Titulo 'Otimização para o trabalho'
    if (-not $Script:EhAdmin) { Write-Host '  Requer administrador.' -ForegroundColor Yellow; return }
    $r = $Script:Dados.Resumo; if (-not $r) { $r = Get-Resumo }
    $plano = Get-PlanoEnergiaAtivo
    $opcoes = @()

    if (($r.Notebook -and $plano.Guid -ne $Script:GuidPlanos.Equilibrado) -or (-not $r.Notebook -and $plano.Guid -ne $Script:GuidPlanos.Alto)) { $opcoes += Get-CorrecaoOpcao 'O:Plano' }
    if ($r.InicioRapido) { $opcoes += Get-CorrecaoOpcao 'O:InicioRapido' }
    $opcoes += Get-CorrecaoOpcao 'O:Inicializacao'
    $opcoes += Get-CorrecaoOpcao 'O:Sugestoes'
    $opcoes += Get-CorrecaoOpcao 'O:StorageSense'
    $opcoes += Get-CorrecaoOpcao 'O:Unidades'

    $opcoes = @($opcoes | ForEach-Object { $Script:Correcoes[$_] } | Where-Object { $_ -and -not $_.Aplicada })
    $sel = @(Select-Opcoes -Titulo 'Otimizações disponíveis' -Itens @($opcoes | ForEach-Object { [pscustomobject]@{ Nome = $_.Nome; Detalhe = $_.Descricao } }) -Aviso 'Tudo o que altera configurações pode ser revertido em "Desfazer".')
    foreach ($i in $sel) { Invoke-Correcao $opcoes[$i] }
}

function Get-CorrecaoOpcao {
    # Opções fixas das abas Otimização, Reparo e Atualizações (ids usados pela janela gráfica)
    param([string]$Id)
    switch ($Id) {
        'O:Plano' {
            $r = $Script:Dados.Resumo; if (-not $r) { $r = Get-Resumo }
            $temAlto = ((powercfg /list) | Out-String) -match $Script:GuidPlanos.Alto
            if ($r.Notebook -or -not $temAlto) {
                return New-Correcao 'Otim:PlanoEq' 'Plano de energia Equilibrado' 'Melhor equilíbrio entre desempenho e consumo (recomendado para notebooks).' 'Reversivel' { Set-PlanoEnergiaSeguro -Guid $Script:GuidPlanos.Equilibrado -Nome 'Equilibrado' }
            }
            return New-Correcao 'Otim:PlanoAlto' 'Plano de energia Alto desempenho' 'Em desktops, mantém o processador sempre pronto.' 'Reversivel' { Set-PlanoEnergiaSeguro -Guid $Script:GuidPlanos.Alto -Nome 'Alto desempenho' }
        }
        'O:InicioRapido' { return Get-CorrecaoPadrao 'InicioRapido' }
        'O:Inicializacao' { return Get-CorrecaoPadrao 'Inicializacao' }
        'O:Sugestoes' { return New-Correcao 'Otim:Sugestoes' 'Desativar sugestões, anúncios e apps instalados automaticamente' 'Desliga dicas, aplicativos patrocinados e sugestões do menu Iniciar para os usuários logados e por política.' 'Reversivel' { Disable-SugestoesWindows } }
        'O:StorageSense' { return New-Correcao 'Otim:StorageSense' 'Ativar Sensor de Armazenamento (temporários)' 'O Windows passa a apagar sozinho arquivos temporários. Lixeira e Downloads ficam de fora.' 'Reversivel' { Enable-SensorArmazenamento } }
        'O:Unidades' {
            return New-Correcao 'Otim:Unidades' 'Otimizar unidades (TRIM no SSD / desfragmentar HD)' 'Usa o otimizador do próprio Windows, que escolhe a operação certa para cada tipo de disco.' 'SemEfeitoPermanente' {
                Get-Volume -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter -and $_.DriveType -eq 'Fixed' -and $_.FileSystem -match 'NTFS|ReFS' } | ForEach-Object {
                    Write-Host "   Otimizando $($_.DriveLetter): ..." -ForegroundColor Gray
                    Optimize-Volume -DriveLetter $_.DriveLetter -ErrorAction SilentlyContinue }
            }
        }
        'R:ReparoWindows' { return Get-CorrecaoPadrao 'ReparoWindows' }
        'R:VerificarDisco' { return Get-CorrecaoPadrao 'VerificarDisco' }
        'R:ReparoWU' { return Get-CorrecaoPadrao 'ReparoWU' }
        'R:RenovarRede' { return Get-CorrecaoPadrao 'RenovarRede' }
        'R:ResetRede' { return Get-CorrecaoPadrao 'ResetRede' }
        'R:LimparFila' { return Get-CorrecaoPadrao 'LimparFila' }
        'R:TesteMemoria' { return Get-CorrecaoPadrao 'TesteMemoria' }
        'A:Windows' { return Get-CorrecaoPadrao 'AtualizarWindows' }
        'A:Drivers' { return Get-CorrecaoPadrao 'AtualizarDrivers' }
        'O:Ultimate' { return New-Correcao 'Otim:Ultimate' 'Plano de energia Desempenho Máximo' 'Ativa o plano oculto de estações de trabalho.' 'Reversivel' { Enable-DesempenhoMaximo } }
        'A:PerfilEmpresa' { return New-Correcao 'AtualPerfil' 'Perfil de atualizações para empresas' 'Adia versões novas por 365 dias e qualidade por 4 dias.' 'Reversivel' { Set-PerfilAtualizacaoEmpresa } }
        'A:Winget' { return New-Correcao 'Winget' 'Atualizar programas (winget)' 'Atualiza os programas instalados pelo gerenciador de pacotes do Windows.' 'NaoReversivel' { Invoke-AtualizacaoProgramas } -Interativa }
    }
    return $null
}

function Enable-DesempenhoMaximo {
    if ($Script:Dados.Resumo -and $Script:Dados.Resumo.Notebook) { throw 'Plano não recomendado para notebooks (aumenta o consumo e o calor).' }
    $lista = (powercfg /list) | Out-String
    $m = [regex]::Match($lista, '([0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12})\s+\((Desempenho M.ximo|Ultimate Performance)\)')
    if ($m.Success) { $guid = $m.Groups[1].Value }
    else {
        $s = (powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61) | Out-String
        $guid = [regex]::Match($s, '[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}').Value
        if (-not $guid) { throw 'Esta edição do Windows não oferece o plano Desempenho Máximo.' }
    }
    Set-PlanoEnergiaSeguro -Guid $guid -Nome 'Desempenho Máximo'
}

function Set-PerfilAtualizacaoEmpresa {
    $k = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $atual = Get-ItemProperty -LiteralPath $k -ErrorAction SilentlyContinue
    if ($atual.WUServer) { throw "Este computador recebe atualizações de um servidor WSUS ($($atual.WUServer)). Ajuste a política no servidor." }
    Set-RegistroSeguro -Caminho $k -Nome 'DeferFeatureUpdates' -Valor 1 -Tipo DWord -Descricao 'Windows Update: adiar versões novas'
    Set-RegistroSeguro -Caminho $k -Nome 'DeferFeatureUpdatesPeriodInDays' -Valor 365 -Tipo DWord -Descricao 'Windows Update: versões novas após 365 dias'
    Set-RegistroSeguro -Caminho $k -Nome 'DeferQualityUpdates' -Valor 1 -Tipo DWord -Descricao 'Windows Update: adiar atualizações de qualidade'
    Set-RegistroSeguro -Caminho $k -Nome 'DeferQualityUpdatesPeriodInDays' -Valor 4 -Tipo DWord -Descricao 'Windows Update: qualidade após 4 dias'
    Write-Host '   Perfil aplicado: versões novas em 365 dias, qualidade e segurança em 4 dias.' -ForegroundColor Gray
}

function Invoke-OpcoesGui {
    param([string[]]$Ids)
    if (-not $Script:EhAdmin) { Write-Host 'Estas ações exigem executar o speedPC como administrador.' -ForegroundColor Yellow; return }
    foreach ($id in $Ids) {
        $chave = Get-CorrecaoOpcao $id
        if ($chave -and $Script:Correcoes.Contains($chave)) { Invoke-Correcao $Script:Correcoes[$chave] }
    }
    Write-Host ''
    Write-Host 'Concluído.' -ForegroundColor Green
}

function Disable-SugestoesWindows {
    $valores = 'SilentInstalledAppsEnabled', 'SystemPaneSuggestionsEnabled', 'SoftLandingEnabled', 'SubscribedContent-338388Enabled', 'SubscribedContent-338389Enabled', 'SubscribedContent-353694Enabled', 'SubscribedContent-353696Enabled', 'SubscribedContent-310093Enabled', 'OemPreInstalledAppsEnabled', 'PreInstalledAppsEnabled'
    foreach ($h in (Get-HivesUsuarios)) {
        $k = "Registry::HKEY_USERS\$($h.Sid)\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
        foreach ($v in $valores) { Set-RegistroSeguro -Caminho $k -Nome $v -Valor 0 -Tipo DWord -Descricao "Sugestões: $v = 0 ($($h.Usuario))" }
    }
    Set-RegistroSeguro -Caminho 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\CloudContent' -Nome 'DisableWindowsConsumerFeatures' -Valor 1 -Tipo DWord -Descricao 'Política: desativar recursos de consumidor'
}

function Enable-SensorArmazenamento {
    foreach ($h in (Get-HivesUsuarios)) {
        $k = "Registry::HKEY_USERS\$($h.Sid)\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy"
        Set-RegistroSeguro -Caminho $k -Nome '01' -Valor 1 -Tipo DWord -Descricao "Sensor de Armazenamento ligado ($($h.Usuario))"
        Set-RegistroSeguro -Caminho $k -Nome '04' -Valor 1 -Tipo DWord -Descricao "Sensor: apagar temporários ($($h.Usuario))"
        Set-RegistroSeguro -Caminho $k -Nome '08' -Valor 0 -Tipo DWord -Descricao "Sensor: não mexer na Lixeira ($($h.Usuario))"
        Set-RegistroSeguro -Caminho $k -Nome '32' -Valor 0 -Tipo DWord -Descricao "Sensor: não mexer em Downloads ($($h.Usuario))"
    }
}

#endregion

#region ===================== Reparo e atualizações =====================

function Get-PastaSistema {
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) { return "$env:SystemRoot\Sysnative" }
    return "$env:SystemRoot\System32"
}

function Invoke-ReparoSistema {
    $sys = Get-PastaSistema
    Write-Host '   Etapa 1/2: DISM /RestoreHealth (de 5 a 30 minutos)...' -ForegroundColor Gray
    $d = Start-Ferramenta "$sys\dism.exe" @('/Online', '/Cleanup-Image', '/RestoreHealth')
    Write-Host '   Etapa 2/2: SFC /scannow (de 5 a 15 minutos)...' -ForegroundColor Gray
    $s = Start-Ferramenta "$sys\sfc.exe" @('/scannow')
    $msg = "DISM código $($d.ExitCode); SFC código $($s.ExitCode)"
    Write-Host "   $msg. Detalhes: $env:SystemRoot\Logs\CBS\CBS.log" -ForegroundColor Gray
    Add-Acao 'Reparo do Windows - resultado' $msg
}

function Invoke-VerificacaoDisco {
    $sys = Get-PastaSistema
    Write-Host "   CHKDSK $env:SystemDrive /scan (online, sem reiniciar)..." -ForegroundColor Gray
    $p = Start-Ferramenta "$sys\chkdsk.exe" @("$env:SystemDrive", '/scan')
    switch ($p.ExitCode) {
        0 { Write-Host '   Nenhum erro encontrado no sistema de arquivos.' -ForegroundColor Green; Add-Acao 'CHKDSK - resultado' 'Sem erros' }
        1 { Write-Host '   Erros encontrados e corrigidos.' -ForegroundColor Green; Add-Acao 'CHKDSK - resultado' 'Erros corrigidos' }
        default {
            Write-Host "   CHKDSK encontrou problemas que exigem reinício (código $($p.ExitCode))." -ForegroundColor Yellow
            Add-Acao 'CHKDSK - resultado' "Código $($p.ExitCode): correção requer reinício"
            if (Confirm-SN "Agendar a correção do disco $env:SystemDrive para o próximo reinício?" 'S') {
                fsutil dirty set $env:SystemDrive | Out-Null
                Add-Acao 'CHKDSK' 'Correção agendada para o próximo reinício'
            }
        }
    }
}

function Invoke-ReparoWindowsUpdate {
    $sufixo = '.bak_' + (Get-Date -Format 'yyyyMMdd_HHmm')
    $servicos = 'wuauserv', 'bits', 'cryptsvc', 'msiserver'
    Write-Host '   Parando serviços do Windows Update...' -ForegroundColor Gray
    Stop-Service -Name $servicos -Force -ErrorAction SilentlyContinue
    try {
        foreach ($p in "$env:SystemRoot\SoftwareDistribution", "$env:SystemRoot\System32\catroot2") {
            if (Test-Path -LiteralPath $p) { Rename-Item -LiteralPath $p -NewName ((Split-Path $p -Leaf) + $sufixo) -ErrorAction Stop; Write-Host "   $p -> $(Split-Path $p -Leaf)$sufixo" -ForegroundColor Gray }
        }
    } finally {
        Start-Service -Name 'cryptsvc', 'bits', 'wuauserv' -ErrorAction SilentlyContinue
    }
}

function Install-AtualizacoesWU {
    param([ValidateSet('Software', 'Driver')][string]$Tipo = 'Software', [switch]$SemConfirmar)
    $rotulo = 'atualizações do Windows'; if ($Tipo -eq 'Driver') { $rotulo = 'drivers' }
    $wu = Get-Service wuauserv -ErrorAction SilentlyContinue
    if ($wu.StartType -eq 'Disabled') { throw 'O serviço Windows Update está desativado.' }
    Write-Host "   Procurando $rotulo no Windows Update (pode levar alguns minutos)..." -ForegroundColor Gray
    $sessao = New-Object -ComObject Microsoft.Update.Session
    $sessao.ClientApplicationID = 'speedPC'
    $res = $sessao.CreateUpdateSearcher().Search("IsInstalled=0 and IsHidden=0 and Type='$Tipo'")
    $lista = @()
    foreach ($u in $res.Updates) {
        $cats = @($u.Categories | ForEach-Object { $_.Name })
        if ($cats -contains 'Upgrades' -or $u.Title -match 'Preview|Visualiza') { continue }   # sem versões novas do Windows nem prévias
        $lista += $u
    }
    if (-not $lista) { Write-Host "   Nenhuma pendência de $rotulo." -ForegroundColor Green; Add-Acao "Windows Update ($rotulo)" 'Nada pendente'; return }
    foreach ($u in $lista) { Write-Host ('   - {0} ({1})' -f $u.Title, (Format-Bytes $u.MaxDownloadSize)) }
    if (-not $SemConfirmar -and -not (Confirm-SN "Instalar $($lista.Count) item(ns)?" 'S')) { Add-Acao "Windows Update ($rotulo)" 'Cancelado'; return }
    $col = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($u in $lista) { if (-not $u.EulaAccepted) { $u.AcceptEula() }; [void]$col.Add($u) }
    Write-Host '   Baixando...' -ForegroundColor Gray
    $dl = $sessao.CreateUpdateDownloader(); $dl.Updates = $col; [void]$dl.Download()
    Write-Host '   Instalando...' -ForegroundColor Gray
    $inst = $sessao.CreateUpdateInstaller(); $inst.Updates = $col
    $r = $inst.Install()
    $ok = 0; $falha = 0
    for ($i = 0; $i -lt $col.Count; $i++) {
        $rc = $r.GetUpdateResult($i).ResultCode
        if ($rc -in 2, 3) { $ok++ } else { $falha++; Write-Host "   Falhou: $($col.Item($i).Title)" -ForegroundColor Yellow }
    }
    $msg = "$ok instalado(s), $falha falha(s)"
    if ($r.RebootRequired) { $msg += '; REINÍCIO NECESSÁRIO'; Write-Host '   Reinicie o computador para concluir.' -ForegroundColor Yellow }
    Write-Host "   $msg" -ForegroundColor Green
    Add-Acao "Windows Update ($rotulo)" $msg
}

function Get-CaminhoWinget {
    # winget é um alias por usuário; ao elevar com outra conta ele pode não estar no PATH
    $c = Get-Command winget -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    $f = Get-ChildItem "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending | Select-Object -First 1
    if ($f) { return $f.FullName }
    return $null
}

function Get-AppsInstalados {
    # Ids (do catálogo winget/msstore) dos programas instalados nesta máquina
    $wg = Get-CaminhoWinget
    if (-not $wg) { Write-Host 'winget não encontrado: instale o "Instalador de Aplicativo" pela Microsoft Store.' -ForegroundColor Yellow; return @() }
    $tmp = Join-Path $env:TEMP ("speedpc_apps_{0}.json" -f [guid]::NewGuid().ToString('N'))
    try {
        $p = Start-Process -FilePath $wg -ArgumentList @('export', '-o', "`"$tmp`"", '--accept-source-agreements', '--include-versions') -WindowStyle Hidden -Wait -PassThru
        if (-not (Test-Path -LiteralPath $tmp)) { Write-Host "Não foi possível listar os programas instalados (código $($p.ExitCode))." -ForegroundColor Yellow; return @() }
        $j = Get-Content -LiteralPath $tmp -Raw -Encoding UTF8 | ConvertFrom-Json
        return @($j.Sources | ForEach-Object { $_.Packages } | ForEach-Object { [string]$_.PackageIdentifier })
    } catch {
        Write-Host "Falha ao listar programas: $($_.Exception.Message)" -ForegroundColor Yellow
        return @()
    } finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
}

function Install-Apps {
    param([string[]]$Ids, [switch]$Desinstalar)
    $wg = Get-CaminhoWinget
    if (-not $wg) { Write-Host 'winget não encontrado: instale o "Instalador de Aplicativo" pela Microsoft Store.' -ForegroundColor Yellow; return }
    $todos = @{}; foreach ($c in $Script:CatalogoApps) { foreach ($a in $c.Apps) { $todos[$a.Id] = $a } }
    $verbo = 'Instalando'; if ($Desinstalar) { $verbo = 'Desinstalando' }
    $ok = 0; $falhas = 0; $i = 0
    foreach ($id in $Ids) {
        $i++
        $app = $todos[$id]; $nome = $id; if ($app) { $nome = $app.Nome }
        $fonte = 'winget'; if ($app -and $app.Fonte) { $fonte = $app.Fonte }
        Write-Host ("> {0} {1} ({2}/{3})..." -f $verbo, $nome, $i, $Ids.Count) -ForegroundColor Gray
        $acao = 'install'; if ($Desinstalar) { $acao = 'uninstall' }
        $argumentos = @($acao, '--id', $id, '--exact', '--silent', '--source', $fonte, '--accept-source-agreements')
        if (-not $Desinstalar) { $argumentos += '--accept-package-agreements' }
        $saida = Join-Path $env:TEMP ("speedpc_wg_{0}.txt" -f [guid]::NewGuid().ToString('N'))
        $p = Start-Process -FilePath $wg -ArgumentList $argumentos -WindowStyle Hidden -Wait -PassThru -RedirectStandardOutput $saida
        $codigo = $p.ExitCode
        switch ($codigo) {
            0 { $ok++; Write-Host "   $nome`: concluído." -ForegroundColor Green }
            -1978335189 { $ok++; Write-Host "   $nome`: já está na versão mais recente." -ForegroundColor Green }
            -1978335135 { $ok++; Write-Host "   $nome`: já está instalado." -ForegroundColor Green }
            -1978335212 { $falhas++; Write-Host "   $nome`: não encontrado no repositório." -ForegroundColor Red }
            default {
                $falhas++
                $msg = ''
                try { $msg = (Get-Content -LiteralPath $saida -Encoding UTF8 -ErrorAction Stop | Where-Object { $_ -match '\w{3}' -and $_ -notmatch '[█▒]' } | Select-Object -Last 1) } catch { }
                Write-Host ("   {0}: falhou (código {1}). {2}" -f $nome, $codigo, $msg) -ForegroundColor Red
            }
        }
        Remove-Item -LiteralPath $saida -Force -ErrorAction SilentlyContinue
    }
    $rotulo = 'Instalar aplicativos'; if ($Desinstalar) { $rotulo = 'Desinstalar aplicativos' }
    Add-Acao $rotulo ("{0} ok, {1} falha(s): {2}" -f $ok, $falhas, ($Ids -join ', '))
    Write-Host ''
    Write-Host ("Concluído: {0} ok, {1} falha(s)." -f $ok, $falhas) -ForegroundColor $(if ($falhas) { 'Yellow' } else { 'Green' })
}

function Invoke-AtualizacaoProgramas {
    $wgCaminho = Get-CaminhoWinget
    if (-not $wgCaminho) { Write-Host '   winget não está disponível nesta máquina (Instalador de Aplicativo da Microsoft Store).' -ForegroundColor Yellow; return }
    $wg = [pscustomobject]@{ Source = $wgCaminho }
    Set-Alias -Name winget -Value $wgCaminho -Scope Local
    Write-Host '   Consultando programas com atualização disponível...' -ForegroundColor Gray
    $lista = (winget upgrade --accept-source-agreements 2>&1 | Out-String) -replace '[^\x20-\x7EÀ-ÿ\r\n]', ''
    Write-Host $lista -ForegroundColor Gray
    if (Confirm-SN 'Atualizar todos os programas listados no registro de atividade?' 'N') {
        $p = Start-Ferramenta $wg.Source @('upgrade', '--all', '--silent', '--accept-source-agreements', '--accept-package-agreements')
        Add-Acao 'Atualizar programas (winget)' "Código $($p.ExitCode)"
    }
}

function Invoke-MenuReparo {
    Write-Titulo 'Reparo do Windows'
    if (-not $Script:EhAdmin) { Write-Host '  Requer administrador.' -ForegroundColor Yellow; return }
    $ops = @(
        (Get-CorrecaoPadrao 'ReparoWindows'), (Get-CorrecaoPadrao 'VerificarDisco'), (Get-CorrecaoPadrao 'ReparoWU'),
        (Get-CorrecaoPadrao 'RenovarRede'), (Get-CorrecaoPadrao 'ResetRede'), (Get-CorrecaoPadrao 'LimparFila'), (Get-CorrecaoPadrao 'TesteMemoria')
    ) | ForEach-Object { $Script:Correcoes[$_] }
    $ops = @($ops)
    $sel = @(Select-Opcoes -Titulo 'Reparos disponíveis' -Itens @($ops | ForEach-Object { [pscustomobject]@{ Nome = $_.Nome; Detalhe = $_.Descricao } }))
    foreach ($i in $sel) { Invoke-Correcao $ops[$i] }
}

function Invoke-MenuAtualizacoes {
    Write-Titulo 'Atualizações e drivers'
    if (-not $Script:EhAdmin) { Write-Host '  Requer administrador.' -ForegroundColor Yellow; return }
    $itens = @(
        [pscustomobject]@{ Nome = 'Atualizações do Windows'; Detalhe = 'Segurança e qualidade, sem versões de prévia e sem troca de versão do Windows.' }
        [pscustomobject]@{ Nome = 'Drivers pelo Windows Update'; Detalhe = 'Drivers publicados pela Microsoft e pelos fabricantes.' }
        [pscustomobject]@{ Nome = 'Programas instalados (winget)'; Detalhe = 'Atualiza navegadores, leitores de PDF etc. pelo gerenciador de pacotes do Windows.' }
    )
    foreach ($i in @(Select-Opcoes -Titulo 'O que atualizar?' -Itens $itens)) {
        try {
            switch ($i) {
                0 { Install-AtualizacoesWU -Tipo Software }
                1 { Install-AtualizacoesWU -Tipo Driver }
                2 { Invoke-AtualizacaoProgramas }
            }
        } catch { Write-Host "   Falhou: $($_.Exception.Message)" -ForegroundColor Red; Add-Acao 'Atualizações' "Falhou: $($_.Exception.Message)" }
    }
}

#endregion

#region ===================== Relatório HTML =====================

function ConvertTo-Html2 {
    # Escapa só os caracteres especiais do HTML; mantém os acentos legíveis no arquivo (UTF-8)
    param($Texto)
    return ([string]$Texto).Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&#39;')
}

function ConvertTo-TabelaHtml {
    param($Linhas)
    $Linhas = @($Linhas | Where-Object { $_ })
    if (-not $Linhas) { return '<p class="vazio">Sem dados.</p>' }
    $props = @($Linhas[0].PSObject.Properties.Name)
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append('<div class="tabela"><table><thead><tr>')
    foreach ($p in $props) { [void]$sb.Append('<th>' + (ConvertTo-Html2 $p) + '</th>') }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($l in $Linhas) {
        [void]$sb.Append('<tr>')
        foreach ($p in $props) { [void]$sb.Append('<td>' + (ConvertTo-Html2 $l.$p) + '</td>') }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table></div>')
    return $sb.ToString()
}

function New-RelatorioHtml {
    param([string]$Pasta = $Script:PastaRelLocal)
    if (-not $Script:DiagnosticoFeito) { Invoke-Diagnostico }
    $r = $Script:Dados.Resumo
    $agora = Get-Date
    $prob = @(Get-AchadosOrdenados -SomenteProblemas)
    $ok = @($Script:Achados | Where-Object { $_.Severidade -eq 'OK' })
    $nCrit = @($prob | Where-Object { $_.Severidade -eq 'Critico' }).Count
    $nAten = @($prob | Where-Object { $_.Severidade -eq 'Atencao' }).Count
    $nInfo = @($prob | Where-Object { $_.Severidade -eq 'Info' }).Count
    $executor = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $modo = 'Manual'; if ($Script:ModoAuto) { $modo = 'Automático (agendado)' }
    $h = { param($t) ConvertTo-Html2 $t }

    $css = @'
:root{--fundo:#f4f6f9;--papel:#fff;--texto:#1d2733;--suave:#5b6878;--linha:#e2e7ee;--crit:#c62828;--crit-bg:#fdecec;--aten:#b26a00;--aten-bg:#fff4e0;--info:#0b6aa2;--info-bg:#e7f3fb;--ok:#2e7d32;--ok-bg:#e8f5e9;--marca:#0f3b5f}
*{box-sizing:border-box}body{margin:0;background:var(--fundo);color:var(--texto);font:14px/1.5 "Segoe UI",system-ui,-apple-system,Roboto,Arial,sans-serif}
.pagina{max-width:1080px;margin:0 auto;padding:24px 16px 48px}
header{background:var(--marca);color:#fff;border-radius:10px;padding:20px 24px;margin-bottom:18px}
header .marca{font-size:12px;letter-spacing:.12em;text-transform:uppercase;opacity:.8}header h1{margin:4px 0 2px;font-size:26px}header p{margin:0;opacity:.85;font-size:13px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin-bottom:18px}
.card{background:var(--papel);border:1px solid var(--linha);border-radius:10px;padding:14px 16px}
.card .rot{font-size:12px;color:var(--suave);text-transform:uppercase;letter-spacing:.05em}.card .val{font-size:20px;font-weight:600;margin-top:2px}.card .sub{font-size:12.5px;color:var(--suave);margin-top:2px}
.card.ruim .val{color:var(--crit)}.card.medio .val{color:var(--aten)}.card.bom .val{color:var(--ok)}
.barra{height:6px;background:var(--linha);border-radius:3px;margin-top:8px;overflow:hidden}.barra span{display:block;height:100%;background:var(--ok)}.card.medio .barra span{background:var(--aten)}.card.ruim .barra span{background:var(--crit)}
section{background:var(--papel);border:1px solid var(--linha);border-radius:10px;padding:18px 20px;margin-bottom:18px}
h2{font-size:17px;margin:0 0 12px}h3{font-size:15px;margin:0}h4{font-size:13px;margin:16px 0 6px;color:var(--suave);text-transform:uppercase;letter-spacing:.05em}
.contagem{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px}.pill{padding:3px 10px;border-radius:999px;font-size:12.5px;font-weight:600}
.pill.critico{background:var(--crit-bg);color:var(--crit)}.pill.atencao{background:var(--aten-bg);color:var(--aten)}.pill.info{background:var(--info-bg);color:var(--info)}.pill.ok{background:var(--ok-bg);color:var(--ok)}
.achado{border:1px solid var(--linha);border-left:5px solid var(--info);border-radius:8px;padding:10px 14px;margin:8px 0;break-inside:avoid}
.achado.critico{border-left-color:var(--crit)}.achado.atencao{border-left-color:var(--aten)}.achado.ok{border-left-color:var(--ok)}
.achado .topo{display:flex;gap:10px;align-items:baseline;flex-wrap:wrap}.achado p{margin:6px 0 0}.achado .rotulo{font-weight:600;color:var(--suave)}.achado .auto{display:inline-block;margin-top:6px;font-size:12px;color:var(--info)}
.tabela{overflow-x:auto}table{border-collapse:collapse;width:100%;font-size:13px}th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--linha);vertical-align:top}th{background:#f7f9fb;font-weight:600;white-space:nowrap}
dl.ficha{display:grid;grid-template-columns:max-content 1fr;gap:4px 18px;margin:0}dl.ficha dt{color:var(--suave)}dl.ficha dd{margin:0;overflow-wrap:anywhere}
ul.ok{columns:2;margin:0;padding-left:18px;color:var(--suave)}ul.ok li{break-inside:avoid}.vazio{color:var(--suave);margin:0}
footer{text-align:center;color:var(--suave);font-size:12px}
@media (max-width:640px){ul.ok{columns:1}dl.ficha{grid-template-columns:1fr}dl.ficha dt{margin-top:6px}}
@media print{body{background:#fff}.pagina{padding:0}section,.card,header{break-inside:avoid}header{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
'@

    $sb = New-Object Text.StringBuilder
    [void]$sb.Append('<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">')
    [void]$sb.Append("<title>speedPC - $(& $h $r.Computador) - $($agora.ToString('dd/MM/yyyy HH:mm'))</title><style>$css</style></head><body><div class=`"pagina`">")
    [void]$sb.Append("<header><div class=`"marca`">speedPC $Script:Versao &middot; Relatório de manutenção</div><h1>$(& $h $r.Computador)</h1>")
    [void]$sb.Append("<p>$(& $h $agora.ToString('dd/MM/yyyy HH:mm')) &middot; Executado por $(& $h $executor) &middot; Modo: $modo</p></header>")

    # Cards
    $cl = 'bom'; if ($r.TempoLigado.TotalDays -ge 7) { $cl = 'ruim' } elseif ($r.TempoLigado.TotalDays -ge 3) { $cl = 'medio' }
    $cd = 'bom'; if ($r.Desligamento.Tipo -eq 'Inesperado') { $cd = 'ruim' } elseif ($r.Desligamento.Tipo -eq 'Desconhecido') { $cd = '' }
    $cr = 'bom'; if ($r.RamPct -ge 90) { $cr = 'ruim' } elseif ($r.RamPct -ge 80) { $cr = 'medio' }
    $cp = 'bom'; if ($nCrit) { $cp = 'ruim' } elseif ($nAten) { $cp = 'medio' }
    $txtDesl = 'Normal'; if ($r.Desligamento.Tipo -eq 'Inesperado') { $txtDesl = 'Inesperado' } elseif ($r.Desligamento.Tipo -eq 'Desconhecido') { $txtDesl = 'Sem registro' }
    $subDesl = ''; if ($r.Desligamento.Data) { $subDesl = $r.Desligamento.Texto -replace '^(Normal|INESPERADO) - ', '' }
    [void]$sb.Append('<div class="cards">')
    [void]$sb.Append("<div class=`"card $cl`"><div class=`"rot`">Ligado há</div><div class=`"val`">$(& $h (Format-Duracao $r.TempoLigado))</div><div class=`"sub`">desde $(& $h $r.Boot.ToString('dd/MM/yyyy HH:mm'))</div></div>")
    [void]$sb.Append("<div class=`"card $cd`"><div class=`"rot`">Último desligamento</div><div class=`"val`">$txtDesl</div><div class=`"sub`">$(& $h $subDesl)</div></div>")
    [void]$sb.Append("<div class=`"card $cr`"><div class=`"rot`">Memória RAM</div><div class=`"val`">$($r.RamPct)%</div><div class=`"sub`">$(& $h (Format-Bytes $r.RamUsada)) de $(& $h (Format-Bytes $r.RamTotal))</div><div class=`"barra`"><span style=`"width:$($r.RamPct)%`"></span></div></div>")
    if ($r.DiscoSistema -and $r.DiscoSistema.Size) {
        $uso = [math]::Round(($r.DiscoSistema.Size - $r.DiscoSistema.FreeSpace) / $r.DiscoSistema.Size * 100)
        $cdk = 'bom'; if ($uso -gt 90) { $cdk = 'ruim' } elseif ($uso -gt 85) { $cdk = 'medio' }
        [void]$sb.Append("<div class=`"card $cdk`"><div class=`"rot`">Disco $env:SystemDrive</div><div class=`"val`">$(& $h (Format-Bytes $r.DiscoSistema.FreeSpace)) livres</div><div class=`"sub`">$uso% usado de $(& $h (Format-Bytes $r.DiscoSistema.Size))</div><div class=`"barra`"><span style=`"width:$uso%`"></span></div></div>")
    }
    [void]$sb.Append("<div class=`"card $cp`"><div class=`"rot`">Problemas</div><div class=`"val`">$($nCrit + $nAten)</div><div class=`"sub`">$nCrit crítico(s), $nAten de atenção</div></div>")
    [void]$sb.Append('</div>')

    if ($r.Desligamento.Observacao) {
        [void]$sb.Append("<section><div class=`"achado atencao`"><p>$(& $h $r.Desligamento.Observacao)</p></div></section>")
    }

    # Problemas
    [void]$sb.Append('<section><h2>Problemas encontrados</h2><div class="contagem">')
    [void]$sb.Append("<span class=`"pill critico`">$nCrit crítico(s)</span><span class=`"pill atencao`">$nAten atenção</span><span class=`"pill info`">$nInfo informativo(s)</span><span class=`"pill ok`">$($ok.Count) OK</span></div>")
    if (-not $prob) { [void]$sb.Append('<p class="vazio">Nenhum problema encontrado.</p>') }
    $catAtual = ''
    foreach ($a in $prob) {
        if ($a.Categoria -ne $catAtual) { $catAtual = $a.Categoria; [void]$sb.Append("<h4>$(& $h $catAtual)</h4>") }
        $s = $Script:Sev[$a.Severidade]
        [void]$sb.Append("<div class=`"achado $($s.Css)`"><div class=`"topo`"><span class=`"pill $($s.Css)`">$($s.Rotulo)</span><h3>$(& $h $a.Titulo)</h3></div>")
        if ($a.Evidencia) { [void]$sb.Append("<p><span class=`"rotulo`">Evidência:</span> $(& $h $a.Evidencia)</p>") }
        if ($a.Recomendacao) { [void]$sb.Append("<p><span class=`"rotulo`">O que fazer:</span> $(& $h $a.Recomendacao)</p>") }
        if ($a.Correcoes) {
            $nomes = @($a.Correcoes | ForEach-Object { $c = $Script:Correcoes[$_]; if ($c) { if ($c.Aplicada) { "$($c.Nome) (aplicada)" } else { $c.Nome } } }) -join '; '
            [void]$sb.Append("<span class=`"auto`">Correção automática: $(& $h $nomes)</span>")
        }
        [void]$sb.Append('</div>')
    }
    [void]$sb.Append('</section>')

    # Ações
    [void]$sb.Append('<section><h2>Ações executadas nesta sessão</h2>')
    if ($Script:Acoes.Count) {
        [void]$sb.Append((ConvertTo-TabelaHtml ($Script:Acoes | ForEach-Object { [pscustomobject]@{ Hora = $_.Hora.ToString('HH:mm:ss'); 'Ação' = $_.Acao; Resultado = $_.Resultado } })))
        if ($Script:ArquivoBackup) { [void]$sb.Append("<p class=`"vazio`">Backup das configurações alteradas: $(& $h $Script:ArquivoBackup) (menu Desfazer).</p>") }
    } else { [void]$sb.Append('<p class="vazio">Nenhuma alteração foi feita: somente diagnóstico.</p>') }
    [void]$sb.Append('</section>')

    # Ficha
    [void]$sb.Append('<section><h2>Ficha da máquina</h2><dl class="ficha">')
    $ficha = [ordered]@{
        'Computador' = $r.Computador; 'Fabricante / modelo' = "$($r.Fabricante) $($r.Modelo)"; 'Número de série' = $r.Serie; 'BIOS' = $r.Bios
        'Windows' = $r.Windows; 'Processador' = "$($r.Processador) ($($r.Nucleos) threads)"; 'Memória RAM' = (Format-Bytes $r.RamTotal)
        'Domínio' = $r.Dominio; 'Usuário logado' = $r.UsuarioLogado; 'Endereços IP' = $r.IPs
        'Último boot' = $r.Boot.ToString('dd/MM/yyyy HH:mm'); 'Último desligamento' = $r.Desligamento.Texto
        'Inicialização Rápida' = $(if ($r.InicioRapido) { 'Ativa' } else { 'Desativada' })
    }
    if ($r.Desligamento.Detalhe) { $ficha['Motivo do desligamento'] = $r.Desligamento.Detalhe }
    foreach ($k in $ficha.Keys) { [void]$sb.Append("<dt>$(& $h $k)</dt><dd>$(& $h $ficha[$k])</dd>") }
    [void]$sb.Append('</dl></section>')

    # Detalhes
    if ($Script:Dados.Tabelas.Count) {
        [void]$sb.Append('<section><h2>Detalhes coletados</h2>')
        foreach ($t in $Script:Dados.Tabelas.Keys) {
            [void]$sb.Append("<h4>$(& $h $t)</h4>")
            [void]$sb.Append((ConvertTo-TabelaHtml $Script:Dados.Tabelas[$t]))
        }
        [void]$sb.Append('</section>')
    }

    # OK
    if ($ok) {
        [void]$sb.Append('<section><h2>Verificações sem problema</h2><ul class="ok">')
        foreach ($a in ($ok | Sort-Object @{ e = { [array]::IndexOf($Script:OrdemCategorias, $_.Categoria) } })) { [void]$sb.Append("<li>$(& $h $a.Categoria): $(& $h $a.Titulo)</li>") }
        [void]$sb.Append('</ul></section>')
    }
    [void]$sb.Append("<footer>Gerado por speedPC $Script:Versao &middot; ferramenta gratuita em PowerShell, sem programas de terceiros.</footer></div></body></html>")

    if (-not (Test-Path -LiteralPath $Pasta)) { New-Item -ItemType Directory -Path $Pasta -Force | Out-Null }
    $arq = Join-Path $Pasta ('speedPC_{0}_{1:yyyy-MM-dd_HHmm}.html' -f $env:COMPUTERNAME, $agora)
    [IO.File]::WriteAllText($arq, $sb.ToString(), (New-Object Text.UTF8Encoding($true)))

    # Resumo em JSON (para consolidar várias máquinas numa planilha/painel)
    $json = [pscustomobject]@{
        Computador = $r.Computador; Data = $agora.ToString('s'); Modo = $modo; Usuario = $r.UsuarioLogado
        Modelo = "$($r.Fabricante) $($r.Modelo)"; Serie = $r.Serie; Windows = $r.Windows
        DiasLigado = [math]::Round($r.TempoLigado.TotalDays, 1); UltimoDesligamento = $r.Desligamento.Tipo
        RamPct = $r.RamPct; Criticos = $nCrit; Atencao = $nAten
        Problemas = @($prob | Where-Object { $_.Severidade -in 'Critico', 'Atencao' } | ForEach-Object { "[$($_.Severidade)] $($_.Titulo)" })
        Relatorio = (Split-Path $arq -Leaf)
    } | ConvertTo-Json -Depth 4
    [IO.File]::WriteAllText((Join-Path $Pasta ("{0}_ultimo.json" -f $env:COMPUTERNAME)), $json, (New-Object Text.UTF8Encoding($true)))
    Add-Acao 'Relatório HTML' $arq
    return $arq
}

#endregion

#region ===================== Modo automático e agendamento =====================

function Invoke-ModoAutomatico {
    $log = Join-Path $Script:PastaLogs ('auto_{0:yyyy-MM-dd_HHmm}.log' -f (Get-Date))
    if (-not (Test-Path -LiteralPath $Script:PastaLogs)) { New-Item -ItemType Directory -Path $Script:PastaLogs -Force | Out-Null }
    try { Start-Transcript -Path $log -Force | Out-Null } catch { }
    Write-Log 'Modo automático iniciado'
    try {
        Get-Resumo | Out-Null
        Invoke-Diagnostico
        if (-not $SemLimpeza) { Invoke-LimpezaSegura -Silencioso }
        $arq = New-RelatorioHtml -Pasta $Script:PastaRelLocal
        if ($PastaRelatorios) {
            try {
                if (-not (Test-Path -LiteralPath $PastaRelatorios)) { New-Item -ItemType Directory -Path $PastaRelatorios -Force -ErrorAction Stop | Out-Null }
                Copy-Item -LiteralPath $arq -Destination $PastaRelatorios -Force -ErrorAction Stop
                Copy-Item -LiteralPath (Join-Path $Script:PastaRelLocal "$($env:COMPUTERNAME)_ultimo.json") -Destination $PastaRelatorios -Force -ErrorAction Stop
                Write-Log "Relatório copiado para $PastaRelatorios"
            } catch { Write-Log "Falha ao copiar para a pasta de rede ${PastaRelatorios}: $($_.Exception.Message)" }
        }
        # mantém só 90 dias de relatórios locais e logs desta ferramenta
        Get-ChildItem -LiteralPath $Script:PastaRelLocal -Filter 'speedPC_*.html' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) } | Remove-Item -Force -ErrorAction SilentlyContinue
        Get-ChildItem -LiteralPath $Script:PastaLogs -Filter 'auto_*.log' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) } | Remove-Item -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Log "Erro no modo automático: $($_.Exception.Message)"
    } finally {
        Write-Log 'Modo automático concluído'
        try { Stop-Transcript | Out-Null } catch { }
    }
}

function Register-ModoAutomatico {
    param([string]$Pasta, [string]$Dia = 'Wednesday', [string]$Horario = '12:00')
    if (-not $Script:EhAdmin) { Write-Host '  Agendar exige administrador.' -ForegroundColor Red; return }
    if (-not $Script:CaminhoScript) { Write-Host '  Execute o script a partir do arquivo .ps1 para agendar.' -ForegroundColor Red; return }
    $destDir = Join-Path $env:ProgramFiles 'speedPC'
    $destino = Join-Path $destDir 'speedPC.ps1'
    if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
    if ($Script:CaminhoScript -ne $destino) { Copy-Item -LiteralPath $Script:CaminhoScript -Destination $destino -Force }
    $argumentos = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$destino`" -Auto"
    if ($Pasta) { $argumentos += " -PastaRelatorios `"$($Pasta.TrimEnd('\'))`"" }
    $acao = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argumentos
    $gatilho = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $Dia -At $Horario
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $config = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 3)
    Register-ScheduledTask -TaskName $Script:NomeTarefa -TaskPath '\speedPC\' -Action $acao -Trigger $gatilho -Principal $principal -Settings $config `
        -Description 'Diagnóstico, limpeza segura e relatório semanal (speedPC).' -Force | Out-Null
    Write-Host "  Tarefa '$Script:NomeTarefa' criada: toda $Dia às $Horario, como SYSTEM." -ForegroundColor Green
    Write-Host "  Script instalado em: $destino" -ForegroundColor Gray
    if ($Pasta) {
        Write-Host "  Relatórios: $Pasta (e cópia local em $Script:PastaRelLocal)" -ForegroundColor Gray
        Write-Host '  IMPORTANTE: a tarefa roda como SYSTEM, então quem grava na pasta de rede é a CONTA DO COMPUTADOR.' -ForegroundColor Yellow
        Write-Host '  Dê permissão de gravação ao grupo "Computadores do domínio" (Domain Computers) no compartilhamento e na pasta.' -ForegroundColor Yellow
    }
    Add-Acao 'Agendamento semanal' "Toda $Dia às $Horario; pasta: $Pasta"
}

function Invoke-MenuAgendamento {
    Write-Titulo 'Modo automático semanal'
    if (-not $Script:EhAdmin) { Write-Host '  Requer administrador.' -ForegroundColor Yellow; return }
    $t = Get-ScheduledTask -TaskPath '\speedPC\' -TaskName $Script:NomeTarefa -ErrorAction SilentlyContinue
    if ($t) {
        $info = $t | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue
        Write-Host "  Agendamento atual: $($t.Actions[0].Arguments)" -ForegroundColor Gray
        if ($info) { Write-Host "  Última execução: $($info.LastRunTime)  |  Próxima: $($info.NextRunTime)" -ForegroundColor Gray }
        Write-Host '  [1] Recriar   [2] Remover   [Enter] Voltar'
        $o = Read-Host '  Opção'
        if ($o -eq '2') { Unregister-ScheduledTask -TaskPath '\speedPC\' -TaskName $Script:NomeTarefa -Confirm:$false; Write-Host '  Agendamento removido.' -ForegroundColor Green; Add-Acao 'Agendamento semanal' 'Removido'; return }
        if ($o -ne '1') { return }
    }
    Write-Host '  O modo automático faz: diagnóstico + limpeza segura + relatório HTML. Não aplica correções.' -ForegroundColor Gray
    $pasta = Read-Host '  Pasta de rede para os relatórios (ex.: \\servidor\manutencao$; Enter = só local)'
    $dias = @{ '1' = 'Monday'; '2' = 'Tuesday'; '3' = 'Wednesday'; '4' = 'Thursday'; '5' = 'Friday'; '6' = 'Saturday'; '7' = 'Sunday' }
    $d = Read-Host '  Dia da semana: 1=Seg 2=Ter 3=Qua 4=Qui 5=Sex 6=Sáb 7=Dom (Enter = Qua)'
    $dia = 'Wednesday'; if ($dias.ContainsKey($d)) { $dia = $dias[$d] }
    $hr = Read-Host '  Horário HH:mm (Enter = 12:00, horário de almoço)'
    if ($hr -notmatch '^\d{1,2}:\d{2}$') { $hr = '12:00' }
    if ($pasta) {
        $pasta = $pasta.Trim().TrimEnd('\')
        try {
            $teste = Join-Path $pasta ".speedpc_teste_$env:COMPUTERNAME"
            Set-Content -LiteralPath $teste -Value 'ok' -ErrorAction Stop; Remove-Item -LiteralPath $teste -Force -ErrorAction SilentlyContinue
            Write-Host '  Gravação na pasta testada com SEU usuário: OK.' -ForegroundColor Green
        } catch { Write-Host "  Aviso: seu usuário não conseguiu gravar em $pasta ($($_.Exception.Message))." -ForegroundColor Yellow }
    }
    Register-ModoAutomatico -Pasta $pasta -Dia $dia -Horario $hr
}

#endregion

#region ===================== Menu principal =====================

function Show-MenuPrincipal {
    Show-Resumo
    while ($true) {
        Write-Host ''
        Write-Host '  ---------------------------------- MENU ----------------------------------' -ForegroundColor DarkCyan
        Write-Host '  [1] Analisar problemas (causa provável + correções)'
        Write-Host '  [2] Limpeza segura (todos os usuários)'
        Write-Host '  [3] Otimização para o trabalho'
        Write-Host '  [4] Reparo do Windows (DISM, SFC, disco, Windows Update, rede, impressão)'
        Write-Host '  [5] Atualizações do Windows e drivers'
        Write-Host '  [6] Gerar relatório HTML'
        Write-Host '  [7] Modo automático semanal (agendar)'
        Write-Host '  [8] Desfazer alterações'
        Write-Host '  [9] Criar ponto de restauração agora'
        Write-Host '  [R] Atualizar resumo     [0] Sair'
        $op = Read-Host '  Opção'
        switch ($op.Trim().ToUpper()) {
            '1' {
                Invoke-Diagnostico
                if ($Script:Correcoes.Count -and $Script:EhAdmin) { Invoke-MenuCorrecoes }
                if (Confirm-SN 'Gerar o relatório HTML desta máquina?' 'S') { $a = New-RelatorioHtml; Write-Host "  Relatório: $a" -ForegroundColor Green; if (Confirm-SN 'Abrir o relatório?' 'S') { Start-Process $a } }
            }
            '2' { Invoke-LimpezaSegura }
            '3' { Invoke-Otimizacao }
            '4' { Invoke-MenuReparo }
            '5' { Invoke-MenuAtualizacoes }
            '6' {
                $destino = $Script:PastaRelLocal
                $p = Read-Host "  Pasta de destino (Enter = $destino)"
                if ($p) { $destino = $p.Trim() }
                try { $a = New-RelatorioHtml -Pasta $destino; Write-Host "  Relatório: $a" -ForegroundColor Green; if (Confirm-SN 'Abrir o relatório?' 'S') { Start-Process $a } }
                catch { Write-Host "  Falha ao salvar: $($_.Exception.Message)" -ForegroundColor Red }
            }
            '7' { Invoke-MenuAgendamento }
            '8' { Invoke-Desfazer }
            '9' { $Script:PontoRestauracaoFeito = $false; New-PontoRestauracao }
            'R' { Get-Resumo | Out-Null; Show-Resumo }
            '0' { Write-Log 'Sessão encerrada'; return }
            default { }
        }
    }
}

#endregion

#region ===================== Interface gráfica (WPF) =====================

$Script:XamlJanela = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="speedPC" Width="1220" Height="800" MinWidth="1020" MinHeight="640"
        WindowStartupLocation="CenterScreen" Background="#0F1115" Foreground="#E5E7EB"
        FontFamily="Segoe UI" FontSize="13" UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
  <Window.Resources>
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="#171A21"/>
      <Setter Property="BorderBrush" Value="#262B35"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="10"/>
      <Setter Property="Padding" Value="18"/>
      <Setter Property="Margin" Value="6"/>
    </Style>
    <Style x:Key="Rotulo" TargetType="TextBlock">
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#8B93A1"/>
    </Style>
    <Style x:Key="Valor" TargetType="TextBlock">
      <Setter Property="FontSize" Value="22"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,8,0,2"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>
    <Style x:Key="Sub" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Foreground" Value="#9CA3AF"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="TituloCard" TargetType="TextBlock">
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,12"/>
    </Style>
    <Style x:Key="TituloPagina" TargetType="TextBlock">
      <Setter Property="FontSize" Value="20"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="Icone" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="Segoe MDL2 Assets"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>

    <Style x:Key="Aba" TargetType="Button">
      <Setter Property="Foreground" Value="#9CA3AF"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Padding" Value="14,13"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Grid Background="Transparent">
              <ContentPresenter Margin="{TemplateBinding Padding}" VerticalAlignment="Center"/>
              <Border x:Name="Linha" Height="2" VerticalAlignment="Bottom" Background="#3B82F6" CornerRadius="1" Margin="10,0" Visibility="Hidden"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter Property="Foreground" Value="#E5E7EB"/></Trigger>
              <Trigger Property="Tag" Value="sel">
                <Setter Property="Foreground" Value="#FFFFFF"/>
                <Setter TargetName="Linha" Property="Visibility" Value="Visible"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Primario" TargetType="Button">
      <Setter Property="Background" Value="#3B82F6"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderBrush" Value="#3B82F6"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="16,8"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="8,0,0,0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Fundo" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="7" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Fundo" Property="Opacity" Value="0.85"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="Fundo" Property="Opacity" Value="0.7"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Fundo" Property="Opacity" Value="0.35"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Secundario" TargetType="Button" BasedOn="{StaticResource Primario}">
      <Setter Property="Background" Value="#232833"/>
      <Setter Property="Foreground" Value="#E5E7EB"/>
      <Setter Property="BorderBrush" Value="#343A46"/>
      <Setter Property="FontWeight" Value="Normal"/>
    </Style>
    <Style x:Key="Link" TargetType="Button">
      <Setter Property="Foreground" Value="#60A5FA"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <ContentPresenter x:Name="C" VerticalAlignment="Center"/>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter Property="Foreground" Value="#93C5FD"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Opcao" TargetType="CheckBox">
      <Setter Property="Foreground" Value="#E5E7EB"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="0,0,0,14"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <Grid Background="Transparent">
              <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <Border x:Name="Caixa" Width="18" Height="18" CornerRadius="4" BorderThickness="1.5" BorderBrush="#4B5563"
                      Background="#0F1115" VerticalAlignment="Top" Margin="0,1,12,0">
                <TextBlock x:Name="Marca" Text="&#xE73E;" FontFamily="Segoe MDL2 Assets" FontSize="11" Foreground="White"
                           HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Hidden"/>
              </Border>
              <ContentPresenter Grid.Column="1"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Caixa" Property="BorderBrush" Value="#60A5FA"/></Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="Caixa" Property="Background" Value="#3B82F6"/>
                <Setter TargetName="Caixa" Property="BorderBrush" Value="#3B82F6"/>
                <Setter TargetName="Marca" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.6"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Chip" TargetType="RadioButton">
      <Setter Property="Foreground" Value="#C9CED6"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="0,0,6,6"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="F" Background="#1F232C" BorderBrush="#343A46" BorderThickness="1" CornerRadius="14" Padding="13,5">
              <ContentPresenter HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="F" Property="BorderBrush" Value="#60A5FA"/></Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="F" Property="Background" Value="#3B82F6"/>
                <Setter TargetName="F" Property="BorderBrush" Value="#3B82F6"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#0F1115"/>
      <Setter Property="Foreground" Value="#E5E7EB"/>
      <Setter Property="BorderBrush" Value="#343A46"/>
      <Setter Property="CaretBrush" Value="White"/>
      <Setter Property="Padding" Value="6,0"/>
      <Setter Property="Height" Value="34"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="B" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="6">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="B" Property="BorderBrush" Value="#3B82F6"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ScrollBar">
      <Setter Property="Width" Value="10"/>
      <Setter Property="MinWidth" Value="10"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Grid Background="Transparent">
              <Track x:Name="PART_Track" IsDirectionReversed="True">
                <Track.Thumb>
                  <Thumb>
                    <Thumb.Template>
                      <ControlTemplate TargetType="Thumb"><Border Background="#353B47" CornerRadius="4" Margin="2"/></ControlTemplate>
                    </Thumb.Template>
                  </Thumb>
                </Track.Thumb>
              </Track>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Barra" TargetType="ProgressBar">
      <Setter Property="Height" Value="6"/>
      <Setter Property="Foreground" Value="#3B82F6"/>
      <Setter Property="Maximum" Value="100"/>
      <Setter Property="Margin" Value="0,12,0,0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ProgressBar">
            <Grid>
              <Border x:Name="PART_Track" Background="#262B35" CornerRadius="3"/>
              <Border x:Name="PART_Indicator" Background="{TemplateBinding Foreground}" CornerRadius="3" HorizontalAlignment="Left"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="ItemLista" TargetType="ListBoxItem">
      <Setter Property="Foreground" Value="#E5E7EB"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListBoxItem">
            <Border x:Name="F" Background="#1C2029" BorderBrush="#2A2F3A" BorderThickness="1" CornerRadius="8" Padding="12,9" Margin="0,0,0,6">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="F" Property="BorderBrush" Value="#3B4350"/></Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="F" Property="Background" Value="#1E2B42"/>
                <Setter TargetName="F" Property="BorderBrush" Value="#3B82F6"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid x:Name="Raiz" Background="#0F1115">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Cabeçalho -->
    <Grid Margin="24,18,24,10">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <StackPanel Orientation="Horizontal">
        <Border Width="38" Height="38" CornerRadius="9">
          <Border.Background>
            <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
              <GradientStop Color="#3B82F6" Offset="0"/>
              <GradientStop Color="#6366F1" Offset="1"/>
            </LinearGradientBrush>
          </Border.Background>
          <TextBlock Text="&#xE90F;" Style="{StaticResource Icone}" FontSize="18" Foreground="White" HorizontalAlignment="Center"/>
        </Border>
        <StackPanel Margin="12,0,0,0" VerticalAlignment="Center">
          <TextBlock Text="speedPC" FontSize="18" FontWeight="SemiBold"/>
          <TextBlock Text="Manutenção de computadores" FontSize="11.5" Foreground="#8B93A1"/>
        </StackPanel>
      </StackPanel>
      <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
        <StackPanel HorizontalAlignment="Right" Margin="0,0,16,0">
          <TextBlock x:Name="TxtPc" Text="Carregando..." FontSize="14" FontWeight="SemiBold" HorizontalAlignment="Right"/>
          <TextBlock x:Name="TxtSub" Text="" FontSize="11.5" Foreground="#8B93A1" HorizontalAlignment="Right"/>
        </StackPanel>
        <Border x:Name="SeloAdmin" CornerRadius="12" Padding="10,4" Background="#13261B" BorderBrush="#1F5135" BorderThickness="1" VerticalAlignment="Center">
          <StackPanel Orientation="Horizontal">
            <TextBlock x:Name="SeloAdminIco" Text="&#xE73E;" Style="{StaticResource Icone}" FontSize="11" Foreground="#4ADE80"/>
            <TextBlock x:Name="SeloAdminTxt" Text="Administrador" FontSize="11.5" Foreground="#4ADE80" Margin="6,0,0,0"/>
          </StackPanel>
        </Border>
      </StackPanel>
    </Grid>

    <!-- Abas -->
    <Border Grid.Row="1" BorderBrush="#232833" BorderThickness="0,0,0,1" Padding="14,0">
      <StackPanel Orientation="Horizontal">
        <Button x:Name="AbaResumo" Style="{StaticResource Aba}" Tag="sel">
          <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE80F;" Style="{StaticResource Icone}" FontSize="13" Margin="0,0,8,0"/><TextBlock Text="Visão geral"/></StackPanel>
        </Button>
        <Button x:Name="AbaDiag" Style="{StaticResource Aba}">
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="&#xE9D9;" Style="{StaticResource Icone}" FontSize="13" Margin="0,0,8,0"/><TextBlock Text="Diagnóstico"/>
            <Border x:Name="SeloDiag" Background="#DC2626" CornerRadius="8" Padding="6,0" Margin="8,0,0,0" Visibility="Collapsed">
              <TextBlock x:Name="SeloDiagTxt" Text="0" FontSize="11" Foreground="White" FontWeight="SemiBold"/>
            </Border>
          </StackPanel>
        </Button>
        <Button x:Name="AbaLimpeza" Style="{StaticResource Aba}">
          <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE74D;" Style="{StaticResource Icone}" FontSize="13" Margin="0,0,8,0"/><TextBlock Text="Limpeza"/></StackPanel>
        </Button>
        <Button x:Name="AbaOtim" Style="{StaticResource Aba}">
          <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE945;" Style="{StaticResource Icone}" FontSize="13" Margin="0,0,8,0"/><TextBlock Text="Otimização"/></StackPanel>
        </Button>
        <Button x:Name="AbaReparo" Style="{StaticResource Aba}">
          <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE90F;" Style="{StaticResource Icone}" FontSize="13" Margin="0,0,8,0"/><TextBlock Text="Reparo"/></StackPanel>
        </Button>
        <Button x:Name="AbaApps" Style="{StaticResource Aba}">
          <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE719;" Style="{StaticResource Icone}" FontSize="13" Margin="0,0,8,0"/><TextBlock Text="Aplicativos"/></StackPanel>
        </Button>
        <Button x:Name="AbaAtual" Style="{StaticResource Aba}">
          <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE895;" Style="{StaticResource Icone}" FontSize="13" Margin="0,0,8,0"/><TextBlock Text="Atualizações"/></StackPanel>
        </Button>
        <Button x:Name="AbaAuto" Style="{StaticResource Aba}">
          <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE787;" Style="{StaticResource Icone}" FontSize="13" Margin="0,0,8,0"/><TextBlock Text="Automação e segurança"/></StackPanel>
        </Button>
      </StackPanel>
    </Border>

    <!-- Páginas -->
    <Grid Grid.Row="2">
      <!-- Visão geral -->
      <ScrollViewer x:Name="PgResumo" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
        <StackPanel Margin="18,16,18,16">
          <UniformGrid Columns="4">
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <StackPanel Orientation="Horizontal">
                  <TextBlock Text="&#xE916;" Style="{StaticResource Icone}" Foreground="#8B93A1" FontSize="12" Margin="0,0,7,0"/>
                  <TextBlock Text="LIGADO HÁ" Style="{StaticResource Rotulo}"/>
                </StackPanel>
                <TextBlock x:Name="ValLigado" Text="-" Style="{StaticResource Valor}"/>
                <TextBlock x:Name="SubLigado" Text="" Style="{StaticResource Sub}"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <StackPanel Orientation="Horizontal">
                  <TextBlock Text="&#xE7E8;" Style="{StaticResource Icone}" Foreground="#8B93A1" FontSize="12" Margin="0,0,7,0"/>
                  <TextBlock Text="ÚLTIMO DESLIGAMENTO" Style="{StaticResource Rotulo}"/>
                </StackPanel>
                <TextBlock x:Name="ValDesl" Text="-" Style="{StaticResource Valor}"/>
                <TextBlock x:Name="SubDesl" Text="" Style="{StaticResource Sub}"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <StackPanel Orientation="Horizontal">
                  <TextBlock Text="&#xE7F4;" Style="{StaticResource Icone}" Foreground="#8B93A1" FontSize="12" Margin="0,0,7,0"/>
                  <TextBlock Text="MEMÓRIA RAM" Style="{StaticResource Rotulo}"/>
                </StackPanel>
                <TextBlock x:Name="ValRam" Text="-" Style="{StaticResource Valor}"/>
                <TextBlock x:Name="SubRam" Text="" Style="{StaticResource Sub}"/>
                <ProgressBar x:Name="BarRam" Style="{StaticResource Barra}" Value="0"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <StackPanel Orientation="Horizontal">
                  <TextBlock Text="&#xEDA2;" Style="{StaticResource Icone}" Foreground="#8B93A1" FontSize="12" Margin="0,0,7,0"/>
                  <TextBlock x:Name="RotDisco" Text="DISCO DO SISTEMA" Style="{StaticResource Rotulo}"/>
                </StackPanel>
                <TextBlock x:Name="ValDisco" Text="-" Style="{StaticResource Valor}"/>
                <TextBlock x:Name="SubDisco" Text="" Style="{StaticResource Sub}"/>
                <ProgressBar x:Name="BarDisco" Style="{StaticResource Barra}" Value="0"/>
              </StackPanel>
            </Border>
          </UniformGrid>

          <Border x:Name="Aviso" Background="#2A2111" BorderBrush="#5C4413" BorderThickness="1" CornerRadius="10" Padding="16,12" Margin="6" Visibility="Collapsed">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="&#xE7BA;" Style="{StaticResource Icone}" Foreground="#FBBF24" FontSize="16" VerticalAlignment="Top" Margin="0,1,12,0"/>
              <TextBlock x:Name="AvisoTxt" Foreground="#FDE68A" TextWrapping="Wrap" MaxWidth="1000"/>
            </StackPanel>
          </Border>

          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="3*"/>
              <ColumnDefinition Width="2*"/>
            </Grid.ColumnDefinitions>
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Text="Quem mais usa memória agora" Style="{StaticResource TituloCard}"/>
                <StackPanel x:Name="ListaRam"/>
              </StackPanel>
            </Border>
            <Border Grid.Column="1" Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Text="Ações rápidas" Style="{StaticResource TituloCard}"/>
                <TextBlock Text="Comece pela análise: ela lê o histórico do Windows e aponta o problema, a evidência e o que fazer." Style="{StaticResource Sub}" Margin="0,0,0,14"/>
                <Button x:Name="BtnAnalisar" Style="{StaticResource Primario}" Margin="0,0,0,8" Padding="16,11" HorizontalContentAlignment="Stretch">
                  <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE721;" Style="{StaticResource Icone}" Margin="0,0,10,0"/><TextBlock Text="Analisar problemas"/></StackPanel>
                </Button>
                <Button x:Name="BtnLimpezaRapida" Style="{StaticResource Secundario}" Margin="0,0,0,8" Padding="16,10">
                  <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE74D;" Style="{StaticResource Icone}" Margin="0,0,10,0"/><TextBlock Text="Limpeza segura"/></StackPanel>
                </Button>
                <Button x:Name="BtnRelatorio" Style="{StaticResource Secundario}" Margin="0,0,0,12" Padding="16,10">
                  <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE8A5;" Style="{StaticResource Icone}" Margin="0,0,10,0"/><TextBlock Text="Gerar relatório HTML"/></StackPanel>
                </Button>
                <Button x:Name="BtnAbrirPasta" Style="{StaticResource Link}" HorizontalAlignment="Center">
                  <TextBlock Text="Abrir pasta de relatórios"/>
                </Button>
              </StackPanel>
            </Border>
          </Grid>

          <Border Style="{StaticResource Card}">
            <StackPanel>
              <TextBlock Text="Ficha da máquina" Style="{StaticResource TituloCard}"/>
              <UniformGrid x:Name="Ficha" Columns="3"/>
            </StackPanel>
          </Border>
        </StackPanel>
      </ScrollViewer>

      <!-- Diagnóstico -->
      <Grid x:Name="PgDiag" Visibility="Collapsed">
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
        <Grid Margin="24,18,24,8">
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
          <StackPanel>
            <TextBlock Text="Análise de problemas" Style="{StaticResource TituloPagina}"/>
            <TextBlock x:Name="TxtDiagSub" Text="Causa provável, evidência e o que fazer, com base no estado da máquina e no histórico do Windows." Style="{StaticResource Sub}" FontSize="12.5" Margin="0,4,0,0"/>
            <WrapPanel x:Name="Pilulas" Margin="0,12,0,0"/>
          </StackPanel>
          <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Top">
            <Button x:Name="BtnReanalisar" Style="{StaticResource Secundario}">
              <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE72C;" Style="{StaticResource Icone}" Margin="0,0,8,0"/><TextBlock Text="Analisar"/></StackPanel>
            </Button>
            <Button x:Name="BtnRelatorio2" Style="{StaticResource Secundario}">
              <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE8A5;" Style="{StaticResource Icone}" Margin="0,0,8,0"/><TextBlock Text="Relatório"/></StackPanel>
            </Button>
            <Button x:Name="BtnCorrigir" Style="{StaticResource Primario}" IsEnabled="False">
              <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE945;" Style="{StaticResource Icone}" Margin="0,0,8,0"/><TextBlock x:Name="BtnCorrigirTxt" Text="Aplicar correções"/></StackPanel>
            </Button>
          </StackPanel>
        </Grid>
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <StackPanel x:Name="ListaAchados" Margin="24,4,24,20"/>
        </ScrollViewer>
      </Grid>

      <!-- Abas de opções (montadas pelo código) -->
      <Grid x:Name="PgLimpeza" Visibility="Collapsed"/>
      <Grid x:Name="PgOtim" Visibility="Collapsed"/>
      <Grid x:Name="PgReparo" Visibility="Collapsed"/>
      <Grid x:Name="PgAtual" Visibility="Collapsed"/>
      <Grid x:Name="PgApps" Visibility="Collapsed"/>

      <!-- Automação e segurança -->
      <ScrollViewer x:Name="PgAuto" Visibility="Collapsed" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
        <Grid Margin="18,16,18,16">
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
          <Border Style="{StaticResource Card}" VerticalAlignment="Top">
            <StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                <TextBlock Text="&#xE787;" Style="{StaticResource Icone}" Foreground="#60A5FA" FontSize="15" Margin="0,0,10,0"/>
                <TextBlock Text="Modo automático semanal" FontSize="15" FontWeight="SemiBold"/>
              </StackPanel>
              <TextBlock Text="Toda semana, sem ninguém por perto: diagnóstico, limpeza segura e relatório HTML salvo numa pasta de rede. Correções nunca são aplicadas sozinhas." Style="{StaticResource Sub}" FontSize="12.5"/>
              <Border Background="#11161F" BorderBrush="#232A36" BorderThickness="1" CornerRadius="8" Padding="12,10" Margin="0,14,0,16">
                <TextBlock x:Name="TxtStatusAg" Text="Verificando agendamento..." Foreground="#CBD5E1" TextWrapping="Wrap"/>
              </Border>
              <TextBlock Text="PASTA DE REDE PARA OS RELATÓRIOS" Style="{StaticResource Rotulo}" Margin="0,0,0,6"/>
              <TextBox x:Name="TxtPasta"/>
              <TextBlock Text="Ex.: \\servidor\manutencao$  ·  em branco = só neste computador" Style="{StaticResource Sub}" Margin="0,5,0,14"/>
              <TextBlock Text="DIA DA SEMANA" Style="{StaticResource Rotulo}" Margin="0,0,0,8"/>
              <WrapPanel x:Name="PainelDias">
                <RadioButton Style="{StaticResource Chip}" GroupName="Dia" Tag="Monday" Content="Seg"/>
                <RadioButton Style="{StaticResource Chip}" GroupName="Dia" Tag="Tuesday" Content="Ter"/>
                <RadioButton Style="{StaticResource Chip}" GroupName="Dia" Tag="Wednesday" Content="Qua" IsChecked="True"/>
                <RadioButton Style="{StaticResource Chip}" GroupName="Dia" Tag="Thursday" Content="Qui"/>
                <RadioButton Style="{StaticResource Chip}" GroupName="Dia" Tag="Friday" Content="Sex"/>
                <RadioButton Style="{StaticResource Chip}" GroupName="Dia" Tag="Saturday" Content="Sáb"/>
                <RadioButton Style="{StaticResource Chip}" GroupName="Dia" Tag="Sunday" Content="Dom"/>
              </WrapPanel>
              <TextBlock Text="HORÁRIO" Style="{StaticResource Rotulo}" Margin="0,8,0,6"/>
              <TextBox x:Name="TxtHora" Text="12:00" Width="90" HorizontalAlignment="Left"/>
              <StackPanel Orientation="Horizontal" Margin="0,18,0,0">
                <Button x:Name="BtnAgendar" Style="{StaticResource Primario}" Margin="0">
                  <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE74E;" Style="{StaticResource Icone}" Margin="0,0,8,0"/><TextBlock Text="Salvar agendamento"/></StackPanel>
                </Button>
                <Button x:Name="BtnRemoverAg" Style="{StaticResource Secundario}" Content="Remover"/>
              </StackPanel>
              <TextBlock Style="{StaticResource Sub}" Margin="0,14,0,0" Text="A tarefa roda como SYSTEM: quem grava na pasta de rede é a conta do computador. Dê permissão de gravação ao grupo &quot;Computadores do domínio&quot; no compartilhamento."/>
            </StackPanel>
          </Border>
          <Border Grid.Column="1" Style="{StaticResource Card}" VerticalAlignment="Top">
            <StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                <TextBlock Text="&#xE7A7;" Style="{StaticResource Icone}" Foreground="#60A5FA" FontSize="15" Margin="0,0,10,0"/>
                <TextBlock Text="Desfazer alterações" FontSize="15" FontWeight="SemiBold"/>
              </StackPanel>
              <TextBlock Text="Cada configuração alterada pelo speedPC tem o valor anterior guardado. Escolha uma sessão para reverter tudo o que foi feito nela." Style="{StaticResource Sub}" FontSize="12.5"/>
              <ListBox x:Name="LstSessoes" Background="Transparent" BorderThickness="0" Margin="0,14,0,0" MaxHeight="230"
                       ItemContainerStyle="{StaticResource ItemLista}" ScrollViewer.HorizontalScrollBarVisibility="Disabled"/>
              <TextBlock x:Name="TxtSemSessoes" Text="Nenhuma alteração registrada neste computador." Style="{StaticResource Sub}" Margin="0,4,0,0"/>
              <TextBlock x:Name="TxtDetSessao" Style="{StaticResource Sub}" Margin="0,8,0,0"/>
              <StackPanel Orientation="Horizontal" Margin="0,16,0,0">
                <Button x:Name="BtnDesfazer" Style="{StaticResource Primario}" Margin="0" IsEnabled="False">
                  <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE7A7;" Style="{StaticResource Icone}" Margin="0,0,8,0"/><TextBlock Text="Desfazer sessão"/></StackPanel>
                </Button>
                <Button x:Name="BtnPonto" Style="{StaticResource Secundario}" Content="Criar ponto de restauração"/>
                <Button x:Name="BtnRstrui" Style="{StaticResource Secundario}" Content="Restauração do Sistema"/>
              </StackPanel>
            </StackPanel>
          </Border>
        </Grid>
      </ScrollViewer>
    </Grid>

    <!-- Registro de atividade -->
    <Border Grid.Row="3" Background="#0B0D11" BorderBrush="#232833" BorderThickness="0,1,0,0">
      <Grid>
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <Grid Margin="24,8,24,8">
          <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="&#xE756;" Style="{StaticResource Icone}" Foreground="#8B93A1" FontSize="12" Margin="0,0,8,0"/>
            <TextBlock Text="Registro de atividade" FontSize="12" FontWeight="SemiBold" Foreground="#C9CED6" VerticalAlignment="Center"/>
          </StackPanel>
          <StackPanel Grid.Column="1" Orientation="Horizontal" Margin="16,0,0,0">
            <ProgressBar x:Name="PrgOcupado" Width="110" Height="4" IsIndeterminate="True" Foreground="#3B82F6" Background="#232833" BorderThickness="0" Visibility="Collapsed" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <TextBlock x:Name="TxtStatus" Text="Pronto" FontSize="12" Foreground="#8B93A1" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
          </StackPanel>
          <Button x:Name="BtnLog" Grid.Column="2" Style="{StaticResource Link}">
            <TextBlock x:Name="BtnLogTxt" Text="Ocultar" FontSize="12"/>
          </Button>
        </Grid>
        <RichTextBox x:Name="Log" Grid.Row="1" Height="150" IsReadOnly="True" Background="#0B0D11" Foreground="#CBD5E1" BorderThickness="0"
                     Padding="20,0,20,10" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" IsDocumentEnabled="False">
          <FlowDocument PagePadding="0" FontFamily="Cascadia Mono, Consolas" FontSize="12"/>
        </RichTextBox>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

# --- Opções das abas (estilo lista de marcação, com os recomendados já marcados) ---
$Script:DefAbas = [ordered]@{
    Limpeza = @{
        Titulo = 'Limpeza segura'
        Desc   = 'Remove temporários e caches de todos os usuários e do Windows. Nunca toca em senhas, histórico, favoritos ou arquivos pessoais.'
        Botao  = 'Iniciar limpeza'
        Grupos = @(
            @{ Nome = 'Sempre incluído'; Itens = @(
                    @{ Id = 'L:Usuarios'; Nome = 'Temporários e caches de todos os usuários'; Desc = 'Pasta Temp (só arquivos com mais de 2 dias), cache dos navegadores, Teams clássico, miniaturas e relatórios de erro.'; Fixo = $true }
                    @{ Id = 'L:Sistema'; Nome = 'Temporários do Windows'; Desc = 'C:\Windows\Temp, downloads antigos do Windows Update (preservados se houver reinício pendente) e cache da Otimização de Entrega.'; Fixo = $true }
                ) }
            @{ Nome = 'Opcional'; Itens = @(
                    @{ Id = 'L:Lixeira'; Nome = 'Esvaziar a Lixeira de todos os usuários'; Desc = 'A Lixeira contém arquivos que os usuários apagaram: combine com o usuário antes.' }
                    @{ Id = 'L:WinSxS'; Nome = 'Limpar componentes do Windows (WinSxS)'; Desc = 'Remove versões antigas de componentes do sistema. Pode levar mais de 10 minutos.' }
                ) }
            @{ Nome = 'Nunca é tocado'; Info = @('Senhas, cookies, histórico e favoritos', 'Downloads, Documentos e Área de Trabalho', 'Anexos temporários do Outlook', 'Despejos de tela azul (evidência para análise)') }
        )
    }
    Otim = @{
        Titulo = 'Otimização para o trabalho'
        Desc   = 'Ajustes seguros para desempenho e estabilidade. Tudo o que altera configuração pode ser revertido em Automação e segurança > Desfazer.'
        Botao  = 'Executar selecionados'
        Grupos = @(
            @{ Nome = 'Essenciais (recomendado)'; Itens = @(
                    @{ Id = 'O:Plano'; Nome = 'Plano de energia adequado'; Desc = 'Equilibrado em notebooks; Alto desempenho em desktops.'; Rec = $true }
                    @{ Id = 'O:InicioRapido'; Nome = 'Desativar Inicialização Rápida'; Desc = 'Faz o "Desligar" reiniciar o Windows de verdade. Evita lentidão acumulada e travamentos ao desligar.'; Rec = $true }
                    @{ Id = 'O:Sugestoes'; Nome = 'Desativar sugestões e anúncios do Windows'; Desc = 'Dicas, aplicativos patrocinados e sugestões do menu Iniciar.'; Rec = $true }
                    @{ Id = 'O:Unidades'; Nome = 'Otimizar unidades'; Desc = 'TRIM no SSD e desfragmentação no HD, com o otimizador do próprio Windows.'; Rec = $true }
                ) }
            @{ Nome = 'Opcionais'; Itens = @(
                    @{ Id = 'O:StorageSense'; Nome = 'Ativar Sensor de Armazenamento'; Desc = 'O Windows passa a apagar temporários sozinho. Lixeira e Downloads ficam de fora.' }
                    @{ Id = 'O:Inicializacao'; Nome = 'Revisar programas de inicialização'; Desc = 'Abre a lista de programas que iniciam com o Windows para você escolher o que desativar.' }
                    @{ Id = 'O:Ultimate'; Nome = 'Plano Desempenho Máximo'; Desc = 'Plano oculto do Windows para estações de trabalho: elimina microatrasos de economia de energia. Só para desktops.'; Cuidado = $true }
                ) }
        )
    }
    Reparo = @{
        Titulo = 'Reparo do Windows'
        Desc   = 'Ferramentas de reparo do próprio Windows. As mais demoradas abrem uma janela própria para você acompanhar o progresso.'
        Botao  = 'Executar selecionados'
        Grupos = @(
            @{ Nome = 'Sistema'; Itens = @(
                    @{ Id = 'R:ReparoWindows'; Nome = 'Reparar arquivos do Windows (DISM + SFC)'; Desc = 'Verifica e corrige arquivos do sistema corrompidos. De 10 a 40 minutos.' }
                    @{ Id = 'R:VerificarDisco'; Nome = 'Verificar o disco (CHKDSK online)'; Desc = 'Analisa o sistema de arquivos sem reiniciar; oferece agendar a correção se precisar.' }
                    @{ Id = 'R:ReparoWU'; Nome = 'Reparar o Windows Update'; Desc = 'Recria as pastas de cache do Windows Update (as antigas ficam guardadas como .bak).' }
                ) }
            @{ Nome = 'Rede'; Itens = @(
                    @{ Id = 'R:RenovarRede'; Nome = 'Renovar IP e limpar cache DNS'; Desc = 'ipconfig /renew, /flushdns e /registerdns.' }
                    @{ Id = 'R:ResetRede'; Nome = 'Redefinir a pilha de rede'; Desc = 'Winsock e TCP/IP. Exige reiniciar e pode afetar VPNs antigas.'; Cuidado = $true }
                ) }
            @{ Nome = 'Impressão e hardware'; Itens = @(
                    @{ Id = 'R:LimparFila'; Nome = 'Destravar fila de impressão'; Desc = 'Reinicia o spooler e remove trabalhos presos (precisarão ser reimpressos).' }
                    @{ Id = 'R:TesteMemoria'; Nome = 'Agendar teste de memória RAM'; Desc = 'Abre o Diagnóstico de Memória do Windows; o teste roda no próximo reinício.' }
                ) }
            @{ Nome = 'Painéis clássicos do Windows'; Atalhos = @(
                    @{ Nome = 'Painel de Controle'; Cmd = 'control.exe' }
                    @{ Nome = 'Conexões de rede'; Cmd = 'ncpa.cpl' }
                    @{ Nome = 'Opções de energia'; Cmd = 'powercfg.cpl' }
                    @{ Nome = 'Som'; Cmd = 'mmsys.cpl' }
                    @{ Nome = 'Propriedades do sistema'; Cmd = 'sysdm.cpl' }
                    @{ Nome = 'Contas de usuário'; Cmd = 'netplwiz.exe' }
                    @{ Nome = 'Gerenciador de dispositivos'; Cmd = 'devmgmt.msc' }
                    @{ Nome = 'Visualizador de eventos'; Cmd = 'eventvwr.msc' }
                    @{ Nome = 'Impressoras'; Cmd = 'control.exe'; Args = 'printers' }
                    @{ Nome = 'Programas e recursos'; Cmd = 'appwiz.cpl' }
                ) }
        )
    }
    Atual = @{
        Titulo = 'Atualizações'
        Desc   = 'Pelo Windows Update e pelo winget, ambos da Microsoft. Antes de instalar, a lista do que foi encontrado aparece no registro e você confirma.'
        Botao  = 'Executar selecionados'
        Grupos = @(
            @{ Nome = 'Windows Update'; Itens = @(
                    @{ Id = 'A:Windows'; Nome = 'Atualizações do Windows'; Desc = 'Segurança e qualidade. Sem versões de prévia e sem troca de versão do Windows.'; Rec = $true }
                    @{ Id = 'A:Drivers'; Nome = 'Drivers'; Desc = 'Drivers publicados pela Microsoft e pelos fabricantes no Windows Update.' }
                ) }
            @{ Nome = 'Programas'; Itens = @(
                    @{ Id = 'A:Winget'; Nome = 'Programas instalados (winget)'; Desc = 'Navegadores, leitores de PDF, compactadores etc.' }
                ) }
            @{ Nome = 'Política de atualizações'; Itens = @(
                    @{ Id = 'A:PerfilEmpresa'; Nome = 'Perfil para empresas'; Desc = 'Adia novas versões do Windows por 365 dias e as atualizações de qualidade por 4 dias, para que patches com problema sejam corrigidos antes de chegar. Não use em PCs gerenciados por WSUS ou Intune.'; Cuidado = $true }
                ) }
        )
    }
}

function Get-Pincel {
    param([string]$Hex)
    if (-not $Script:G.Pinceis.ContainsKey($Hex)) {
        $b = New-Object Windows.Media.SolidColorBrush ([Windows.Media.ColorConverter]::ConvertFromString($Hex))
        $b.Freeze()
        $Script:G.Pinceis[$Hex] = $b
    }
    return $Script:G.Pinceis[$Hex]
}

function New-Espaco { param([double]$E = 0, [double]$C = 0, [double]$D = 0, [double]$B = 0) return New-Object Windows.Thickness($E, $C, $D, $B) }

function New-Texto {
    param([string]$Texto, [double]$Tamanho = 13, [string]$Cor = '#E5E7EB', [string]$Peso = 'Normal', [switch]$Quebra, $Margem = $null, [switch]$Icone)
    $t = New-Object Windows.Controls.TextBlock
    $t.Text = $Texto
    $t.FontSize = $Tamanho
    $t.Foreground = Get-Pincel $Cor
    $t.FontWeight = [Windows.FontWeights]::$Peso
    if ($Quebra) { $t.TextWrapping = 'Wrap' }
    if ($Margem) { $t.Margin = $Margem }
    if ($Icone) { $t.FontFamily = New-Object Windows.Media.FontFamily 'Segoe MDL2 Assets'; $t.VerticalAlignment = 'Center' }
    return $t
}

function Get-Estilo { param([string]$Nome) return $Script:G.Janela.FindResource($Nome) }

function Get-CorNivel {
    param([string]$Nivel)
    switch ($Nivel) { 'bom' { '#4ADE80' } 'medio' { '#FBBF24' } 'ruim' { '#F87171' } default { '#E5E7EB' } }
}

$Script:SevGui = @{
    Critico = @{ Cor = '#F87171'; Fundo = '#2A1517'; Icone = [string][char]0xEA39; Rotulo = 'CRÍTICO' }
    Atencao = @{ Cor = '#FBBF24'; Fundo = '#2A2111'; Icone = [string][char]0xE7BA; Rotulo = 'ATENÇÃO' }
    Info    = @{ Cor = '#38BDF8'; Fundo = '#0F2230'; Icone = [string][char]0xE946; Rotulo = 'INFO' }
    OK      = @{ Cor = '#4ADE80'; Fundo = '#13261B'; Icone = [string][char]0xE73E; Rotulo = 'OK' }
}

# --- Registro de atividade ---
function Add-LogLinha {
    param([string]$Texto, [string]$Cor = '#CBD5E1', [switch]$SemQuebra)
    $doc = $Script:G.Ui.Log.Document
    if (-not $Script:G.ParAtual) {
        $par = New-Object Windows.Documents.Paragraph
        $par.Margin = New-Espaco
        $doc.Blocks.Add($par)
        $Script:G.ParAtual = $par
    }
    $run = New-Object Windows.Documents.Run $Texto
    $run.Foreground = Get-Pincel $Cor
    $Script:G.ParAtual.Inlines.Add($run)
    if (-not $SemQuebra) { $Script:G.ParAtual = $null }
    while ($doc.Blocks.Count -gt 1500) { [void]$doc.Blocks.Remove($doc.Blocks.FirstBlock) }
    $Script:G.Ui.Log.ScrollToEnd()
}

function Add-LogInfo {
    param($Registro)
    $m = $Registro.MessageData
    if ($m -is [System.Management.Automation.HostInformationMessage]) {
        $cores = @{ Red = '#F87171'; DarkRed = '#F87171'; Yellow = '#FBBF24'; DarkYellow = '#FBBF24'; Green = '#4ADE80'; DarkGreen = '#4ADE80'
            Cyan = '#67E8F9'; DarkCyan = '#22D3EE'; White = '#F8FAFC'; Gray = '#CBD5E1'; DarkGray = '#8B93A1'; Blue = '#60A5FA'; Magenta = '#E879F9' }
        $cor = '#CBD5E1'
        if ($null -ne $m.ForegroundColor -and $cores.ContainsKey([string]$m.ForegroundColor)) { $cor = $cores[[string]$m.ForegroundColor] }
        $txt = [string]$m.Message
        Add-LogLinha $txt $cor -SemQuebra:([bool]$m.NoNewLine)
        if (-not $m.NoNewLine -and $txt.Trim()) { $Script:G.Ui.TxtStatus.Text = ($txt.Trim() -replace '^>\s*', '') }
    } else {
        Add-LogLinha ([string]$m)
    }
}

# --- Tarefas em segundo plano (o motor roda num runspace próprio para a janela não travar) ---
function Set-Ocupado {
    param([bool]$Sim, [string]$Texto = 'Pronto')
    $Script:G.Ui.PrgOcupado.Visibility = $(if ($Sim) { 'Visible' } else { 'Collapsed' })
    $Script:G.Ui.TxtStatus.Text = $Texto
    foreach ($b in $Script:G.BotoesAcao) { $b.IsEnabled = -not $Sim }
    if (-not $Sim) { Update-BotaoCorrigir }
}

function Start-Tarefa {
    param([string]$Comando, [string]$Status, [scriptblock]$AoTerminar, [switch]$MostrarLog)
    if ($Script:G.Tarefa) {
        Show-DialogoAviso 'Aguarde' 'Já existe uma tarefa em andamento. Aguarde ela terminar.'
        return
    }
    if ($MostrarLog) { Set-LogVisivel $true }
    $ps = [powershell]::Create()
    $ps.Runspace = $Script:G.Runspace
    [void]$ps.AddScript($Comando)
    $Script:G.Tarefa = @{ PS = $ps; Handle = $ps.BeginInvoke(); AoTerminar = $AoTerminar; Lidos = 0; ErrLidos = 0 }
    Set-Ocupado $true $Status
    Add-LogLinha ''
    Add-LogLinha ("» $Status") '#60A5FA'
}

function Read-SaidaTarefa {
    $t = $Script:G.Tarefa
    $info = $t.PS.Streams.Information
    while ($t.Lidos -lt $info.Count) { Add-LogInfo $info[$t.Lidos]; $t.Lidos++ }
    $err = $t.PS.Streams.Error
    while ($t.ErrLidos -lt $err.Count) { Add-LogLinha ('Erro: ' + $err[$t.ErrLidos].Exception.Message) '#F87171'; $t.ErrLidos++ }
}

function Invoke-Pulso {
    if ($Script:G.EmDialogo) { return }
    if ($Script:G.JobIcones) { Test-JobIcones }
    $t = $Script:G.Tarefa
    if (-not $t) { return }
    Read-SaidaTarefa
    $p = $Script:G.Ponte.Pedido
    if ($p) {
        $Script:G.Ponte.Pedido = $null
        $Script:G.EmDialogo = $true
        $resp = $null
        try { $resp = Show-Pedido $p } finally { $Script:G.EmDialogo = $false }
        $Script:G.Ponte.Resposta = $resp.Valor
    }
    if ($t.Handle.IsCompleted) {
        $saida = $null
        try { $saida = $t.PS.EndInvoke($t.Handle) } catch { Add-LogLinha ('Erro: ' + $_.Exception.InnerException.Message) '#F87171' }
        Read-SaidaTarefa
        $t.PS.Dispose()
        $Script:G.Tarefa = $null
        Set-Ocupado $false 'Pronto'
        $ultimo = $null
        if ($saida -and $saida.Count -gt 0) { $ultimo = $saida[$saida.Count - 1] }
        if ($t.AoTerminar) {
            try { & $t.AoTerminar $ultimo } catch { Add-LogLinha ('Erro na interface: ' + $_.Exception.Message) '#F87171' }
        }
    }
}

function Show-Pedido {
    param($Pedido)
    switch ($Pedido.Tipo) {
        'Confirmar' { return @{ Valor = [bool](Show-DialogoConfirmar -Texto $Pedido.Dados.Texto -Padrao $Pedido.Dados.Padrao) } }
        'Selecao' {
            $sel = @(Show-DialogoSelecao -Titulo $Pedido.Dados.Titulo -Itens $Pedido.Dados.Itens -Aviso $Pedido.Dados.Aviso)
            return @{ Valor = [int[]]$sel }
        }
    }
    return @{ Valor = $false }
}

# --- Diálogos ---
function New-JanelaDialogo {
    param([string]$Titulo, [double]$Largura = 600)
    $w = New-Object Windows.Window
    $w.Title = "speedPC - $Titulo"
    $w.Width = $Largura
    $w.SizeToContent = 'Height'
    $w.MaxHeight = 720
    $w.ResizeMode = 'NoResize'
    $w.ShowInTaskbar = $false
    $w.Background = Get-Pincel '#171A21'
    $w.Foreground = Get-Pincel '#E5E7EB'
    $w.FontFamily = New-Object Windows.Media.FontFamily 'Segoe UI'
    $w.FontSize = 13
    $w.UseLayoutRounding = $true
    $w.Resources.Add([Windows.Controls.Primitives.ScrollBar], $Script:G.Janela.FindResource([Windows.Controls.Primitives.ScrollBar]))
    if ($Script:G.Janela.IsVisible) { $w.Owner = $Script:G.Janela; $w.WindowStartupLocation = 'CenterOwner' } else { $w.WindowStartupLocation = 'CenterScreen' }
    return $w
}

function New-Botao {
    param([string]$Texto, [string]$Estilo = 'Secundario', [string]$Icone = '')
    $b = New-Object Windows.Controls.Button
    $b.Style = Get-Estilo $Estilo
    if ($Icone) {
        $sp = New-Object Windows.Controls.StackPanel; $sp.Orientation = 'Horizontal'
        [void]$sp.Children.Add((New-Texto $Icone 13 -Icone -Margem (New-Espaco 0 0 8 0) -Cor $(if ($Estilo -eq 'Primario') { '#FFFFFF' } else { '#E5E7EB' })))
        [void]$sp.Children.Add((New-Texto $Texto 13 -Cor $(if ($Estilo -eq 'Primario') { '#FFFFFF' } else { '#E5E7EB' }) -Peso $(if ($Estilo -eq 'Primario') { 'SemiBold' } else { 'Normal' })))
        $b.Content = $sp
    } else { $b.Content = $Texto }
    return $b
}

function New-ConteudoSelecao {
    # Monta o conteúdo do diálogo de escolha (também usado na captura de telas)
    param([string]$Titulo, [object[]]$Itens, [string]$Aviso)
    $raiz = New-Object Windows.Controls.Grid
    $raiz.Margin = New-Espaco 24 20 24 20
    foreach ($h in 'Auto', '*', 'Auto') { $rd = New-Object Windows.Controls.RowDefinition; $rd.Height = $h; $raiz.RowDefinitions.Add($rd) }
    $topo = New-Object Windows.Controls.StackPanel
    [void]$topo.Children.Add((New-Texto $Titulo 17 -Peso SemiBold -Quebra))
    [void]$topo.Children.Add((New-Texto 'Marque o que deve ser executado.' 12.5 '#9CA3AF' -Margem (New-Espaco 0 4 0 0)))
    if ($Aviso) {
        $av = New-Object Windows.Controls.Border
        $av.Background = Get-Pincel '#2A2111'; $av.BorderBrush = Get-Pincel '#5C4413'; $av.BorderThickness = New-Espaco 1 1 1 1
        $av.CornerRadius = New-Object Windows.CornerRadius 8; $av.Padding = New-Espaco 12 9 12 9; $av.Margin = New-Espaco 0 12 0 0
        $av.Child = New-Texto $Aviso 12.5 '#FDE68A' -Quebra
        [void]$topo.Children.Add($av)
    }
    $raiz.Children.Add($topo) | Out-Null
    $sv = New-Object Windows.Controls.ScrollViewer
    $sv.VerticalScrollBarVisibility = 'Auto'; $sv.HorizontalScrollBarVisibility = 'Disabled'; $sv.MaxHeight = 430; $sv.Margin = New-Espaco 0 16 0 8
    [Windows.Controls.Grid]::SetRow($sv, 1)
    $lista = New-Object Windows.Controls.StackPanel
    $checks = @()
    foreach ($it in @($Itens)) {
        $cb = New-Object Windows.Controls.CheckBox
        $cb.Style = Get-Estilo 'Opcao'
        $sp = New-Object Windows.Controls.StackPanel
        [void]$sp.Children.Add((New-Texto ([string]$it.Nome) 13.5 -Quebra))
        if ($it.Detalhe) { [void]$sp.Children.Add((New-Texto ([string]$it.Detalhe) 12 '#9CA3AF' -Quebra -Margem (New-Espaco 0 2 0 0))) }
        $cb.Content = $sp
        [void]$lista.Children.Add($cb)
        $checks += $cb
    }
    $sv.Content = $lista
    $raiz.Children.Add($sv) | Out-Null
    $rodape = New-Object Windows.Controls.DockPanel
    [Windows.Controls.Grid]::SetRow($rodape, 2)
    $todos = New-Object Windows.Controls.Button; $todos.Style = Get-Estilo 'Link'; $todos.Content = 'Marcar todos'
    $todos.Tag = $checks
    $todos.add_Click({ param($s, $e) $marcar = @($s.Tag | Where-Object { -not $_.IsChecked }).Count -gt 0; foreach ($c in $s.Tag) { $c.IsChecked = $marcar } })
    [void]$rodape.Children.Add($todos)
    $dir = New-Object Windows.Controls.StackPanel; $dir.Orientation = 'Horizontal'; $dir.HorizontalAlignment = 'Right'
    $cancelar = New-Botao 'Cancelar' 'Secundario'
    $ok = New-Botao 'Executar' 'Primario' ([string][char]0xE768)
    $cancelar.add_Click({ param($s, $e) [Windows.Window]::GetWindow($s).Close() })
    $ok.add_Click({ param($s, $e) $w = [Windows.Window]::GetWindow($s); $w.Tag = 'ok'; $w.Close() })
    [void]$dir.Children.Add($cancelar); [void]$dir.Children.Add($ok)
    [void]$rodape.Children.Add($dir)
    $raiz.Children.Add($rodape) | Out-Null
    return @{ Raiz = $raiz; Checks = $checks }
}

function Show-DialogoSelecao {
    param([string]$Titulo, [object[]]$Itens, [string]$Aviso)
    $w = New-JanelaDialogo $Titulo 640
    $c = New-ConteudoSelecao -Titulo $Titulo -Itens $Itens -Aviso $Aviso
    $w.Content = $c.Raiz
    [void]$w.ShowDialog()
    if ($w.Tag -ne 'ok') { return }
    for ($i = 0; $i -lt $c.Checks.Count; $i++) { if ($c.Checks[$i].IsChecked) { $i } }
}

function Show-DialogoConfirmar {
    param([string]$Texto, [string]$Padrao = 'N', [string]$Titulo = 'Confirmar', [string]$Sim = 'Sim', [string]$Nao = 'Não')
    $w = New-JanelaDialogo $Titulo 500
    $sp = New-Object Windows.Controls.StackPanel; $sp.Margin = New-Espaco 24 22 24 20
    $linha = New-Object Windows.Controls.StackPanel; $linha.Orientation = 'Horizontal'
    [void]$linha.Children.Add((New-Texto ([string][char]0xE9CE) 22 '#60A5FA' -Icone -Margem (New-Espaco 0 0 14 0)))
    $tx = New-Texto $Texto 13.5 -Quebra; $tx.MaxWidth = 390; $tx.VerticalAlignment = 'Center'
    [void]$linha.Children.Add($tx)
    [void]$sp.Children.Add($linha)
    $bt = New-Object Windows.Controls.StackPanel; $bt.Orientation = 'Horizontal'; $bt.HorizontalAlignment = 'Right'; $bt.Margin = New-Espaco 0 22 0 0
    $bNao = New-Botao $Nao 'Secundario'; $bSim = New-Botao $Sim 'Primario'
    $bNao.add_Click({ param($s, $e) [Windows.Window]::GetWindow($s).Close() })
    $bSim.add_Click({ param($s, $e) $j = [Windows.Window]::GetWindow($s); $j.Tag = 'ok'; $j.Close() })
    if ($Padrao -eq 'S') { $bSim.IsDefault = $true } else { $bNao.IsDefault = $true }
    [void]$bt.Children.Add($bNao); [void]$bt.Children.Add($bSim)
    [void]$sp.Children.Add($bt)
    $w.Content = $sp
    [void]$w.ShowDialog()
    return ($w.Tag -eq 'ok')
}

function Show-DialogoAviso {
    param([string]$Titulo, [string]$Texto)
    $w = New-JanelaDialogo $Titulo 460
    $sp = New-Object Windows.Controls.StackPanel; $sp.Margin = New-Espaco 24 22 24 20
    [void]$sp.Children.Add((New-Texto $Texto 13.5 -Quebra))
    $b = New-Botao 'OK' 'Primario'; $b.HorizontalAlignment = 'Right'; $b.Margin = New-Espaco 0 20 0 0; $b.IsDefault = $true
    $b.add_Click({ param($s, $e) [Windows.Window]::GetWindow($s).Close() })
    [void]$sp.Children.Add($b)
    $w.Content = $sp
    [void]$w.ShowDialog()
}

# --- Navegação ---
function Set-Aba {
    param([string]$Nome)
    foreach ($n in $Script:G.Abas) {
        $Script:G.Ui["Pg$n"].Visibility = $(if ($n -eq $Nome) { 'Visible' } else { 'Collapsed' })
        $Script:G.Ui["Aba$n"].Tag = $(if ($n -eq $Nome) { 'sel' } else { $null })
    }
    if ($Nome -eq 'Auto' -and -not $Script:G.Tarefa -and $Script:G.MotorPronto) { Update-Automacao }
    if ($Nome -eq 'Apps' -and -not $Script:G.AppsLidos -and -not $Script:G.Tarefa -and $Script:G.MotorPronto) { Update-AppsInstalados }
}

function Set-LogVisivel {
    param([bool]$Visivel)
    $Script:G.Ui.Log.Visibility = $(if ($Visivel) { 'Visible' } else { 'Collapsed' })
    $Script:G.Ui.BtnLogTxt.Text = $(if ($Visivel) { 'Ocultar' } else { 'Mostrar' })
}

# --- Páginas de opções ---
function New-PaginaOpcoes {
    param([string]$Chave)
    $def = $Script:DefAbas[$Chave]
    $pg = $Script:G.Ui["Pg$Chave"]
    $pg.Children.Clear()
    foreach ($h in 'Auto', '*') { $rd = New-Object Windows.Controls.RowDefinition; $rd.Height = $h; $pg.RowDefinitions.Add($rd) }

    $topo = New-Object Windows.Controls.Grid; $topo.Margin = New-Espaco 24 18 24 6
    foreach ($w in '*', 'Auto') { $cd = New-Object Windows.Controls.ColumnDefinition; $cd.Width = $w; $topo.ColumnDefinitions.Add($cd) }
    $esq = New-Object Windows.Controls.StackPanel
    $t = New-Texto $def.Titulo 20 -Peso SemiBold; [void]$esq.Children.Add($t)
    $d = New-Texto $def.Desc 12.5 '#9CA3AF' -Quebra -Margem (New-Espaco 0 4 24 0); $d.MaxWidth = 760; [void]$esq.Children.Add($d)
    [void]$topo.Children.Add($esq)
    $dir = New-Object Windows.Controls.StackPanel; $dir.Orientation = 'Horizontal'; $dir.VerticalAlignment = 'Top'
    [Windows.Controls.Grid]::SetColumn($dir, 1)
    $checks = New-Object System.Collections.Generic.List[object]
    if ($Chave -ne 'Limpeza') {
        $bRec = New-Botao 'Recomendados' 'Secundario'; $bRec.Tag = $Chave
        $bRec.add_Click({ param($s, $e) foreach ($c in $Script:G.Checks[$s.Tag]) { if ($c.IsEnabled) { $c.IsChecked = [bool]$c.Uid } } })
        $bNada = New-Botao 'Limpar seleção' 'Secundario'; $bNada.Tag = $Chave
        $bNada.add_Click({ param($s, $e) foreach ($c in $Script:G.Checks[$s.Tag]) { if ($c.IsEnabled) { $c.IsChecked = $false } } })
        [void]$dir.Children.Add($bRec); [void]$dir.Children.Add($bNada)
    }
    $bExec = New-Botao $def.Botao 'Primario' ([string][char]0xE768); $bExec.Tag = $Chave
    $bExec.add_Click({ param($s, $e) Invoke-ExecutarAba $s.Tag })
    [void]$dir.Children.Add($bExec)
    $Script:G.BotoesAcao += $bExec
    [void]$topo.Children.Add($dir)
    [void]$pg.Children.Add($topo)

    $sv = New-Object Windows.Controls.ScrollViewer; $sv.VerticalScrollBarVisibility = 'Auto'; $sv.HorizontalScrollBarVisibility = 'Disabled'
    [Windows.Controls.Grid]::SetRow($sv, 1)
    $pilha = New-Object Windows.Controls.StackPanel; $pilha.Margin = New-Espaco 18 8 18 16
    $grade = New-Object Windows.Controls.Primitives.UniformGrid
    $grade.Columns = [math]::Min(3, @($def.Grupos | Where-Object { -not $_.Atalhos }).Count); $grade.VerticalAlignment = 'Top'
    [void]$pilha.Children.Add($grade)
    foreach ($g in $def.Grupos) {
        $card = New-Object Windows.Controls.Border; $card.Style = Get-Estilo 'Card'; $card.VerticalAlignment = 'Stretch'
        $sp = New-Object Windows.Controls.StackPanel
        [void]$sp.Children.Add((New-Texto $g.Nome 14 -Peso SemiBold -Margem (New-Espaco 0 0 0 14)))
        if ($g.Info) {
            foreach ($linha in $g.Info) {
                $l = New-Object Windows.Controls.StackPanel; $l.Orientation = 'Horizontal'; $l.Margin = New-Espaco 0 0 0 10
                [void]$l.Children.Add((New-Texto ([string][char]0xE72E) 12 '#4ADE80' -Icone -Margem (New-Espaco 0 0 10 0)))
                [void]$l.Children.Add((New-Texto $linha 13 '#C9CED6'))
                [void]$sp.Children.Add($l)
            }
        }
        if ($g.Atalhos) {
            [void]$sp.Children.Add((New-Texto 'Abrem na hora, sem alterar nada.' 12 '#9CA3AF' -Margem (New-Espaco 0 -8 0 12)))
            $wp = New-Object Windows.Controls.WrapPanel
            foreach ($a in $g.Atalhos) {
                $b = New-Botao $a.Nome 'Secundario' ([string][char]0xE8A7)
                $b.Margin = New-Espaco 0 0 8 8; $b.Padding = New-Espaco 12 6 12 6; $b.Tag = $a
                $b.add_Click({ param($s, $e)
                        try { if ($s.Tag.Args) { Start-Process $s.Tag.Cmd $s.Tag.Args } else { Start-Process $s.Tag.Cmd } } catch { Show-DialogoAviso 'Atalho' $_.Exception.Message } })
                [void]$wp.Children.Add($b)
            }
            [void]$sp.Children.Add($wp)
        }
        foreach ($it in @($g.Itens)) {
            if (-not $it) { continue }
            $cb = New-Object Windows.Controls.CheckBox; $cb.Style = Get-Estilo 'Opcao'; $cb.Tag = $it.Id
            $cb.Uid = $(if ($it.Rec) { '1' } else { '' })
            $cb.IsChecked = [bool]($it.Rec -or $it.Fixo)
            if ($it.Fixo) { $cb.IsEnabled = $false }
            $conteudo = New-Object Windows.Controls.StackPanel
            $linhaNome = New-Object Windows.Controls.WrapPanel
            [void]$linhaNome.Children.Add((New-Texto $it.Nome 13.5))
            if ($it.Cuidado) {
                $tag = New-Object Windows.Controls.Border; $tag.Background = Get-Pincel '#2A2111'; $tag.CornerRadius = New-Object Windows.CornerRadius 4
                $tag.Padding = New-Espaco 6 1 6 1; $tag.Margin = New-Espaco 8 1 0 0; $tag.Child = New-Texto 'CUIDADO' 10 '#FBBF24' -Peso SemiBold
                [void]$linhaNome.Children.Add($tag)
            }
            [void]$conteudo.Children.Add($linhaNome)
            [void]$conteudo.Children.Add((New-Texto $it.Desc 12 '#9CA3AF' -Quebra -Margem (New-Espaco 0 3 0 0)))
            $cb.Content = $conteudo
            [void]$sp.Children.Add($cb)
            $checks.Add($cb)
        }
        $card.Child = $sp
        if ($g.Atalhos) { [void]$pilha.Children.Add($card) } else { [void]$grade.Children.Add($card) }
    }
    $sv.Content = $pilha
    [void]$pg.Children.Add($sv)
    $Script:G.Checks[$Chave] = $checks
}

function Invoke-ExecutarAba {
    param([string]$Chave)
    $def = $Script:DefAbas[$Chave]
    $ids = @($Script:G.Checks[$Chave] | Where-Object { $_.IsChecked -and $_.IsEnabled } | ForEach-Object { [string]$_.Tag })
    if (-not $Script:G.Estado.EhAdmin) { Show-DialogoAviso 'Administrador necessário' 'Abra o speedPC como administrador para executar esta ação.'; return }
    if ($Chave -eq 'Limpeza') {
        $lix = '$false'; if ($ids -contains 'L:Lixeira') { $lix = '$true' }
        $sxs = '$false'; if ($ids -contains 'L:WinSxS') { $sxs = '$true' }
        $cmd = "Invoke-LimpezaSegura -Silencioso -IncluirLixeira:$lix -Profunda:$sxs; Get-Resumo | Out-Null; Get-EstadoGui"
    } else {
        if (-not $ids) { Show-DialogoAviso $def.Titulo 'Marque pelo menos uma opção.'; return }
        $lista = ($ids | ForEach-Object { "'$_'" }) -join ','
        $cmd = "Invoke-OpcoesGui -Ids @($lista); Get-Resumo | Out-Null; Get-EstadoGui"
    }
    Start-Tarefa -Comando $cmd -Status "$($def.Titulo) em andamento" -MostrarLog -AoTerminar { param($s) Update-Tudo $s }
}

# --- Aplicativos (instalação pelo winget, estilo lista de marcação) ---
function New-PaginaApps {
    $pg = $Script:G.Ui.PgApps
    $pg.Children.Clear()
    foreach ($h in 'Auto', '*') { $rd = New-Object Windows.Controls.RowDefinition; $rd.Height = $h; $pg.RowDefinitions.Add($rd) }
    $topo = New-Object Windows.Controls.Grid; $topo.Margin = New-Espaco 24 18 24 6
    foreach ($w in '*', 'Auto') { $cd = New-Object Windows.Controls.ColumnDefinition; $cd.Width = $w; $topo.ColumnDefinitions.Add($cd) }
    $esq = New-Object Windows.Controls.StackPanel
    [void]$esq.Children.Add((New-Texto 'Aplicativos' 20 -Peso SemiBold))
    $d = New-Texto 'Instala, atualiza ou remove programas pelo winget, o gerenciador de pacotes oficial do Windows: sem procurar sites de download e sem instaladores com propaganda.' 12.5 '#9CA3AF' -Quebra -Margem (New-Espaco 0 4 24 0)
    $d.MaxWidth = 700; [void]$esq.Children.Add($d)
    # Busca
    $busca = New-Object Windows.Controls.Grid; $busca.Margin = New-Espaco 0 12 0 0; $busca.Width = 340; $busca.HorizontalAlignment = 'Left'
    $tb = New-Object Windows.Controls.TextBox; $tb.Padding = New-Espaco 30 0 6 0
    $lupa = New-Texto ([string][char]0xE721) 12 '#8B93A1' -Icone -Margem (New-Espaco 11 0 0 0); $lupa.IsHitTestVisible = $false
    $dica = New-Texto 'Buscar aplicativo...' 13 '#6B7280' -Margem (New-Espaco 32 0 0 0); $dica.VerticalAlignment = 'Center'; $dica.IsHitTestVisible = $false
    [void]$busca.Children.Add($tb); [void]$busca.Children.Add($lupa); [void]$busca.Children.Add($dica)
    $Script:G.AppsDica = $dica
    $tb.add_TextChanged({ param($s, $e) Set-FiltroApps $s.Text })
    [void]$esq.Children.Add($busca)
    [void]$topo.Children.Add($esq)

    $dir = New-Object Windows.Controls.StackPanel; $dir.Orientation = 'Horizontal'; $dir.VerticalAlignment = 'Top'
    [Windows.Controls.Grid]::SetColumn($dir, 1)
    $bLer = New-Botao 'Ver instalados' 'Secundario' ([string][char]0xE72C)
    $bLer.add_Click({ Update-AppsInstalados })
    $bLimpar = New-Botao 'Limpar seleção' 'Secundario'
    $bLimpar.add_Click({ foreach ($c in $Script:G.AppChecks) { $c.IsChecked = $false } })
    $bRem = New-Botao 'Desinstalar' 'Secundario' ([string][char]0xE74D)
    $bRem.add_Click({ Invoke-AppsSelecionados -Desinstalar })
    $bInst = New-Botao 'Instalar / atualizar' 'Primario' ([string][char]0xE896)
    $bInst.add_Click({ Invoke-AppsSelecionados })
    $Script:G.AppsBotaoTxt = ($bInst.Content.Children[1])
    foreach ($b in $bLer, $bLimpar, $bRem, $bInst) { [void]$dir.Children.Add($b) }
    $Script:G.BotoesAcao += @($bLer, $bRem, $bInst)
    [void]$topo.Children.Add($dir)
    [void]$pg.Children.Add($topo)

    $sv = New-Object Windows.Controls.ScrollViewer; $sv.VerticalScrollBarVisibility = 'Auto'; $sv.HorizontalScrollBarVisibility = 'Disabled'
    [Windows.Controls.Grid]::SetRow($sv, 1)
    $grade = New-Object Windows.Controls.Primitives.UniformGrid; $grade.Columns = 3; $grade.Margin = New-Espaco 18 8 18 16; $grade.VerticalAlignment = 'Top'
    $Script:G.AppChecks = New-Object System.Collections.Generic.List[object]
    $Script:G.AppSelos = @{}
    $Script:G.AppCards = @()
    foreach ($cat in $Script:CatalogoApps) {
        $card = New-Object Windows.Controls.Border; $card.Style = Get-Estilo 'Card'
        $sp = New-Object Windows.Controls.StackPanel
        [void]$sp.Children.Add((New-Texto $cat.Cat 14 -Peso SemiBold -Margem (New-Espaco 0 0 0 14)))
        $doCard = @()
        foreach ($a in $cat.Apps) {
            $cb = New-Object Windows.Controls.CheckBox; $cb.Style = Get-Estilo 'Opcao'; $cb.Tag = $a.Id; $cb.Margin = New-Espaco 0 0 0 10
            $cb.Uid = ("{0} {1} {2}" -f $a.Nome, $a.Desc, $a.Id).ToLower()
            $grid = New-Object Windows.Controls.Grid
            foreach ($w in '40', '*') { $cd = New-Object Windows.Controls.ColumnDefinition; $cd.Width = $w; $grid.ColumnDefinitions.Add($cd) }
            [void]$grid.Children.Add((New-IconeApp $a))
            $conteudo = New-Object Windows.Controls.StackPanel; $conteudo.VerticalAlignment = 'Center'
            [Windows.Controls.Grid]::SetColumn($conteudo, 1)
            [void]$grid.Children.Add($conteudo)
            $linha = New-Object Windows.Controls.WrapPanel
            [void]$linha.Children.Add((New-Texto $a.Nome 13.5))
            $selo = New-Object Windows.Controls.Border; $selo.Background = Get-Pincel '#13261B'; $selo.CornerRadius = New-Object Windows.CornerRadius 4
            $selo.Padding = New-Espaco 6 1 6 1; $selo.Margin = New-Espaco 8 1 0 0; $selo.Visibility = 'Collapsed'
            $selo.Child = New-Texto 'INSTALADO' 10 '#4ADE80' -Peso SemiBold
            [void]$linha.Children.Add($selo)
            [void]$conteudo.Children.Add($linha)
            [void]$conteudo.Children.Add((New-Texto $a.Desc 12 '#9CA3AF' -Quebra -Margem (New-Espaco 0 2 0 0)))
            $cb.Content = $grid
            $cb.add_Checked({ Update-ContagemApps })
            $cb.add_Unchecked({ Update-ContagemApps })
            [void]$sp.Children.Add($cb)
            $Script:G.AppChecks.Add($cb)
            $Script:G.AppSelos[$a.Id] = $selo
            $doCard += $cb
        }
        $card.Child = $sp
        $card.Tag = $doCard
        $Script:G.AppCards += $card
        [void]$grade.Children.Add($card)
    }
    $sv.Content = $grade
    [void]$pg.Children.Add($sv)
}

# Ícones dos aplicativos: inicial colorida na hora; o ícone real é baixado uma vez (em segundo plano)
# e guardado em %LOCALAPPDATA%\speedPC\Icones. Sem internet, a inicial continua.
$Script:PastaIcones = Join-Path $env:LOCALAPPDATA 'speedPC\Icones'

function Get-UrlIcone {
    param($App)
    if ($App.Icone) { return $App.Icone }
    if ($App.Site) { return 'https://www.google.com/s2/favicons?sz=64&domain_url=' + [uri]::EscapeDataString($App.Site) }
    return $null
}

function Get-ArquivoIcone { param([string]$Id) return Join-Path $Script:PastaIcones (($Id -replace '[^\w.-]', '_') + '.png') }

function New-IconeApp {
    param($App)
    $caixa = New-Object Windows.Controls.Grid; $caixa.Width = 28; $caixa.Height = 28; $caixa.HorizontalAlignment = 'Left'; $caixa.VerticalAlignment = 'Top'; $caixa.Margin = New-Espaco 0 1 0 0
    $paleta = '#2563EB', '#7C3AED', '#DB2777', '#EA580C', '#16A34A', '#0891B2', '#CA8A04', '#DC2626'
    $h = 0; foreach ($ch in $App.Nome.ToCharArray()) { $h = ($h * 31 + [int]$ch) % 9973 }
    $fundo = New-Object Windows.Controls.Border; $fundo.CornerRadius = New-Object Windows.CornerRadius 7; $fundo.Background = Get-Pincel $paleta[$h % $paleta.Count]
    $palavras = @($App.Nome -split '[\s\-\.]+' | Where-Object { $_ })
    $ini = if ($palavras.Count -ge 2) { "$($palavras[0][0])$($palavras[1][0])" } else { $App.Nome.Substring(0, [math]::Min(2, $App.Nome.Length)) }
    $t = New-Texto $ini.ToUpper() 11 '#FFFFFF' -Peso SemiBold; $t.HorizontalAlignment = 'Center'; $t.VerticalAlignment = 'Center'
    $fundo.Child = $t
    [void]$caixa.Children.Add($fundo)
    $img = New-Object Windows.Controls.Image; $img.Width = 28; $img.Height = 28
    [Windows.Media.RenderOptions]::SetBitmapScalingMode($img, 'HighQuality')
    [void]$caixa.Children.Add($img)
    $Script:G.AppIcones[$App.Id] = @{ Imagem = $img; Fundo = $fundo }
    Set-IconeApp $App.Id
    return $caixa
}

function Set-IconeApp {
    param([string]$Id)
    $alvo = $Script:G.AppIcones[$Id]; if (-not $alvo) { return }
    $arq = Get-ArquivoIcone $Id
    if (-not (Test-Path -LiteralPath $arq) -or (Get-Item -LiteralPath $arq).Length -lt 100) { return }
    try {
        $bi = New-Object Windows.Media.Imaging.BitmapImage
        $bi.BeginInit(); $bi.UriSource = New-Object Uri $arq; $bi.CacheOption = 'OnLoad'; $bi.DecodePixelWidth = 56; $bi.EndInit(); $bi.Freeze()
        $alvo.Imagem.Source = $bi
        $alvo.Fundo.Visibility = 'Hidden'
    } catch { }
}

function Start-DownloadIcones {
    param([switch]$Esperar)
    $pendentes = @()
    foreach ($c in $Script:CatalogoApps) {
        foreach ($a in $c.Apps) {
            $url = Get-UrlIcone $a; if (-not $url) { continue }
            $arq = Get-ArquivoIcone $a.Id
            $marca = "$arq.falhou"
            if (Test-Path -LiteralPath $arq) { continue }
            if ((Test-Path -LiteralPath $marca) -and (Get-Item -LiteralPath $marca).LastWriteTime -gt (Get-Date).AddDays(-7)) { continue }
            $pendentes += [pscustomobject]@{ Id = $a.Id; Url = $url; Arquivo = $arq }
        }
    }
    if (-not $pendentes) { return }
    $script = {
        param($Lista, $Pasta)
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
        if (-not (Test-Path -LiteralPath $Pasta)) { New-Item -ItemType Directory -Path $Pasta -Force | Out-Null }
        $wc = New-Object Net.WebClient
        try { $wc.Proxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials } catch { }
        $ok = @()
        foreach ($i in $Lista) {
            try {
                $tmp = "$($i.Arquivo).tmp"
                $wc.DownloadFile($i.Url, $tmp)
                Move-Item -LiteralPath $tmp -Destination $i.Arquivo -Force
                $ok += $i.Id
            } catch {
                Remove-Item -LiteralPath "$($i.Arquivo).tmp" -Force -ErrorAction SilentlyContinue
                Set-Content -LiteralPath "$($i.Arquivo).falhou" -Value '' -ErrorAction SilentlyContinue
            }
        }
        $ok
    }
    $ps = [powershell]::Create()
    [void]$ps.AddScript($script).AddArgument($pendentes).AddArgument($Script:PastaIcones)
    if ($Esperar) {
        foreach ($id in $ps.Invoke()) { Set-IconeApp $id }
        $ps.Dispose()
        return
    }
    $Script:G.JobIcones = @{ PS = $ps; Handle = $ps.BeginInvoke() }
}

function Test-JobIcones {
    $j = $Script:G.JobIcones
    if (-not $j -or -not $j.Handle.IsCompleted) { return }
    $Script:G.JobIcones = $null
    try { foreach ($id in $j.PS.EndInvoke($j.Handle)) { Set-IconeApp $id } } catch { }
    $j.PS.Dispose()
}

function Set-FiltroApps {
    param([string]$Texto)
    $t = $Texto.Trim().ToLower()
    $Script:G.AppsDica.Visibility = $(if ($Texto) { 'Collapsed' } else { 'Visible' })
    foreach ($card in $Script:G.AppCards) {
        $visiveis = 0
        foreach ($cb in $card.Tag) {
            $mostra = (-not $t) -or $cb.Uid.Contains($t)
            $cb.Visibility = $(if ($mostra) { 'Visible' } else { 'Collapsed' })
            if ($mostra) { $visiveis++ }
        }
        $card.Visibility = $(if ($visiveis) { 'Visible' } else { 'Collapsed' })
    }
}

function Update-ContagemApps {
    $n = @($Script:G.AppChecks | Where-Object { $_.IsChecked }).Count
    $Script:G.AppsBotaoTxt.Text = $(if ($n) { "Instalar / atualizar ($n)" } else { 'Instalar / atualizar' })
}

function Show-AppsInstalados {
    param($Ids)
    $set = @{}; foreach ($i in @($Ids)) { if ($i) { $set[$i.ToLower()] = $true } }
    $n = 0
    foreach ($id in @($Script:G.AppSelos.Keys)) {
        $tem = $set.ContainsKey($id.ToLower())
        $Script:G.AppSelos[$id].Visibility = $(if ($tem) { 'Visible' } else { 'Collapsed' })
        if ($tem) { $n++ }
    }
    $Script:G.AppsLidos = $true
    Add-LogLinha "$n aplicativo(s) do catálogo já instalado(s) nesta máquina." '#8B93A1'
}

function Update-AppsInstalados {
    Start-Tarefa -Comando '@(Get-AppsInstalados) -join "|"' -Status 'Verificando aplicativos instalados' -AoTerminar {
        param($s) Show-AppsInstalados (([string]$s) -split '\|') }
}

function Invoke-AppsSelecionados {
    param([switch]$Desinstalar)
    $sel = @($Script:G.AppChecks | Where-Object { $_.IsChecked })
    if (-not $sel) { Show-DialogoAviso 'Aplicativos' 'Marque pelo menos um aplicativo.'; return }
    $nomes = @($sel | ForEach-Object { $_.Content.Children[1].Children[0].Children[0].Text })
    $verbo = 'instalar ou atualizar'; if ($Desinstalar) { $verbo = 'DESINSTALAR' }
    $texto = "Vai $verbo $($sel.Count) aplicativo(s): $($nomes -join ', ')."
    if (-not $Desinstalar) { $texto += ' Os instaladores rodam em modo silencioso; alguns podem pedir confirmação do Windows.' }
    $rotuloSim = 'Instalar'; if ($Desinstalar) { $rotuloSim = 'Desinstalar' }
    if (-not (Show-DialogoConfirmar -Titulo 'Aplicativos' -Texto $texto -Sim $rotuloSim)) { return }
    $lista = ($sel | ForEach-Object { Get-TextoPs ([string]$_.Tag) }) -join ','
    $extra = ''; if ($Desinstalar) { $extra = ' -Desinstalar' }
    $status = 'Instalando aplicativos'; if ($Desinstalar) { $status = 'Desinstalando aplicativos' }
    Start-Tarefa -Comando "Install-Apps -Ids @($lista)$extra; @(Get-AppsInstalados) -join '|'" -Status $status -MostrarLog -AoTerminar {
        param($s)
        Show-AppsInstalados (([string]$s) -split '\|')
        foreach ($c in $Script:G.AppChecks) { $c.IsChecked = $false }
    }
}

# --- Visão geral ---
function Update-Resumo {
    $e = $Script:G.Estado; if (-not $e -or -not $e.Resumo) { return }
    $r = $e.Resumo; $ui = $Script:G.Ui
    $ui.TxtPc.Text = $r.Computador
    $ui.TxtSub.Text = ('{0}  ·  {1} {2}' -f ($r.Windows -replace '^Microsoft ', ''), $r.Fabricante, $r.Modelo)
    if ($e.EhAdmin) {
        $ui.SeloAdmin.Background = Get-Pincel '#13261B'; $ui.SeloAdmin.BorderBrush = Get-Pincel '#1F5135'
        $ui.SeloAdminIco.Text = [string][char]0xE73E; $ui.SeloAdminIco.Foreground = Get-Pincel '#4ADE80'
        $ui.SeloAdminTxt.Text = 'Administrador'; $ui.SeloAdminTxt.Foreground = Get-Pincel '#4ADE80'
    } else {
        $ui.SeloAdmin.Background = Get-Pincel '#2A2111'; $ui.SeloAdmin.BorderBrush = Get-Pincel '#5C4413'
        $ui.SeloAdminIco.Text = [string][char]0xE7BA; $ui.SeloAdminIco.Foreground = Get-Pincel '#FBBF24'
        $ui.SeloAdminTxt.Text = 'Sem administrador'; $ui.SeloAdminTxt.Foreground = Get-Pincel '#FBBF24'
    }
    $n = 'bom'; if ($r.TempoLigado.TotalDays -ge 7) { $n = 'ruim' } elseif ($r.TempoLigado.TotalDays -ge 3) { $n = 'medio' }
    $ui.ValLigado.Text = Format-Duracao $r.TempoLigado; $ui.ValLigado.Foreground = Get-Pincel (Get-CorNivel $n)
    $ui.SubLigado.Text = 'desde ' + $r.Boot.ToString('dd/MM/yyyy HH:mm')
    $d = $r.Desligamento
    switch ($d.Tipo) {
        'Normal' { $ui.ValDesl.Text = 'Normal'; $ui.ValDesl.Foreground = Get-Pincel (Get-CorNivel 'bom') }
        'Inesperado' { $ui.ValDesl.Text = 'Inesperado'; $ui.ValDesl.Foreground = Get-Pincel (Get-CorNivel 'ruim') }
        default { $ui.ValDesl.Text = 'Sem registro'; $ui.ValDesl.Foreground = Get-Pincel '#E5E7EB' }
    }
    $ui.SubDesl.Text = ($d.Texto -replace '^(Normal|INESPERADO) - ', '')
    if ($d.Observacao) { $ui.Aviso.Visibility = 'Visible'; $ui.AvisoTxt.Text = $d.Observacao } else { $ui.Aviso.Visibility = 'Collapsed' }
    $n = 'bom'; if ($r.RamPct -ge 90) { $n = 'ruim' } elseif ($r.RamPct -ge 80) { $n = 'medio' }
    $ui.ValRam.Text = "$($r.RamPct)%"; $ui.ValRam.Foreground = Get-Pincel (Get-CorNivel $n)
    $ui.SubRam.Text = '{0} de {1}' -f (Format-Bytes $r.RamUsada), (Format-Bytes $r.RamTotal)
    $ui.BarRam.Value = $r.RamPct; $ui.BarRam.Foreground = Get-Pincel (Get-CorNivel $n)
    if ($r.DiscoSistema -and $r.DiscoSistema.Size) {
        $uso = [math]::Round(($r.DiscoSistema.Size - $r.DiscoSistema.FreeSpace) / $r.DiscoSistema.Size * 100)
        $n = 'bom'; if ($uso -gt 90) { $n = 'ruim' } elseif ($uso -gt 85) { $n = 'medio' }
        $ui.RotDisco.Text = "DISCO $env:SystemDrive"
        $ui.ValDisco.Text = (Format-Bytes $r.DiscoSistema.FreeSpace) + ' livres'; $ui.ValDisco.Foreground = Get-Pincel (Get-CorNivel $n)
        $ui.SubDisco.Text = "$uso% usado de $(Format-Bytes $r.DiscoSistema.Size)"
        $ui.BarDisco.Value = $uso; $ui.BarDisco.Foreground = Get-Pincel (Get-CorNivel $n)
    }
    # Maiores consumidores de RAM
    $ui.ListaRam.Children.Clear()
    $tops = @($r.TopMemoria | Select-Object -First 6)
    $max = [double]($tops | Measure-Object Bytes -Maximum).Maximum; if ($max -le 0) { $max = 1 }
    foreach ($p in $tops) {
        $g = New-Object Windows.Controls.Grid; $g.Margin = New-Espaco 0 0 0 12
        foreach ($w in '*', 'Auto') { $cd = New-Object Windows.Controls.ColumnDefinition; $cd.Width = $w; $g.ColumnDefinitions.Add($cd) }
        $nome = $p.Nome; if ($p.Instancias -gt 1) { $nome += "  ·  $($p.Instancias) processos" }
        $tn = New-Texto $nome 13; $tn.TextTrimming = 'CharacterEllipsis'
        [void]$g.Children.Add($tn)
        $tv = New-Texto (Format-Bytes $p.Bytes) 13 '#C9CED6' -Peso SemiBold; [Windows.Controls.Grid]::SetColumn($tv, 1); [void]$g.Children.Add($tv)
        $pb = New-Object Windows.Controls.ProgressBar; $pb.Style = Get-Estilo 'Barra'; $pb.Height = 4; $pb.Margin = New-Espaco 0 22 0 0
        $pb.Value = [math]::Round($p.Bytes / $max * 100); $pb.Foreground = Get-Pincel '#6366F1'
        [Windows.Controls.Grid]::SetColumnSpan($pb, 2); [void]$g.Children.Add($pb)
        [void]$ui.ListaRam.Children.Add($g)
    }
    # Ficha
    $ui.Ficha.Children.Clear()
    $ficha = [ordered]@{
        'Fabricante e modelo' = "$($r.Fabricante) $($r.Modelo)"; 'Número de série' = $r.Serie; 'Windows' = ($r.Windows -replace '^Microsoft ', '')
        'Processador' = "$($r.Processador) ($($r.Nucleos) threads)"; 'BIOS' = $r.Bios; 'Domínio' = $r.Dominio
        'Usuário logado' = $r.UsuarioLogado; 'Endereços IP' = $r.IPs; 'Inicialização Rápida' = $(if ($r.InicioRapido) { 'Ativa' } else { 'Desativada' })
    }
    foreach ($k in $ficha.Keys) {
        $sp = New-Object Windows.Controls.StackPanel; $sp.Margin = New-Espaco 0 0 18 14
        [void]$sp.Children.Add((New-Texto $k.ToUpper() 10.5 '#8B93A1' -Peso SemiBold))
        $v = [string]$ficha[$k]; if (-not $v.Trim()) { $v = '-' }
        [void]$sp.Children.Add((New-Texto $v 13 -Quebra -Margem (New-Espaco 0 3 0 0)))
        [void]$ui.Ficha.Children.Add($sp)
    }
}

# --- Diagnóstico ---
function New-Pilula {
    param([string]$Texto, [string]$Cor, [string]$Fundo)
    $b = New-Object Windows.Controls.Border
    $b.Background = Get-Pincel $Fundo; $b.CornerRadius = New-Object Windows.CornerRadius 12
    $b.Padding = New-Espaco 10 3 10 3; $b.Margin = New-Espaco 0 0 8 0
    $b.Child = New-Texto $Texto 12 $Cor -Peso SemiBold
    return $b
}

function New-CartaoAchado {
    param($A)
    $s = $Script:SevGui[$A.Severidade]
    $card = New-Object Windows.Controls.Border
    $card.Background = Get-Pincel '#171A21'; $card.BorderBrush = Get-Pincel '#262B35'; $card.BorderThickness = New-Espaco 1 1 1 1
    $card.CornerRadius = New-Object Windows.CornerRadius 10; $card.Margin = New-Espaco 0 0 0 10
    $g = New-Object Windows.Controls.Grid
    foreach ($w in '4', '*') { $cd = New-Object Windows.Controls.ColumnDefinition; $cd.Width = $w; $g.ColumnDefinitions.Add($cd) }
    $faixa = New-Object Windows.Controls.Border; $faixa.Background = Get-Pincel $s.Cor; $faixa.CornerRadius = New-Object Windows.CornerRadius 10, 0, 0, 10
    [void]$g.Children.Add($faixa)
    $sp = New-Object Windows.Controls.StackPanel; $sp.Margin = New-Espaco 16 14 18 14
    [Windows.Controls.Grid]::SetColumn($sp, 1)
    $topo = New-Object Windows.Controls.DockPanel
    $pil = New-Object Windows.Controls.Border; $pil.Background = Get-Pincel $s.Fundo; $pil.CornerRadius = New-Object Windows.CornerRadius 5
    $pil.Padding = New-Espaco 7 2 8 2; $pil.Margin = New-Espaco 0 0 12 0; $pil.VerticalAlignment = 'Top'
    $pl = New-Object Windows.Controls.StackPanel; $pl.Orientation = 'Horizontal'
    [void]$pl.Children.Add((New-Texto $s.Icone 10.5 $s.Cor -Icone -Margem (New-Espaco 0 0 5 0)))
    [void]$pl.Children.Add((New-Texto $s.Rotulo 11 $s.Cor -Peso SemiBold))
    $pil.Child = $pl
    [void]$topo.Children.Add($pil)
    [void]$topo.Children.Add((New-Texto $A.Titulo 14.5 -Peso SemiBold -Quebra))
    [void]$sp.Children.Add($topo)
    foreach ($par in @(@('Evidência', $A.Evidencia), @('O que fazer', $A.Recomendacao))) {
        if (-not $par[1]) { continue }
        $tb = New-Object Windows.Controls.TextBlock; $tb.TextWrapping = 'Wrap'; $tb.Margin = New-Espaco 0 9 0 0; $tb.FontSize = 13
        $r1 = New-Object Windows.Documents.Run ($par[0] + '   '); $r1.Foreground = Get-Pincel '#8B93A1'; $r1.FontWeight = [Windows.FontWeights]::SemiBold
        $r2 = New-Object Windows.Documents.Run ([string]$par[1]); $r2.Foreground = Get-Pincel '#CBD5E1'
        $tb.Inlines.Add($r1); $tb.Inlines.Add($r2)
        [void]$sp.Children.Add($tb)
    }
    if ($A.Correcoes) {
        $cb = New-Object Windows.Controls.Border; $cb.Background = Get-Pincel '#152238'; $cb.CornerRadius = New-Object Windows.CornerRadius 6
        $cb.Padding = New-Espaco 9 5 10 5; $cb.Margin = New-Espaco 0 12 0 0; $cb.HorizontalAlignment = 'Left'
        $cl = New-Object Windows.Controls.StackPanel; $cl.Orientation = 'Horizontal'
        [void]$cl.Children.Add((New-Texto ([string][char]0xE945) 11 '#60A5FA' -Icone -Margem (New-Espaco 0 0 7 0)))
        $ct = New-Texto ('Correção automática: ' + $A.Correcoes) 12 '#93C5FD' -Quebra; $ct.MaxWidth = 900
        [void]$cl.Children.Add($ct)
        $cb.Child = $cl
        [void]$sp.Children.Add($cb)
    }
    [void]$g.Children.Add($sp)
    $card.Child = $g
    return $card
}

function Show-DiagnosticoEmAndamento {
    $l = $Script:G.Ui.ListaAchados; $l.Children.Clear()
    $card = New-Object Windows.Controls.Border; $card.Style = Get-Estilo 'Card'; $card.Margin = New-Espaco 0 0 0 0; $card.Padding = New-Espaco 28 28 28 28
    $sp = New-Object Windows.Controls.StackPanel; $sp.HorizontalAlignment = 'Center'
    [void]$sp.Children.Add((New-Texto ([string][char]0xE9D9) 30 '#60A5FA' -Icone -Margem (New-Espaco 0 0 0 12)))
    $t = New-Texto 'Analisando o computador...' 16 -Peso SemiBold; $t.HorizontalAlignment = 'Center'; [void]$sp.Children.Add($t)
    $s = New-Texto 'Lendo o histórico do Windows, memória, discos, rede, câmera e impressão. Leva de 30 a 60 segundos.' 12.5 '#9CA3AF' -Quebra -Margem (New-Espaco 0 6 0 14)
    $s.TextAlignment = 'Center'; $s.MaxWidth = 460; [void]$sp.Children.Add($s)
    $pb = New-Object Windows.Controls.ProgressBar; $pb.IsIndeterminate = $true; $pb.Height = 4; $pb.Width = 260
    $pb.Foreground = Get-Pincel '#3B82F6'; $pb.Background = Get-Pincel '#262B35'; $pb.BorderThickness = New-Espaco 0 0 0 0
    [void]$sp.Children.Add($pb)
    $card.Child = $sp
    [void]$l.Children.Add($card)
}

function Update-BotaoCorrigir {
    $e = $Script:G.Estado
    $n = 0; if ($e) { $n = [int]$e.Pendentes }
    $Script:G.Ui.BtnCorrigirTxt.Text = $(if ($n -gt 0) { "Aplicar correções ($n)" } else { 'Aplicar correções' })
    $Script:G.Ui.BtnCorrigir.IsEnabled = ($n -gt 0 -and $e.EhAdmin -and -not $Script:G.Tarefa)
}

function Update-Diagnostico {
    $e = $Script:G.Estado; $ui = $Script:G.Ui
    $ui.ListaAchados.Children.Clear(); $ui.Pilulas.Children.Clear()
    Update-BotaoCorrigir
    if (-not $e -or -not $e.DiagnosticoFeito) {
        $card = New-Object Windows.Controls.Border; $card.Style = Get-Estilo 'Card'; $card.Margin = New-Espaco 0 0 0 0; $card.Padding = New-Espaco 28 30 28 30
        $sp = New-Object Windows.Controls.StackPanel; $sp.HorizontalAlignment = 'Center'
        $i = New-Texto ([string][char]0xE721) 30 '#60A5FA' -Icone -Margem (New-Espaco 0 0 0 12); $i.HorizontalAlignment = 'Center'; [void]$sp.Children.Add($i)
        $t = New-Texto 'Nenhuma análise feita ainda' 16 -Peso SemiBold; $t.HorizontalAlignment = 'Center'; [void]$sp.Children.Add($t)
        $s = New-Texto 'A análise aponta lentidão, travamentos, problemas de câmera e microfone, rede e impressão, com a evidência e a correção possível.' 12.5 '#9CA3AF' -Quebra -Margem (New-Espaco 0 6 0 16)
        $s.TextAlignment = 'Center'; $s.MaxWidth = 520; [void]$sp.Children.Add($s)
        $b = New-Botao 'Analisar agora' 'Primario' ([string][char]0xE721); $b.HorizontalAlignment = 'Center'; $b.Margin = New-Espaco 0 0 0 0
        $b.add_Click({ Start-Analise })
        $Script:G.BotoesAcao += $b
        [void]$sp.Children.Add($b)
        $card.Child = $sp
        [void]$ui.ListaAchados.Children.Add($card)
        $ui.SeloDiag.Visibility = 'Collapsed'
        return
    }
    $prob = @($e.Achados | Where-Object { $_.Severidade -ne 'OK' })
    $ok = @($e.Achados | Where-Object { $_.Severidade -eq 'OK' })
    $nC = @($prob | Where-Object { $_.Severidade -eq 'Critico' }).Count
    $nA = @($prob | Where-Object { $_.Severidade -eq 'Atencao' }).Count
    $nI = @($prob | Where-Object { $_.Severidade -eq 'Info' }).Count
    [void]$ui.Pilulas.Children.Add((New-Pilula "$nC crítico(s)" '#F87171' '#2A1517'))
    [void]$ui.Pilulas.Children.Add((New-Pilula "$nA atenção" '#FBBF24' '#2A2111'))
    [void]$ui.Pilulas.Children.Add((New-Pilula "$nI informativo(s)" '#38BDF8' '#0F2230'))
    [void]$ui.Pilulas.Children.Add((New-Pilula "$($ok.Count) OK" '#4ADE80' '#13261B'))
    $ui.SeloDiag.Visibility = $(if (($nC + $nA) -gt 0) { 'Visible' } else { 'Collapsed' })
    $ui.SeloDiagTxt.Text = [string]($nC + $nA)
    if (-not $prob) {
        $card = New-Object Windows.Controls.Border; $card.Style = Get-Estilo 'Card'; $card.Margin = New-Espaco 0 0 0 10
        $l = New-Object Windows.Controls.StackPanel; $l.Orientation = 'Horizontal'
        [void]$l.Children.Add((New-Texto ([string][char]0xE73E) 18 '#4ADE80' -Icone -Margem (New-Espaco 0 0 12 0)))
        [void]$l.Children.Add((New-Texto 'Nenhum problema encontrado nesta máquina.' 14 -Peso SemiBold))
        $card.Child = $l
        [void]$ui.ListaAchados.Children.Add($card)
    }
    $cat = ''
    foreach ($a in $prob) {
        if ($a.Categoria -ne $cat) {
            $cat = $a.Categoria
            [void]$ui.ListaAchados.Children.Add((New-Texto $cat.ToUpper() 11.5 '#8B93A1' -Peso SemiBold -Margem (New-Espaco 2 10 0 8)))
        }
        [void]$ui.ListaAchados.Children.Add((New-CartaoAchado $a))
    }
    if ($ok) {
        [void]$ui.ListaAchados.Children.Add((New-Texto "VERIFICAÇÕES SEM PROBLEMA ($($ok.Count))" 11.5 '#8B93A1' -Peso SemiBold -Margem (New-Espaco 2 14 0 8)))
        $wp = New-Object Windows.Controls.WrapPanel
        foreach ($a in $ok) {
            $chip = New-Object Windows.Controls.Border; $chip.Background = Get-Pincel '#13201A'; $chip.BorderBrush = Get-Pincel '#1F3B2B'
            $chip.BorderThickness = New-Espaco 1 1 1 1; $chip.CornerRadius = New-Object Windows.CornerRadius 14; $chip.Padding = New-Espaco 10 4 12 4; $chip.Margin = New-Espaco 0 0 8 8
            $cl = New-Object Windows.Controls.StackPanel; $cl.Orientation = 'Horizontal'
            [void]$cl.Children.Add((New-Texto ([string][char]0xE73E) 10 '#4ADE80' -Icone -Margem (New-Espaco 0 0 7 0)))
            [void]$cl.Children.Add((New-Texto $a.Titulo 12 '#B7E4C7'))
            $chip.Child = $cl
            [void]$wp.Children.Add($chip)
        }
        [void]$ui.ListaAchados.Children.Add($wp)
    }
}

function Update-Tudo {
    param($Estado)
    if ($Estado -and $Estado.PSObject.Properties['Resumo']) { $Script:G.Estado = $Estado }
    Update-Resumo
    Update-Diagnostico
}

function Start-Analise {
    Set-Aba 'Diag'
    Show-DiagnosticoEmAndamento
    Start-Tarefa -Comando 'Invoke-Diagnostico -Silencioso; Get-EstadoGui' -Status 'Analisando o computador' -AoTerminar { param($s) Update-Tudo $s }
}

function Start-Relatorio {
    Start-Tarefa -Comando '$a = New-RelatorioHtml; [pscustomobject]@{ Arquivo = $a; Estado = (Get-EstadoGui) }' -Status 'Gerando relatório HTML' -AoTerminar {
        param($s)
        if ($s -and $s.Estado) { Update-Tudo $s.Estado }
        if ($s -and $s.Arquivo -and (Test-Path -LiteralPath $s.Arquivo)) { Add-LogLinha "Relatório salvo em $($s.Arquivo)" '#4ADE80'; Start-Process $s.Arquivo }
    }
}

# --- Automação e segurança ---
function Update-Automacao {
    Start-Tarefa -Comando '[pscustomobject]@{ Agendamento = (Get-StatusAgendamento); Sessoes = @(Get-SessoesBackup) }' -Status 'Lendo agendamento e backups' -AoTerminar {
        param($s)
        if ($s) { Show-Automacao $s }
    }
}

function Show-Automacao {
    param($Dados)
    $ui = $Script:G.Ui
    $ag = $Dados.Agendamento
    if ($ag) {
        $ui.TxtStatusAg.Text = $ag.Texto
        if ($ag.Existe) {
            $ui.TxtPasta.Text = $ag.Pasta; $ui.TxtHora.Text = $ag.Hora
            foreach ($rb in $ui.PainelDias.Children) { $rb.IsChecked = ([string]$rb.Tag -eq $ag.Dia) }
        }
        $ui.BtnRemoverAg.IsEnabled = [bool]$ag.Existe
    }
    $ui.LstSessoes.Items.Clear()
    $sessoes = @($Dados.Sessoes)
    foreach ($s in $sessoes) {
        $it = New-Object Windows.Controls.ListBoxItem
        $it.Tag = $s
        $sp = New-Object Windows.Controls.StackPanel
        [void]$sp.Children.Add((New-Texto ('{0:dd/MM/yyyy HH:mm}  ·  {1} alteração(ões)' -f $s.Data, $s.Qtd) 13 -Peso SemiBold))
        [void]$sp.Children.Add((New-Texto ("por $($s.Executor)") 11.5 '#9CA3AF' -Margem (New-Espaco 0 2 0 0)))
        $it.Content = $sp
        [void]$ui.LstSessoes.Items.Add($it)
    }
    $ui.TxtSemSessoes.Visibility = $(if ($sessoes) { 'Collapsed' } else { 'Visible' })
    $ui.LstSessoes.Visibility = $(if ($sessoes) { 'Visible' } else { 'Collapsed' })
    $ui.TxtDetSessao.Text = ''
    $ui.BtnDesfazer.IsEnabled = $false
}

function Get-TextoPs { param([string]$Texto) return "'" + ($Texto -replace "'", "''") + "'" }

# --- Inicialização da janela ---
function Start-Gui {
    param([string]$Captura)
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $Script:G = @{
        Pinceis = @{}; Ui = @{}; Checks = @{}; BotoesAcao = @(); Tarefa = $null; EmDialogo = $false; Estado = $null; ParAtual = $null; MotorPronto = $false
        Abas = @('Resumo', 'Diag', 'Limpeza', 'Otim', 'Reparo', 'Apps', 'Atual', 'Auto'); AppsLidos = $false; AppIcones = @{}; JobIcones = $null
        Ponte = [hashtable]::Synchronized(@{ Pedido = $null; Resposta = $null })
    }
    $janela = [Windows.Markup.XamlReader]::Parse($Script:XamlJanela)
    $Script:G.Janela = $janela
    foreach ($m in [regex]::Matches($Script:XamlJanela, 'x:Name="(\w+)"')) {
        $nome = $m.Groups[1].Value
        $el = $janela.FindName($nome)
        if ($el) { $Script:G.Ui[$nome] = $el }
    }
    $janela.Title = "speedPC $Script:Versao"
    $ui = $Script:G.Ui
    foreach ($k in 'Limpeza', 'Otim', 'Reparo', 'Atual') { New-PaginaOpcoes $k }
    New-PaginaApps
    $Script:G.BotoesAcao += @($ui.BtnAnalisar, $ui.BtnLimpezaRapida, $ui.BtnRelatorio, $ui.BtnReanalisar, $ui.BtnRelatorio2, $ui.BtnAgendar, $ui.BtnRemoverAg, $ui.BtnPonto)

    # Eventos
    foreach ($n in $Script:G.Abas) {
        $ui["Aba$n"].Uid = $n
        $ui["Aba$n"].add_Click({ param($s, $e) Set-Aba $s.Uid })
    }
    $ui.BtnAnalisar.add_Click({ Start-Analise })
    $ui.BtnReanalisar.add_Click({ Start-Analise })
    $ui.BtnLimpezaRapida.add_Click({ Set-Aba 'Limpeza' })
    $ui.BtnRelatorio.add_Click({ Start-Relatorio })
    $ui.BtnRelatorio2.add_Click({ Start-Relatorio })
    $ui.BtnAbrirPasta.add_Click({
            $p = $Script:PastaRelLocal
            if (-not (Test-Path -LiteralPath $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
            Start-Process explorer.exe $p })
    $ui.BtnCorrigir.add_Click({
            Start-Tarefa -Comando 'Invoke-MenuCorrecoes; Get-EstadoGui' -Status 'Aplicando correções' -MostrarLog -AoTerminar { param($s) Update-Tudo $s } })
    $ui.BtnLog.add_Click({ Set-LogVisivel ($Script:G.Ui.Log.Visibility -ne 'Visible') })
    $ui.LstSessoes.add_SelectionChanged({
            $it = $Script:G.Ui.LstSessoes.SelectedItem
            if ($it) {
                $Script:G.Ui.TxtDetSessao.Text = (@($it.Tag.Itens | Select-Object -First 12 | ForEach-Object { "•  $_" }) -join "`n")
                $Script:G.Ui.BtnDesfazer.IsEnabled = (-not $Script:G.Tarefa)
            } })
    $ui.BtnDesfazer.add_Click({
            $it = $Script:G.Ui.LstSessoes.SelectedItem
            if (-not $it) { return }
            $s = $it.Tag
            if (-not (Show-DialogoConfirmar -Titulo 'Desfazer alterações' -Texto ("Reverter as {0} alterações feitas em {1:dd/MM/yyyy HH:mm}? Algumas só valem depois de reiniciar ou fazer logoff." -f $s.Qtd, $s.Data) -Sim 'Desfazer')) { return }
            Start-Tarefa -Comando ("Undo-SessaoBackup -Arquivo {0}; [pscustomobject]@{{ Agendamento = (Get-StatusAgendamento); Sessoes = @(Get-SessoesBackup) }}" -f (Get-TextoPs $s.Arquivo)) -Status 'Desfazendo alterações' -MostrarLog -AoTerminar { param($r) if ($r) { Show-Automacao $r } } })
    $ui.BtnPonto.add_Click({
            Start-Tarefa -Comando '$Script:PontoRestauracaoFeito = $false; New-PontoRestauracao' -Status 'Criando ponto de restauração' -MostrarLog })
    $ui.BtnRstrui.add_Click({ Start-Process "$env:SystemRoot\System32\rstrui.exe" })
    $ui.BtnAgendar.add_Click({
            $ui = $Script:G.Ui
            $hora = $ui.TxtHora.Text.Trim()
            if ($hora -notmatch '^([01]?\d|2[0-3]):[0-5]\d$') { Show-DialogoAviso 'Horário inválido' 'Informe o horário no formato HH:mm, por exemplo 12:30.'; return }
            $dia = 'Wednesday'; foreach ($rb in $ui.PainelDias.Children) { if ($rb.IsChecked) { $dia = [string]$rb.Tag } }
            $pasta = $ui.TxtPasta.Text.Trim().TrimEnd('\')
            $cmd = "Write-Host (Test-PastaRede {0}) -ForegroundColor Gray; Register-ModoAutomatico -Pasta {0} -Dia '{1}' -Horario '{2}'; [pscustomobject]@{{ Agendamento = (Get-StatusAgendamento); Sessoes = @(Get-SessoesBackup) }}" -f (Get-TextoPs $pasta), $dia, $hora
            Start-Tarefa -Comando $cmd -Status 'Salvando agendamento' -MostrarLog -AoTerminar { param($r) if ($r) { Show-Automacao $r } } })
    $ui.BtnRemoverAg.add_Click({
            if (-not (Show-DialogoConfirmar -Titulo 'Remover agendamento' -Texto 'Remover a manutenção semanal automática deste computador?' -Sim 'Remover')) { return }
            $cmd = "Unregister-ScheduledTask -TaskPath '\speedPC\' -TaskName `$Script:NomeTarefa -Confirm:`$false -ErrorAction SilentlyContinue; Add-Acao 'Agendamento semanal' 'Removido'; Write-Host 'Agendamento removido.' -ForegroundColor Green; [pscustomobject]@{ Agendamento = (Get-StatusAgendamento); Sessoes = @(Get-SessoesBackup) }"
            Start-Tarefa -Comando $cmd -Status 'Removendo agendamento' -AoTerminar { param($r) if ($r) { Show-Automacao $r } } })

    # Motor em segundo plano
    $iss = [initialsessionstate]::CreateDefault()
    try { $iss.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass } catch { }
    $rs = [runspacefactory]::CreateRunspace($iss)
    $rs.ApartmentState = 'STA'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('CaminhoManut', $Script:CaminhoScript)
    $rs.SessionStateProxy.SetVariable('PonteGui', $Script:G.Ponte)
    $Script:G.Runspace = $rs
    $carregar = '. $CaminhoManut -Motor; $Script:ModoGui = $true; $Script:Ponte = $PonteGui; Get-Resumo | Out-Null; Get-EstadoGui'

    Update-Diagnostico
    Add-LogLinha "speedPC $Script:Versao  ·  $env:COMPUTERNAME" '#8B93A1'

    if ($Captura) {
        # Modo de verificação visual: gera imagens das telas sem abrir a janela
        $ps = [powershell]::Create(); $ps.Runspace = $rs
        [void]$ps.AddScript($carregar.Replace('Get-EstadoGui', 'Invoke-Diagnostico -Silencioso; Get-EstadoGui'))
        $saida = $ps.Invoke()
        foreach ($rec in $ps.Streams.Information) { Add-LogInfo $rec }
        foreach ($er in $ps.Streams.Error) { Add-LogLinha ('Erro: ' + $er.Exception.Message) '#F87171' }
        $Script:G.MotorPronto = $true
        Update-Tudo $saida[$saida.Count - 1]
        $ps2 = [powershell]::Create(); $ps2.Runspace = $rs
        [void]$ps2.AddScript('[pscustomobject]@{ Agendamento = (Get-StatusAgendamento); Sessoes = @(Get-SessoesBackup) }')
        Show-Automacao ($ps2.Invoke()[0])
        Start-DownloadIcones -Esperar
        $ps3 = [powershell]::Create(); $ps3.Runspace = $rs
        [void]$ps3.AddScript('@(Get-AppsInstalados) -join "|"')
        Show-AppsInstalados (([string]($ps3.Invoke()[0])) -split '\|')
        $Script:G.AppChecks[0].IsChecked = $true; $Script:G.AppChecks[6].IsChecked = $true
        $Script:G.MotorPronto = $false   # evita disparar leituras em segundo plano durante a captura
        if (-not (Test-Path -LiteralPath $Captura)) { New-Item -ItemType Directory -Path $Captura -Force | Out-Null }
        foreach ($aba in $Script:G.Abas) {
            Set-Aba $aba
            Save-Captura $janela.Content (Join-Path $Captura "tela_$aba.png")
        }
        $dlg = New-ConteudoSelecao -Titulo 'Correções automáticas disponíveis' -Itens @(
            [pscustomobject]@{ Nome = 'Desativar Inicialização Rápida'; Detalhe = 'Faz o "Desligar" reiniciar o Windows de verdade. (Reversão: pode ser desfeita)' }
            [pscustomobject]@{ Nome = 'Verificar o disco (CHKDSK online)'; Detalhe = 'Analisa o sistema de arquivos sem precisar reiniciar. (Reversão: sem alteração de configuração)' }
            [pscustomobject]@{ Nome = 'Atualizar drivers (Windows Update)'; Detalhe = 'Procura e instala drivers publicados no Windows Update. (Reversão: não reversível)' }
        ) -Aviso 'Antes da primeira alteração é criado um ponto de restauração e um backup de cada configuração.'
        $dlg.Checks[0].IsChecked = $true
        $borda = New-Object Windows.Controls.Border; $borda.Background = Get-Pincel '#171A21'; $borda.Child = $dlg.Raiz
        $janela.Content = $null
        $janela.Content = $borda
        Save-Captura $borda (Join-Path $Captura 'dialogo_selecao.png') 640 470
        $rs.Dispose()
        return
    }

    $timer = New-Object Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(150)
    $timer.add_Tick({ Invoke-Pulso })
    $Script:G.Timer = $timer
    $janela.add_Loaded({
            $Script:G.Timer.Start()
            Hide-Console
            Start-DownloadIcones
            Start-Tarefa -Comando $carregar -Status 'Lendo informações do computador' -AoTerminar {
                param($s)
                $Script:G.MotorPronto = $true
                Update-Tudo $s
                Add-LogLinha 'Pronto. Comece por "Analisar problemas".' '#4ADE80'
            }
        })
    $janela.add_Closing({
            param($s, $e)
            if ($Script:G.Tarefa -and -not (Show-DialogoConfirmar -Titulo 'Sair' -Texto 'Há uma tarefa em andamento. Fechar mesmo assim? Ela será interrompida.' -Sim 'Fechar')) { $e.Cancel = $true }
        })
    [void]$janela.ShowDialog()
    try { $Script:G.Timer.Stop() } catch { }
    try { if ($Script:G.Tarefa) { $Script:G.Tarefa.PS.Stop() } } catch { }
    try { $rs.Dispose() } catch { }
}

function Save-Captura {
    param($Elemento, [string]$Arquivo, [double]$Largura = 1220, [double]$Altura = 800)
    $Elemento.Measure((New-Object Windows.Size($Largura, $Altura)))
    $Elemento.Arrange((New-Object Windows.Rect(0, 0, $Largura, $Altura)))
    $Elemento.UpdateLayout()
    $bmp = New-Object Windows.Media.Imaging.RenderTargetBitmap([int]$Largura, [int]$Altura, 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
    $bmp.Render($Elemento)
    $enc = New-Object Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bmp))
    $fs = [IO.File]::Create($Arquivo)
    try { $enc.Save($fs) } finally { $fs.Close() }
}

function Hide-Console {
    try {
        if (-not ('speedPC.Console' -as [type])) {
            Add-Type -Namespace speedPC -Name Console -MemberDefinition '[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow(); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);'
        }
        $h = [speedPC.Console]::GetConsoleWindow()
        if ($h -ne [IntPtr]::Zero) { [void][speedPC.Console]::ShowWindow($h, 0) }
    } catch { }
}

#endregion

#region ===================== Execução =====================

if ($Motor) { return }   # carregado pela janela gráfica: só as funções

try { $Host.UI.RawUI.WindowTitle = "speedPC $Script:Versao - $env:COMPUTERNAME" } catch { }

if ($RemoverAgendamento) {
    Unregister-ScheduledTask -TaskPath '\speedPC\' -TaskName $Script:NomeTarefa -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host 'Agendamento removido.'
    return
}
if ($Agendar) { Register-ModoAutomatico -Pasta $PastaRelatorios -Dia $DiaSemana -Horario $Hora; return }
if ($Script:ModoAuto) { Invoke-ModoAutomatico; return }
if ($SomenteRelatorio) {
    Get-Resumo | Out-Null
    Show-Resumo
    Invoke-Diagnostico
    $destino = $Script:PastaRelLocal; if ($PastaRelatorios) { $destino = $PastaRelatorios }
    $arq = New-RelatorioHtml -Pasta $destino
    Write-Host ''
    Write-Host "  Relatório: $arq" -ForegroundColor Green
    return
}

if ($CapturarTelas) { Start-Gui -Captura $CapturarTelas; return }

if (-not $Script:EhAdmin -and $PSCommandPath) {
    Write-Host 'Solicitando permissão de administrador...' -ForegroundColor Yellow
    $argumentos = @('-NoProfile', '-ExecutionPolicy', 'Bypass')
    if (-not $Console) { $argumentos += @('-WindowStyle', 'Hidden') }
    $argumentos += @('-File', "`"$PSCommandPath`"")
    if ($Console) { $argumentos += '-Console' }
    try {
        Start-Process powershell.exe -Verb RunAs -ArgumentList $argumentos -ErrorAction Stop
        return
    } catch {
        Write-Host 'Elevação recusada: continuando com recursos limitados (sem correções).' -ForegroundColor Yellow
        Start-Sleep -Seconds 2
    }
}

Write-Log "Sessão interativa iniciada por $([Security.Principal.WindowsIdentity]::GetCurrent().Name)"
if ($Console) { Show-MenuPrincipal; return }
try {
    Start-Gui
} catch {
    Write-Host "Não foi possível abrir a janela gráfica: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'Abrindo o modo texto...' -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    Show-MenuPrincipal
}

#endregion
