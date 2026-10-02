#CERTO4

#requires -Version 5.1
<#
===============================================================================
WORKFLOW DOCUMENTAL - PROTÓTIPO CSV
Compatível com Windows PowerShell 5.1 e PowerShell ISE

Persistência:
- CSV local
- Sem JSON
- Sem SQLite
- Sem DLL externa

Fluxo:
Terceiro -> Líder -> Supervisor -> Coordenador
===============================================================================
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
try {
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
}
catch { }

# =============================================================================
# CONFIGURAÇÃO
# =============================================================================

$script:AppName = "Workflow Documental"

$script:Dir = $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($script:Dir)) {
    if (
        $null -ne $psISE -and
        $null -ne $psISE.CurrentFile -and
        -not [string]::IsNullOrWhiteSpace($psISE.CurrentFile.FullPath)
    ) {
        $script:Dir = Split-Path -Parent $psISE.CurrentFile.FullPath
    }
    else {
        throw "Não foi possível identificar a pasta do script. Salve o arquivo .ps1 e execute-o por F5 no PowerShell ISE."
    }
}

$script:RootPath = '\\clx01fs\SistemaAssinatura$'

# Subpastas criadas automaticamente pelo script, se houver permissão.
$script:DataPath = Join-Path $script:RootPath "Dados"
$script:PdfPath = Join-Path $script:RootPath "PDFs"

$script:UsersFile = Join-Path $script:DataPath "usuarios.csv"
$script:DocumentsFile = Join-Path $script:DataPath "documentos.csv"
$script:StepsFile = Join-Path $script:DataPath "etapas.csv"
$script:HistoryFile = Join-Path $script:DataPath "historico.csv"
$script:NotificationsFile = Join-Path $script:DataPath "notificacoes.csv"
$script:DeletedDocumentsFile = Join-Path $script:DataPath "documentos_excluidos.csv"
$script:RecycleBinFile = Join-Path $script:DataPath "lixeira.csv"
$script:RecycleBinPath = Join-Path $script:RootPath "Lixeira"
$script:RecycleBinRetentionDays = 30

$script:CurrentUser = $null
$script:MainForm = $null
$script:ContentPanel = $null
$script:NotificationButton = $null

$script:ColorBackground = [System.Drawing.Color]::FromArgb(246, 246, 247)
$script:ColorSurface = [System.Drawing.Color]::White
$script:ColorBorder = [System.Drawing.Color]::FromArgb(228, 228, 231)
$script:ColorText = [System.Drawing.Color]::FromArgb(24, 24, 27)
$script:ColorMenu = $script:ColorText
$script:ColorPrimary = [System.Drawing.Color]::FromArgb(15, 61, 122)
$script:ColorInfo = [System.Drawing.Color]::FromArgb(79, 70, 229)
$script:ColorSuccess = [System.Drawing.Color]::FromArgb(22, 163, 74)
$script:ColorWarning = [System.Drawing.Color]::FromArgb(217, 119, 6)
$script:ColorDanger = [System.Drawing.Color]::FromArgb(220, 38, 38)
$script:ColorMuted = [System.Drawing.Color]::FromArgb(113, 113, 122)
$script:ColorHover = [System.Drawing.Color]::FromArgb(244, 244, 245)
$script:ColorActive = [System.Drawing.Color]::FromArgb(237, 237, 240)
$script:ColorWhite = [System.Drawing.Color]::White
$script:ColorBlack = $script:ColorText

$script:FontUi = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$script:FontUiBold = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$script:FontButton = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
try {
    $script:FontIcon = New-Object System.Drawing.Font("Segoe MDL2 Assets", 12, [System.Drawing.FontStyle]::Regular)
}
catch {
    $script:FontIcon = $script:FontUi
}

$script:SidebarExpandedWidth = 252
$script:SidebarCollapsedWidth = 76
$script:SidebarExpanded = $true
$script:SidebarTargetWidth = $script:SidebarExpandedWidth
$script:SidebarSyncing = $false
$script:AccountMenuClosedAt = [datetime]::MinValue
$script:NavItems = $null
$script:NavSections = $null

# =============================================================================
# FUNÇÕES BÁSICAS
# =============================================================================

function Get-NowText {
    return (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
}

function Show-AppMessage {
    param(
        [string]$Message,
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $script:AppName,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $Icon
    ) | Out-Null
}

function Confirm-AppMessage {
    param([string]$Message)

    $result = [System.Windows.Forms.MessageBox]::Show(
        $Message,
        "Confirmação",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Get-StatusColor {
    param([string]$Status)

    switch ($Status) {
        "PENDENTE"     { return $script:ColorWarning }
        "CONCLUIDA"    { return $script:ColorSuccess }
        "RECUSADA"     { return $script:ColorDanger }
        "FINALIZADO"   { return $script:ColorSuccess }
        "RECUSADO"     { return $script:ColorDanger }
        "EM_APROVACAO" { return $script:ColorInfo }
        default         { return $script:ColorMuted }
    }
}

function Get-StepText {
    param([string]$Type)

    if ($Type -eq "VERIFICACAO") {
        return "VERIFICAÇÃO"
    }

    return "APROVAÇÃO"
}

function Convert-ToBoolean {
    param($Value)

    if ($Value -is [bool]) {
        return $Value
    }

    return ([string]$Value).ToLowerInvariant() -eq "true"
}

function Send-CorporateEmail {
   param (
       [string]$ToEmail,
       [string]$Subject,
       [string]$Body
   )
   if ([string]::IsNullOrWhiteSpace($ToEmail)) {
       [System.Windows.Forms.MessageBox]::Show("O e-mail de destino está vazio!", "Erro de Envio")
       return
   }
   try {
       $smtpServer = "smtp.office365.com"
       $smtpPort = 587
       $msg = New-Object System.Net.Mail.MailMessage
       $msg.From = "emariafss@castrolandaservices.coop.br"
       $msg.To.Add($ToEmail)
       $msg.Subject = $Subject
       $msg.Body = $Body
       $msg.IsBodyHtml = $false
       $client = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
       $client.EnableSsl = $true
       $client.UseDefaultCredentials = $false
       # IMPORTANTE: Substitui "tua_senha_corporativa" pela senha real que usas para entrar no teu e-mail da empresa
       $password = "tua_senha_corporativa"
       $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
       $client.Credentials = New-Object System.Net.NetworkCredential("emariafss@castrolandaservices.coop.br", $securePassword)
       $client.Send($msg)
       $msg.Dispose()
       [System.Windows.Forms.MessageBox]::Show("E-mail enviado com sucesso para: $ToEmail", "Sucesso SMTP")
   }
   catch {
       [System.Windows.Forms.MessageBox]::Show("Falha ao enviar para $ToEmail. Erro: $($_.Exception.Message)", "Erro SMTP")
   }
}

# =============================================================================
# ARMAZENAMENTO CSV
# =============================================================================
function Initialize-DemoUsers {
   $users = Get-Users
   $myEmail = "emariafss@castrolandaservice.coop.br"
   $hasMyUser = $false
   foreach ($u in $users) {
       if ($u.email.ToLowerInvariant() -eq $myEmail) {
           $hasMyUser = $true
           break
       }
   }
   if ($users.Count -gt 0 -and $hasMyUser) {
       return
   }
   $now = Get-NowText
   $users = @(
       [pscustomobject]@{
           id        = 1
           nome      = "Maria"
           email     = $myEmail = "emariafss@castrolandaservices.coop.br"
           codigo    = "123456"
           cargo     = "ADMINISTRADOR"
           ativo     = "true"
           criado_em = $now
       },
       [pscustomobject]@{
           id        = 2
           nome      = "Admin"
           email     = "admin@exemplo.local"
           codigo    = "admin123"
           cargo     = "ADMINISTRADOR"
           ativo     = "true"
           criado_em = $now
       },
       [pscustomobject]@{
           id        = 3
           nome      = "João"
           email     = "joao@exemplo.local"
           codigo    = "joao123"
           cargo     = "TERCEIRO"
           ativo     = "true"
           criado_em = $now
       },
       [pscustomobject]@{
           id        = 4
           nome      = "Maria Teste"
           email     = "maria@exemplo.local"
           codigo    = "maria123"
           cargo     = "LIDER"
           ativo     = "true"
           criado_em = $now
       },
       [pscustomobject]@{
           id        = 5
           nome      = "Carlos"
           email     = "carlos@exemplo.local"
           codigo    = "carlos123"
           cargo     = "SUPERVISOR"
           ativo     = "true"
           criado_em = $now
       },
       [pscustomobject]@{
           id        = 6
           nome      = "Ana"
           email     = "ana@exemplo.local"
           codigo    = "ana123"
           cargo     = "COORDENADOR"
           ativo     = "true"
           criado_em = $now
       }
   )
   Save-CsvRows -Rows $users -FilePath $script:UsersFile
}

function Get-CsvRows {
    param([string]$FilePath)

    if (-not (Test-Path -Path $FilePath)) {
        return @()
    }

    $rows = Import-Csv -Path $FilePath -Delimiter ";"

    if ($null -eq $rows) {
        return @()
    }

    return @($rows)
}

function Save-CsvRows {
    param(
        [object[]]$Rows,
        [string]$FilePath
    )

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        return
    }

    $Rows | Export-Csv `
        -Path $FilePath `
        -Delimiter ";" `
        -NoTypeInformation `
        -Encoding UTF8
}

function Get-NextId {
    param([object[]]$Rows)

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        return 1
    }

    $max = 0

    foreach ($row in $Rows) {
        if ($null -ne $row -and $null -ne $row.id) {
            $currentId = [int]$row.id

            if ($currentId -gt $max) {
                $max = $currentId
            }
        }
    }

    return ($max + 1)
}

# =============================================================================
# DADOS
# =============================================================================

function Get-Users {
    return @(Get-CsvRows -FilePath $script:UsersFile)
}

function Get-Documents {
    return @(Get-CsvRows -FilePath $script:DocumentsFile)
}

function Get-Steps {
    return @(Get-CsvRows -FilePath $script:StepsFile)
}

function Get-History {
    return @(Get-CsvRows -FilePath $script:HistoryFile)
}

function Get-Notifications {
    return @(Get-CsvRows -FilePath $script:NotificationsFile)
}

function Get-UserById {
    param([int]$UserId)

    foreach ($user in (Get-Users)) {
        if ([int]$user.id -eq $UserId) {
            return $user
        }
    }

    return $null
}

function Get-DocumentById {
    param([int]$DocumentId)

    foreach ($document in (Get-Documents)) {
        if ([int]$document.id -eq $DocumentId) {
            return $document
        }
    }

    return $null
}

function Get-CurrentStep {
    param([int]$DocumentId)

    $document = Get-DocumentById -DocumentId $DocumentId

    if ($null -eq $document) {
        return $null
    }

    foreach ($step in (Get-Steps)) {
        if ([int]$step.id -eq [int]$document.etapa_atual_id) {
            return $step
        }
    }

    return $null
}

function Get-DocumentSteps {
    param([int]$DocumentId)

    $result = @()

    foreach ($step in (Get-Steps)) {
        if ([int]$step.documento_id -eq $DocumentId) {
            $result += $step
        }
    }

    return @($result | Sort-Object { [int]$_.ordem })
}

# =============================================================================
# USUÁRIOS FICTÍCIOS
# =============================================================================

function Initialize-DemoUsers {
    $users = Get-Users

    if ($users.Count -gt 0) {
        return
    }

    $now = Get-NowText

    $users = @(
        [pscustomobject]@{
            id        = 1
            nome      = "Admin"
            email     = "admin@exemplo.local"
            codigo    = "admin123"
            cargo     = "ADMINISTRADOR"
            ativo     = "true"
            criado_em = $now
        },
        [pscustomobject]@{
            id        = 2
            nome      = "João"
            email     = "joao@exemplo.local"
            codigo    = "joao123"
            cargo     = "TERCEIRO"
            ativo     = "true"
            criado_em = $now
        },
        [pscustomobject]@{
            id        = 3
            nome      = "Maria"
            email     = "maria@exemplo.local"
            codigo    = "maria123"
            cargo     = "LIDER"
            ativo     = "true"
            criado_em = $now
        },
        [pscustomobject]@{
            id        = 4
            nome      = "Carlos"
            email     = "carlos@exemplo.local"
            codigo    = "carlos123"
            cargo     = "SUPERVISOR"
            ativo     = "true"
            criado_em = $now
        },
        [pscustomobject]@{
            id        = 5
            nome      = "Ana"
            email     = "ana@exemplo.local"
            codigo    = "ana123"
            cargo     = "COORDENADOR"
            ativo     = "true"
            criado_em = $now
        }
    )

    Save-CsvRows -Rows $users -FilePath $script:UsersFile
}

# =============================================================================
# HISTÓRICO E NOTIFICAÇÕES
# =============================================================================

function Add-Notification {

   param(

        [int]$UserId,

        [int]$DocumentId,

        [string]$Message,

        [string]$DirectEmail = ""

   )

   $notifications = Get-Notifications

   $newEntry = [pscustomobject]@{

        id         = Get-NextId -Rows $notifications

        usuario_id = $UserId

        documento_id = $DocumentId

        mensagem   = $Message

        lida       = "false"

        criado_em  = Get-NowText

   }

   $notifications = @($notifications) + @($newEntry)

   Save-CsvRows -Rows $notifications -FilePath $script:NotificationsFile

   # Obtém o utilizador e define o destinatário direto (suportando qualquer domínio da cooperativa)

   $targetUser = Get-UserById -UserId $UserId

   $destinatarioFinal = if (-not [string]::IsNullOrWhiteSpace($DirectEmail)) { $DirectEmail } else { $targetUser.email }

   $subjectText = "Workflow Documental - Nova Pendência (Doc #$DocumentId)"

   $bodyText = "Olá," + [Environment]::NewLine + [Environment]::NewLine +

                "Existe uma nova pendência a aguardar ação no sistema de Workflow Documental:" + [Environment]::NewLine +

                "- ID do Documento: $DocumentId" + [Environment]::NewLine +

                "- Mensagem / Etapa: $Message" + [Environment]::NewLine + [Environment]::NewLine +

                "Por favor, aceda ao sistema para proceder com a análise."

   # Envia exclusivamente para o destinatário final

   if (-not [string]::IsNullOrWhiteSpace($destinatarioFinal)) {

        Send-CorporateEmail -ToEmail $destinatarioFinal -Subject $subjectText -Body $bodyText

   }

}
 
function Add-History {
   param(
       [int]$DocumentId,
       [int]$UserId,
       [string]$EventType,
       [string]$Description
   )
   $history = Get-History
   $newEntry = [pscustomobject]@{
       id           = Get-NextId -Rows $history
       documento_id = $DocumentId
       usuario_id   = $UserId
       tipo_evento  = $EventType
       descricao    = $Description
       criado_em    = Get-NowText
   }
   $history = @($history) + @($newEntry)
   Save-CsvRows -Rows $history -FilePath $script:HistoryFile
}

function Update-NotificationCounter {
    if ($null -eq $script:CurrentUser -or $null -eq $script:NotificationButton) {
        return
    }

    $count = 0

    foreach ($notification in (Get-Notifications)) {
        if (
            [int]$notification.usuario_id -eq [int]$script:CurrentUser.id -and
            (Convert-ToBoolean -Value $notification.lida) -eq $false
        ) {
            $count++
        }
    }

    $script:NotificationButton.Text = "Notificações ($count)"
}

# =============================================================================
# FUNÇÃO CENTRAL DO WORKFLOW
# =============================================================================

function Process-WorkflowAction {
    param(
        [int]$DocumentId,
        [int]$UserId,

        [ValidateSet("CONFIRMAR_VERIFICACAO", "APROVAR", "RECUSAR")]
        [string]$Action,

        [string]$Reason = ""
    )

    $users = Get-Users
    $documents = Get-Documents
    $steps = Get-Steps

    $user = $null
    $document = $null
    $currentStep = $null

    foreach ($item in $users) {
        if ([int]$item.id -eq $UserId) {
            $user = $item
            break
        }
    }

    foreach ($item in $documents) {
        if ([int]$item.id -eq $DocumentId) {
            $document = $item
            break
        }
    }

    if ($null -eq $user) {
        throw "Usuário não encontrado."
    }

    if ((Convert-ToBoolean -Value $user.ativo) -ne $true) {
        throw "Usuário inativo."
    }

    if ($null -eq $document) {
        throw "Documento não encontrado."
    }

    if ($document.status -ne "EM_APROVACAO") {
        throw "Documento não permite ação. Status atual: $($document.status)"
    }

    foreach ($step in $steps) {
        if ([int]$step.id -eq [int]$document.etapa_atual_id) {
            $currentStep = $step
            break
        }
    }

    if ($null -eq $currentStep) {
        throw "Etapa atual não encontrada."
    }

    if ($currentStep.status -ne "PENDENTE") {
        throw "A etapa atual não está pendente."
    }

    if ([int]$currentStep.responsavel_id -ne $UserId) {
        throw "Você não é o responsável pela etapa atual."
    }

    if ($Action -eq "CONFIRMAR_VERIFICACAO" -and $currentStep.tipo -ne "VERIFICACAO") {
        throw "Ação inválida para esta etapa."
    }

    if ($Action -eq "APROVAR" -and $currentStep.tipo -ne "APROVACAO") {
        throw "Ação inválida para esta etapa."
    }

    if ($Action -eq "RECUSAR" -and [string]::IsNullOrWhiteSpace($Reason)) {
        throw "O motivo da recusa é obrigatório."
    }

    $now = Get-NowText

    # RECUSA
    if ($Action -eq "RECUSAR") {
        foreach ($step in $steps) {
            if ([int]$step.id -eq [int]$currentStep.id) {
                $step.status = "RECUSADA"
                $step.concluido_em = $now
            }
        }

        foreach ($item in $documents) {
            if ([int]$item.id -eq $DocumentId) {
                $item.status = "RECUSADO"
            }
        }

        Save-CsvRows -Rows $steps -FilePath $script:StepsFile
        Save-CsvRows -Rows $documents -FilePath $script:DocumentsFile

        Add-History `
            -DocumentId $DocumentId `
            -UserId $UserId `
            -EventType "RECUSA" `
            -Description "Documento recusado por $($user.nome). Motivo: $Reason"

        foreach ($admin in $users) {
            if (
                $admin.cargo -eq "ADMINISTRADOR" -and
                (Convert-ToBoolean -Value $admin.ativo) -eq $true
            ) {
                Add-Notification `
                    -UserId ([int]$admin.id) `
                    -DocumentId $DocumentId `
                    -Message "Documento $($document.numero) foi recusado por $($user.nome)."
            }
        }

        return "Documento recusado. Fluxo encerrado."
    }

    # CONCLUSÃO
    foreach ($step in $steps) {
        if ([int]$step.id -eq [int]$currentStep.id) {
            $step.status = "CONCLUIDA"
            $step.concluido_em = $now
        }
    }

    if ($Action -eq "CONFIRMAR_VERIFICACAO") {
        Add-History `
            -DocumentId $DocumentId `
            -UserId $UserId `
            -EventType "VERIFICACAO" `
            -Description "Verificação confirmada por $($user.nome)."
    }
    else {
        Add-History `
            -DocumentId $DocumentId `
            -UserId $UserId `
            -EventType "APROVACAO" `
            -Description "Documento aprovado por $($user.nome)."
    }

    # PRÓXIMA ETAPA
    $nextStep = $null
    $lowestOrder = 999999

    foreach ($step in $steps) {
        if (
            [int]$step.documento_id -eq $DocumentId -and
            [int]$step.ordem -gt [int]$currentStep.ordem -and
            [int]$step.ordem -lt $lowestOrder
        ) {
            $nextStep = $step
            $lowestOrder = [int]$step.ordem
        }
    }

    if ($null -ne $nextStep) {
        foreach ($step in $steps) {
            if ([int]$step.id -eq [int]$nextStep.id) {
                $step.status = "PENDENTE"
                $step.iniciado_em = $now
            }
        }

        foreach ($item in $documents) {
            if ([int]$item.id -eq $DocumentId) {
                $item.etapa_atual_id = [int]$nextStep.id
            }
        }

        Save-CsvRows -Rows $steps -FilePath $script:StepsFile
        Save-CsvRows -Rows $documents -FilePath $script:DocumentsFile

        $nextUser = Get-UserById -UserId ([int]$nextStep.responsavel_id)

        Add-History `
            -DocumentId $DocumentId `
            -UserId $UserId `
            -EventType "ETAPA_LIBERADA" `
            -Description "Documento encaminhado para $($nextUser.nome)."

        Add-Notification `
            -UserId ([int]$nextStep.responsavel_id) `
            -DocumentId $DocumentId `
            -Message "Documento $($document.numero) está aguardando sua ação."

        return "Etapa concluída. Documento encaminhado para $($nextUser.nome)."
    }

    # FINALIZAÇÃO
    foreach ($item in $documents) {
        if ([int]$item.id -eq $DocumentId) {
            $item.status = "FINALIZADO"
            $item.finalizado_em = $now
        }
    }

    Save-CsvRows -Rows $steps -FilePath $script:StepsFile
    Save-CsvRows -Rows $documents -FilePath $script:DocumentsFile

    Add-History `
        -DocumentId $DocumentId `
        -UserId $UserId `
        -EventType "FINALIZADO" `
        -Description "Documento finalizado por $($user.nome)."

    return "Documento finalizado com sucesso."
}

# =============================================================================
# INTERFACE
# =============================================================================

function New-AppLabel {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 250,
        [int]$Height = 28,
        [int]$FontSize = 10,
        [bool]$Bold = $false,
        $ForeColor = $null
    )

    if ($null -eq $ForeColor) {
        $ForeColor = $script:ColorBlack
    }

    $fontStyle = [System.Drawing.FontStyle]::Regular

    if ($Bold) {
        $fontStyle = [System.Drawing.FontStyle]::Bold
    }

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, $Height)
    $label.ForeColor = $ForeColor
    $label.Font = New-Object System.Drawing.Font("Segoe UI", $FontSize, $fontStyle)

    return $label
}

function Get-ShiftedColor {
    param(
        [System.Drawing.Color]$Color,
        [int]$Delta
    )

    $red = [Math]::Min(255, [Math]::Max(0, $Color.R + $Delta))
    $green = [Math]::Min(255, [Math]::Max(0, $Color.G + $Delta))
    $blue = [Math]::Min(255, [Math]::Max(0, $Color.B + $Delta))
    return [System.Drawing.Color]::FromArgb($red, $green, $blue)
}

function Get-RoundPath {
    param(
        [System.Drawing.Rectangle]$Bounds,
        [int]$Radius = 10
    )

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $diameter = $Radius * 2
    $path.AddArc($Bounds.X, $Bounds.Y, $diameter, $diameter, 180, 90)
    $path.AddArc(($Bounds.Right - $diameter), $Bounds.Y, $diameter, $diameter, 270, 90)
    $path.AddArc(($Bounds.Right - $diameter), ($Bounds.Bottom - $diameter), $diameter, $diameter, 0, 90)
    $path.AddArc($Bounds.X, ($Bounds.Bottom - $diameter), $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function Set-RoundRegion {
    param(
        [System.Windows.Forms.Control]$Control,
        [int]$Radius = 10
    )

    if ($null -eq $Control) {
        return
    }

    $width = $Control.Width
    $height = $Control.Height
    if ($width -lt ($Radius * 2) -or $height -lt ($Radius * 2)) {
        $Control.Region = $null
        return
    }

    $path = Get-RoundPath -Bounds (New-Object System.Drawing.Rectangle(0, 0, $width, $height)) -Radius $Radius
    $Control.Region = New-Object System.Drawing.Region($path)
    $path.Dispose()
}

function Enable-DoubleBuffer {
    param([System.Windows.Forms.Control]$Control)

    if ($null -eq $Control) {
        return
    }

    $flags = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
    $property = [System.Windows.Forms.Control].GetProperty("DoubleBuffered", $flags)
    if ($null -ne $property) {
        $property.SetValue($Control, $true, $null)
    }
}

function Enable-ResizeRedraw {
    param([System.Windows.Forms.Control]$Control)

    if ($null -eq $Control) {
        return
    }

    Enable-DoubleBuffer -Control $Control
    $flags = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
    $method = [System.Windows.Forms.Control].GetMethod("SetStyle", $flags)
    if ($null -eq $method) {
        return
    }

    $style = [System.Windows.Forms.ControlStyles]::ResizeRedraw -bor `
        [System.Windows.Forms.ControlStyles]::OptimizedDoubleBuffer -bor `
        [System.Windows.Forms.ControlStyles]::AllPaintingInWmPaint
    $method.Invoke($Control, @($style, $true)) | Out-Null
}

function New-AppButton {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 180,
        [int]$Height = 38,
        $BackColor = $null
    )

    if ($null -eq $BackColor) {
        $BackColor = $script:ColorPrimary
    }

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, $Height)
    $button.BackColor = $BackColor
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.Font = $script:FontButton
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand

    $brightness = ($BackColor.R * 0.299) + ($BackColor.G * 0.587) + ($BackColor.B * 0.114)
    $hoverDelta = 18
    if ($brightness -gt 186) {
        $button.ForeColor = $script:ColorText
        $button.FlatAppearance.BorderSize = 1
        $button.FlatAppearance.BorderColor = $script:ColorBorder
        $hoverDelta = -12
    }
    else {
        $button.ForeColor = $script:ColorWhite
        $button.FlatAppearance.BorderSize = 0
    }

    $button.FlatAppearance.MouseOverBackColor = Get-ShiftedColor -Color $BackColor -Delta $hoverDelta
    $button.FlatAppearance.MouseDownBackColor = Get-ShiftedColor -Color $BackColor -Delta ($hoverDelta * 2)
    return $button
}

function Set-AppGridLook {
    param($Grid)

    if ($null -eq $Grid) {
        return
    }

    $Grid.BackgroundColor = $script:ColorSurface
    $Grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $Grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $Grid.GridColor = $script:ColorBorder
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::None
    $Grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $Grid.ColumnHeadersHeight = 38
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = $script:ColorBackground
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = $script:ColorMuted
    $Grid.ColumnHeadersDefaultCellStyle.Font = $script:FontUi
    $Grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $script:ColorBackground
    $Grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = $script:ColorMuted
    $Grid.DefaultCellStyle.BackColor = $script:ColorSurface
    $Grid.DefaultCellStyle.ForeColor = $script:ColorText
    $Grid.DefaultCellStyle.Font = $script:FontUi
    $Grid.DefaultCellStyle.SelectionBackColor = $script:ColorActive
    $Grid.DefaultCellStyle.SelectionForeColor = $script:ColorText
    $Grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
    $Grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $Grid.AlternatingRowsDefaultCellStyle.ForeColor = $script:ColorText
    $Grid.RowTemplate.Height = 36
    $Grid.AllowUserToResizeRows = $false

    if ($Grid.Dock -eq [System.Windows.Forms.DockStyle]::None) {
        $parentWidth = 0
        if ($null -ne $Grid.Parent) {
            $parentWidth = $Grid.Parent.ClientSize.Width
        }
        elseif ($null -ne $script:ContentPanel) {
            $parentWidth = $script:ContentPanel.ClientSize.Width
        }

        if ($parentWidth -gt 240) {
            $Grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
            $Grid.Width = $parentWidth - $Grid.Left - 24
        }
    }
}

function Set-ActiveNav {
    param([string]$Key)

    $titles = @{
        "DASHBOARD"           = "Dashboard"
        "DOCUMENTOS"          = "Documentos"
        "NOTIFICACOES"        = "Notificações"
        "CADASTRO_DOCUMENTO"  = "Cadastrar documento"
        "DETALHES_DOCUMENTO"  = "Documento"
        "GERENCIAR_USUARIOS"  = "Usuários"
        "EXCLUSAO_ARQUIVOS"   = "Exclusão de arquivos"
        "LIXEIRA"             = "Lixeira"
    }

    if ($null -ne $script:HeaderTitle -and $titles.ContainsKey($Key)) {
        $script:HeaderTitle.Text = $titles[$Key]
    }

    if ($null -eq $script:NavItems) {
        return
    }

    foreach ($item in $script:NavItems) {
        $active = ($item.Tag.Key -eq $Key)
        $item.Tag.Active = $active
        $caption = $item.Controls["caption"]
        $icon = $item.Controls["icon"]

        if ($active) {
            $item.BackColor = $script:ColorActive
            if ($null -ne $caption) {
                $caption.ForeColor = $script:ColorText
                $caption.Font = $script:FontUiBold
            }
            if ($null -ne $icon) {
                $icon.ForeColor = $script:ColorText
            }
        }
        else {
            $item.BackColor = $script:ColorSurface
            if ($null -ne $caption) {
                $caption.ForeColor = $script:ColorMuted
                $caption.Font = $script:FontUi
            }
            if ($null -ne $icon) {
                $icon.ForeColor = $script:ColorMuted
            }
        }
    }
}

function Clear-Content {
    if ($null -ne $script:ContentPanel) {
        $script:ContentPanel.Controls.Clear()
    }

    Set-ActiveNav -Key $script:CurrentPage
}

# =============================================================================
# LOGIN
# =============================================================================

function Show-Login {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Login - Workflow Documental"
    $form.Size = New-Object System.Drawing.Size(460, 350)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.Font = $script:FontUi
    $form.BackColor = $script:ColorSurface
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $form.Controls.Add((New-AppLabel -Text "Workflow Documental" -X 40 -Y 30 -Width 370 -Height 35 -FontSize 18 -Bold $true -ForeColor $script:ColorText))
    $form.Controls.Add((New-AppLabel -Text "Entre com e-mail e código de acesso" -X 40 -Y 70 -Width 350 -Height 25 -FontSize 10 -ForeColor $script:ColorMuted))
    $form.Controls.Add((New-AppLabel -Text "E-mail" -X 40 -Y 120 -Width 150 -Height 25 -FontSize 10 -Bold $true))
    $form.Controls.Add((New-AppLabel -Text "Código de acesso" -X 40 -Y 185 -Width 180 -Height 25 -FontSize 10 -Bold $true))

    $emailBox = New-Object System.Windows.Forms.TextBox
    $emailBox.Location = New-Object System.Drawing.Point(40, 145)
    $emailBox.Size = New-Object System.Drawing.Size(360, 28)

    $codeBox = New-Object System.Windows.Forms.TextBox
    $codeBox.Location = New-Object System.Drawing.Point(40, 210)
    $codeBox.Size = New-Object System.Drawing.Size(360, 28)
    $codeBox.UseSystemPasswordChar = $true

    $loginButton = New-AppButton -Text "Entrar" -X 40 -Y 270 -Width 360 -Height 40 -BackColor $script:ColorPrimary

    $loginButton.Tag = [pscustomobject]@{
        Form = $form
        EmailBox = $emailBox
        CodeBox = $codeBox
    }

    $loginButton.Add_Click({
        try {
            $context = $this.Tag
            $email = $context.EmailBox.Text.Trim().ToLowerInvariant()
            $code = $context.CodeBox.Text

            $userFound = $null

            foreach ($user in (Get-Users)) {
                if (
                    $user.email.ToLowerInvariant() -eq $email -and
                    $user.codigo -eq $code
                ) {
                    $userFound = $user
                    break
                }
            }

            if ($null -eq $userFound) {
                Show-AppMessage -Message "E-mail ou código de acesso inválidos." -Icon Warning
                return
            }

            if ((Convert-ToBoolean -Value $userFound.ativo) -ne $true) {
                Show-AppMessage -Message "Usuário inativo." -Icon Warning
                return
            }

            $script:CurrentUser = $userFound
            $context.Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $context.Form.Close()
        }
        catch {
            Show-AppMessage -Message $_.Exception.Message -Icon Error
        }
    })

    $form.Controls.Add($emailBox)
    $form.Controls.Add($codeBox)
    $form.Controls.Add($loginButton)
    $form.AcceptButton = $loginButton

    return $form.ShowDialog()
}

# =============================================================================
# DASHBOARD
# =============================================================================

function New-MetricCard {
    param(
        [string]$Caption,
        [string]$Value,
        $Accent
    )

    $card = New-Object System.Windows.Forms.Panel
    $card.Dock = [System.Windows.Forms.DockStyle]::Fill
    $card.Margin = New-Object System.Windows.Forms.Padding(8, 4, 8, 8)
    $card.BackColor = $script:ColorSurface
    $card.Controls.Add((New-AppLabel -Text $Caption -X 16 -Y 14 -Width 180 -Height 20 -FontSize 9 -ForeColor $script:ColorMuted))
    $card.Controls.Add((New-AppLabel -Text $Value -X 16 -Y 36 -Width 160 -Height 36 -FontSize 20 -Bold $true -ForeColor $Accent))
    Enable-ResizeRedraw -Control $card
    $card.Add_Resize({
        Set-RoundRegion -Control $this -Radius 12
        $this.Invalidate()
        if ($null -ne $this.Parent) {
            $this.Parent.Invalidate()
        }
    })
    $card.Add_Paint({
        $graphics = $_.Graphics
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $bounds = New-Object System.Drawing.Rectangle(0, 0, ($this.Width - 1), ($this.Height - 1))
        if ($bounds.Width -lt 24 -or $bounds.Height -lt 24) {
            return
        }
        $pen = New-Object System.Drawing.Pen($script:ColorBorder)
        $path = Get-RoundPath -Bounds $bounds -Radius 12
        $graphics.DrawPath($pen, $path)
        $path.Dispose()
        $pen.Dispose()
    })
    return $card
}

function Show-Dashboard {
    $script:CurrentPage = "DASHBOARD"
    $script:CurrentDocumentId = $null

    Clear-Content
    Update-NotificationCounter

    $documents = Get-Documents
    $steps = Get-Steps

    $myPending = 0
    $inApproval = 0
    $finished = 0
    $rejected = 0

    foreach ($step in $steps) {
        if ([int]$step.responsavel_id -eq [int]$script:CurrentUser.id -and $step.status -eq "PENDENTE") {
            $myPending++
        }
    }

    foreach ($document in $documents) {
        if ($document.status -eq "EM_APROVACAO") { $inApproval++ }
        if ($document.status -eq "FINALIZADO") { $finished++ }
        if ($document.status -eq "RECUSADO") { $rejected++ }
    }

    $page = New-Object System.Windows.Forms.Panel
    $page.Dock = [System.Windows.Forms.DockStyle]::Fill
    $page.BackColor = $script:ColorBackground

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = [System.Windows.Forms.DockStyle]::Top
    $header.Height = 72
    $header.BackColor = $script:ColorBackground
    $header.Controls.Add((New-AppLabel -Text "Olá, $($script:CurrentUser.nome)" -X 24 -Y 8 -Width 640 -Height 34 -FontSize 20 -Bold $true -ForeColor $script:ColorText))
    $header.Controls.Add((New-AppLabel -Text "Acompanhe o fluxo e as suas pendências." -X 24 -Y 42 -Width 640 -Height 22 -FontSize 10 -ForeColor $script:ColorMuted))

    $cardsHost = New-Object System.Windows.Forms.TableLayoutPanel
    $cardsHost.Dock = [System.Windows.Forms.DockStyle]::Top
    $cardsHost.Height = 112
    $cardsHost.ColumnCount = 4
    $cardsHost.RowCount = 1
    $cardsHost.BackColor = $script:ColorBackground
    $cardsHost.Padding = New-Object System.Windows.Forms.Padding(16, 0, 16, 0)
    if ($cardsHost.RowStyles.Count -eq 0) {
        [void]$cardsHost.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    }
    else {
        $cardsHost.RowStyles[0].SizeType = [System.Windows.Forms.SizeType]::Percent
        $cardsHost.RowStyles[0].Height = 100
    }

    for ($columnIndex = 0; $columnIndex -lt 4; $columnIndex++) {
        if ($cardsHost.ColumnStyles.Count -le $columnIndex) {
            [void]$cardsHost.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
        }
        else {
            $cardsHost.ColumnStyles[$columnIndex].SizeType = [System.Windows.Forms.SizeType]::Percent
            $cardsHost.ColumnStyles[$columnIndex].Width = 25
        }
    }

    $metricCards = @(
        @{ Caption = "Pendências"; Value = "$myPending"; Color = $script:ColorWarning },
        @{ Caption = "Em aprovação"; Value = "$inApproval"; Color = $script:ColorInfo },
        @{ Caption = "Finalizados"; Value = "$finished"; Color = $script:ColorSuccess },
        @{ Caption = "Recusados"; Value = "$rejected"; Color = $script:ColorDanger }
    )

    for ($columnIndex = 0; $columnIndex -lt $metricCards.Count; $columnIndex++) {
        $metric = $metricCards[$columnIndex]
        $card = New-MetricCard -Caption $metric.Caption -Value $metric.Value -Accent $metric.Color
        $cardsHost.Controls.Add($card, $columnIndex, 0)
    }

    $section = New-Object System.Windows.Forms.Panel
    $section.Dock = [System.Windows.Forms.DockStyle]::Top
    $section.Height = 36
    $section.BackColor = $script:ColorBackground
    $section.Controls.Add((New-AppLabel -Text "Minhas pendências" -X 24 -Y 4 -Width 400 -Height 28 -FontSize 12 -Bold $true -ForeColor $script:ColorText))

    $footer = New-Object System.Windows.Forms.Panel
    $footer.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footer.Height = 64
    $footer.BackColor = $script:ColorBackground

    $gridHost = New-Object System.Windows.Forms.Panel
    $gridHost.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridHost.Padding = New-Object System.Windows.Forms.Padding(24, 0, 24, 8)
    $gridHost.BackColor = $script:ColorBackground

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $false
    Set-AppGridLook -Grid $grid

    foreach ($columnName in @("ID", "Número", "Pessoa", "Etapa", "Iniciada em")) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.Name = $columnName
        $column.HeaderText = $columnName
        [void]$grid.Columns.Add($column)
    }

    foreach ($step in $steps) {
        if ([int]$step.responsavel_id -eq [int]$script:CurrentUser.id -and $step.status -eq "PENDENTE") {
            $document = Get-DocumentById -DocumentId ([int]$step.documento_id)

            if ($null -ne $document -and $document.status -eq "EM_APROVACAO") {
                [void]$grid.Rows.Add(
                    $document.id,
                    $document.numero,
                    $document.pessoa_relacionada,
                    (Get-StepText -Type $step.tipo),
                    $step.iniciado_em
                )
            }
        }
    }

    $openButton = New-AppButton -Text "Abrir documento" -X 24 -Y 12 -Width 180 -Height 38 -BackColor $script:ColorPrimary
    $openButton.Tag = $grid

    $openButton.Add_Click({
        try {
            $selectedGrid = $this.Tag

            if ($selectedGrid.SelectedRows.Count -eq 0) {
                Show-AppMessage -Message "Selecione um documento." -Icon Warning
                return
            }

            Show-DocumentDetails -DocumentId ([int]$selectedGrid.SelectedRows[0].Cells["ID"].Value)
        }
        catch {
            Show-AppMessage -Message $_.Exception.Message -Icon Error
        }
    })

    $footer.Controls.Add($openButton)
    $gridHost.Controls.Add($grid)
    $page.Controls.Add($gridHost)
    $page.Controls.Add($footer)
    $page.Controls.Add($section)
    $page.Controls.Add($cardsHost)
    $page.Controls.Add($header)
    $script:ContentPanel.Controls.Add($page)
}

# =============================================================================
# CADASTRO DE DOCUMENTOS
# =============================================================================

function Show-RegisterDocument {
    $script:CurrentPage = "CADASTRO_DOCUMENTO"
    $script:CurrentDocumentId = $null

    if ($script:CurrentUser.cargo -ne "ADMINISTRADOR") {
        Show-AppMessage -Message "Somente administradores podem cadastrar documentos." -Icon Warning
        return
    }

    Clear-Content

    $script:ContentPanel.Controls.Add((New-AppLabel -Text "CADASTRAR DOCUMENTO" -X 25 -Y 20 -Width 600 -Height 35 -FontSize 18 -Bold $true -ForeColor $script:ColorMenu))

    $labels = @(
        @{ Text = "Número do documento"; Y = 85 },
        @{ Text = "Pessoa relacionada"; Y = 140 },
        @{ Text = "Data"; Y = 195 },
        @{ Text = "Arquivo PDF"; Y = 250 },
        @{ Text = "Terceiro"; Y = 305 },
        @{ Text = "Líder"; Y = 360 },
        @{ Text = "Supervisor"; Y = 415 },
        @{ Text = "Coordenador"; Y = 470 }
    )

    foreach ($labelData in $labels) {
        $script:ContentPanel.Controls.Add((New-AppLabel -Text $labelData.Text -X 25 -Y $labelData.Y -Width 220 -Height 25 -FontSize 10 -Bold $true))
    }

    $numberBox = New-Object System.Windows.Forms.TextBox
    $numberBox.Location = New-Object System.Drawing.Point(250, 82)
    $numberBox.Size = New-Object System.Drawing.Size(280, 28)

    $personBox = New-Object System.Windows.Forms.TextBox
    $personBox.Location = New-Object System.Drawing.Point(250, 137)
    $personBox.Size = New-Object System.Drawing.Size(280, 28)

    $datePicker = New-Object System.Windows.Forms.DateTimePicker
    $datePicker.Location = New-Object System.Drawing.Point(250, 192)
    $datePicker.Size = New-Object System.Drawing.Size(280, 28)
    $datePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short

    $pdfBox = New-Object System.Windows.Forms.TextBox
    $pdfBox.Location = New-Object System.Drawing.Point(250, 247)
    $pdfBox.Size = New-Object System.Drawing.Size(400, 28)
    $pdfBox.ReadOnly = $true

    $browseButton = New-AppButton -Text "SELECIONAR PDF" -X 665 -Y 245 -Width 190 -Height 33 -BackColor $script:ColorPrimary
    $browseButton.Tag = $pdfBox

    $browseButton.Add_Click({
        try {
            $targetBox = $this.Tag
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Filter = "Arquivos PDF (*.pdf)|*.pdf"
            $dialog.Title = "Selecione o arquivo PDF"

            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $targetBox.Text = $dialog.FileName
            }

            $dialog.Dispose()
        }
        catch {
            Show-AppMessage -Message $_.Exception.Message -Icon Error
        }
    })

    $thirdCombo = New-Object System.Windows.Forms.ComboBox
    $thirdCombo.Location = New-Object System.Drawing.Point(250, 302)
    $thirdCombo.Size = New-Object System.Drawing.Size(280, 28)
    $thirdCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $thirdCombo.DisplayMember = "nome"

    $leaderCombo = New-Object System.Windows.Forms.ComboBox
    $leaderCombo.Location = New-Object System.Drawing.Point(250, 357)
    $leaderCombo.Size = New-Object System.Drawing.Size(280, 28)
    $leaderCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $leaderCombo.DisplayMember = "nome"

    $supervisorCombo = New-Object System.Windows.Forms.ComboBox
    $supervisorCombo.Location = New-Object System.Drawing.Point(250, 412)
    $supervisorCombo.Size = New-Object System.Drawing.Size(280, 28)
    $supervisorCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $supervisorCombo.DisplayMember = "nome"

    $coordinatorCombo = New-Object System.Windows.Forms.ComboBox
    $coordinatorCombo.Location = New-Object System.Drawing.Point(250, 467)
    $coordinatorCombo.Size = New-Object System.Drawing.Size(280, 28)
    $coordinatorCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $coordinatorCombo.DisplayMember = "nome"

    foreach ($user in (Get-Users)) {
        if ((Convert-ToBoolean -Value $user.ativo) -eq $true -and $user.cargo -eq "TERCEIRO") {
            [void]$thirdCombo.Items.Add($user)
        }

        if ((Convert-ToBoolean -Value $user.ativo) -eq $true -and $user.cargo -eq "LIDER") {
            [void]$leaderCombo.Items.Add($user)
        }

        if ((Convert-ToBoolean -Value $user.ativo) -eq $true -and $user.cargo -eq "SUPERVISOR") {
            [void]$supervisorCombo.Items.Add($user)
        }

        if ((Convert-ToBoolean -Value $user.ativo) -eq $true -and $user.cargo -eq "COORDENADOR") {
            [void]$coordinatorCombo.Items.Add($user)
        }
    }

    if ($thirdCombo.Items.Count -gt 0) { $thirdCombo.SelectedIndex = 0 }
    if ($leaderCombo.Items.Count -gt 0) { $leaderCombo.SelectedIndex = 0 }
    if ($supervisorCombo.Items.Count -gt 0) { $supervisorCombo.SelectedIndex = 0 }
    if ($coordinatorCombo.Items.Count -gt 0) { $coordinatorCombo.SelectedIndex = 0 }

    $saveButton = New-AppButton -Text "CADASTRAR DOCUMENTO" -X 600 -Y 465 -Width 255 -Height 38 -BackColor $script:ColorSuccess

    $saveButton.Tag = [pscustomobject]@{
        NumberBox = $numberBox
        PersonBox = $personBox
        DatePicker = $datePicker
        PdfBox = $pdfBox
        ThirdCombo = $thirdCombo
        LeaderCombo = $leaderCombo
        SupervisorCombo = $supervisorCombo
        CoordinatorCombo = $coordinatorCombo
    }

    $saveButton.Add_Click({
        try {
            $context = $this.Tag
            $number = $context.NumberBox.Text.Trim()
            $person = $context.PersonBox.Text.Trim()
            $sourcePdf = $context.PdfBox.Text.Trim()

            if ([string]::IsNullOrWhiteSpace($number)) {
                Show-AppMessage -Message "Informe o número do documento." -Icon Warning
                return
            }

            if ([string]::IsNullOrWhiteSpace($person)) {
                Show-AppMessage -Message "Informe a pessoa relacionada." -Icon Warning
                return
            }

            if ([string]::IsNullOrWhiteSpace($sourcePdf) -or -not (Test-Path -Path $sourcePdf)) {
                Show-AppMessage -Message "Selecione um arquivo PDF válido." -Icon Warning
                return
            }

            if (
                $null -eq $context.ThirdCombo.SelectedItem -or
                $null -eq $context.LeaderCombo.SelectedItem -or
                $null -eq $context.SupervisorCombo.SelectedItem -or
                $null -eq $context.CoordinatorCombo.SelectedItem
            ) {
                Show-AppMessage -Message "Defina todos os responsáveis." -Icon Warning
                return
            }

            $documents = Get-Documents

            foreach ($existingDocument in $documents) {
                if ($existingDocument.numero -eq $number) {
                    Show-AppMessage -Message "Já existe documento com esse número." -Icon Warning
                    return
                }
            }

            $newDocumentId = Get-NextId -Rows $documents
            $destinationPdf = Join-Path $script:PdfPath ("{0}_{1}.pdf" -f $number, $newDocumentId)

            Copy-Item -Path $sourcePdf -Destination $destinationPdf -Force -ErrorAction Stop

            $steps = Get-Steps
            $firstStepId = Get-NextId -Rows $steps
            $now = Get-NowText

            $thirdUser = $context.ThirdCombo.SelectedItem
            $leaderUser = $context.LeaderCombo.SelectedItem
            $supervisorUser = $context.SupervisorCombo.SelectedItem
            $coordinatorUser = $context.CoordinatorCombo.SelectedItem

            $newDocument = [pscustomobject]@{
                id                 = $newDocumentId
                numero             = $number
                pessoa_relacionada = $person
                data_documento     = $context.DatePicker.Value.ToString("yyyy-MM-dd")
                arquivo_pdf        = $destinationPdf
                status             = "EM_APROVACAO"
                etapa_atual_id     = $firstStepId
                criado_por         = [int]$script:CurrentUser.id
                criado_em          = $now
                finalizado_em      = ""
            }

            $step1 = [pscustomobject]@{
                id             = $firstStepId
                documento_id   = $newDocumentId
                ordem          = 1
                tipo           = "VERIFICACAO"
                responsavel_id = [int]$thirdUser.id
                status         = "PENDENTE"
                iniciado_em    = $now
                concluido_em   = ""
            }

            $step2 = [pscustomobject]@{
                id             = ($firstStepId + 1)
                documento_id   = $newDocumentId
                ordem          = 2
                tipo           = "APROVACAO"
                responsavel_id = [int]$leaderUser.id
                status         = "AGUARDANDO"
                iniciado_em    = ""
                concluido_em   = ""
            }

            $step3 = [pscustomobject]@{
                id             = ($firstStepId + 2)
                documento_id   = $newDocumentId
                ordem          = 3
                tipo           = "APROVACAO"
                responsavel_id = [int]$supervisorUser.id
                status         = "AGUARDANDO"
                iniciado_em    = ""
                concluido_em   = ""
            }

            $step4 = [pscustomobject]@{
                id             = ($firstStepId + 3)
                documento_id   = $newDocumentId
                ordem          = 4
                tipo           = "APROVACAO"
                responsavel_id = [int]$coordinatorUser.id
                status         = "AGUARDANDO"
                iniciado_em    = ""
                concluido_em   = ""
            }

            $documents = @($documents) + @($newDocument)
            $steps = @($steps) + @($step1, $step2, $step3, $step4)

            Save-CsvRows -Rows $documents -FilePath $script:DocumentsFile
            Save-CsvRows -Rows $steps -FilePath $script:StepsFile

            Add-History `
                -DocumentId $newDocumentId `
                -UserId ([int]$script:CurrentUser.id) `
                -EventType "CRIADO" `
                -Description "Documento criado por $($script:CurrentUser.nome)."

            Add-Notification `
                -UserId ([int]$thirdUser.id) `
                -DocumentId $newDocumentId `
                -Message "Novo documento $number aguardando sua verificação."

            Show-AppMessage -Message "Documento cadastrado. Etapa liberada para $($thirdUser.nome)." -Icon Information

            Show-DocumentDetails -DocumentId $newDocumentId
        }
        catch {
            Show-AppMessage -Message $_.Exception.Message -Icon Error
        }
    })

    $script:ContentPanel.Controls.Add($numberBox)
    $script:ContentPanel.Controls.Add($personBox)
    $script:ContentPanel.Controls.Add($datePicker)
    $script:ContentPanel.Controls.Add($pdfBox)
    $script:ContentPanel.Controls.Add($browseButton)
    $script:ContentPanel.Controls.Add($thirdCombo)
    $script:ContentPanel.Controls.Add($leaderCombo)
    $script:ContentPanel.Controls.Add($supervisorCombo)
    $script:ContentPanel.Controls.Add($coordinatorCombo)
    $script:ContentPanel.Controls.Add($saveButton)
}

# =============================================================================
# DETALHES DO DOCUMENTO
# =============================================================================

function Show-DocumentDetails {
    param([int]$DocumentId)

    $script:CurrentPage = "DETALHES_DOCUMENTO"
    $script:CurrentDocumentId = $DocumentId

    Clear-Content
    Update-NotificationCounter

    $document = Get-DocumentById -DocumentId $DocumentId

    if ($null -eq $document) {
        Show-AppMessage -Message "Documento não encontrado." -Icon Error
        Show-Dashboard
        return
    }

    $currentStep = Get-CurrentStep -DocumentId $DocumentId
    $responsibleName = "-"

    if ($null -ne $currentStep) {
        $responsible = Get-UserById -UserId ([int]$currentStep.responsavel_id)

        if ($null -ne $responsible) {
            $responsibleName = $responsible.nome
        }
    }

    $script:ContentPanel.Controls.Add((New-AppLabel -Text "DOCUMENTO $($document.numero)" -X 25 -Y 15 -Width 550 -Height 35 -FontSize 18 -Bold $true -ForeColor $script:ColorMenu))
    $script:ContentPanel.Controls.Add((New-AppLabel -Text "STATUS: $($document.status)" -X 675 -Y 20 -Width 210 -Height 28 -FontSize 11 -Bold $true -ForeColor (Get-StatusColor -Status $document.status)))
    $script:ContentPanel.Controls.Add((New-AppLabel -Text "Pessoa: $($document.pessoa_relacionada)" -X 25 -Y 60 -Width 400 -Height 25 -FontSize 10))
    $script:ContentPanel.Controls.Add((New-AppLabel -Text "Data: $($document.data_documento)" -X 25 -Y 85 -Width 400 -Height 25 -FontSize 10))
    $script:ContentPanel.Controls.Add((New-AppLabel -Text "Responsável atual: $responsibleName" -X 25 -Y 110 -Width 500 -Height 25 -FontSize 10 -Bold $true -ForeColor $script:ColorInfo))
    $script:ContentPanel.Controls.Add((New-AppLabel -Text "FLUXO" -X 25 -Y 155 -Width 300 -Height 28 -FontSize 12 -Bold $true -ForeColor $script:ColorMenu))

    $positionY = 190

    foreach ($step in (Get-DocumentSteps -DocumentId $DocumentId)) {
        $stepUser = Get-UserById -UserId ([int]$step.responsavel_id)
        $stepUserName = "-"

        if ($null -ne $stepUser) {
            $stepUserName = $stepUser.nome
        }

        $symbol = "○"
        $detail = "Aguardando"

        if ($step.status -eq "CONCLUIDA") {
            $symbol = "✓"
            $detail = "Concluída: $($step.concluido_em)"
        }
        elseif ($step.status -eq "PENDENTE") {
            $symbol = "●"
            $detail = "Aguardando ação"
        }
        elseif ($step.status -eq "RECUSADA") {
            $symbol = "✕"
            $detail = "Recusada: $($step.concluido_em)"
        }

        $script:ContentPanel.Controls.Add((New-AppLabel -Text $symbol -X 40 -Y $positionY -Width 35 -Height 30 -FontSize 16 -Bold $true -ForeColor (Get-StatusColor -Status $step.status)))

        $stepText = "$($step.ordem). $(Get-StepText -Type $step.tipo) - $stepUserName`r`n$detail"

        $script:ContentPanel.Controls.Add((New-AppLabel -Text $stepText -X 80 -Y $positionY -Width 350 -Height 42 -FontSize 10 -Bold $true -ForeColor $script:ColorMenu))

        $positionY += 55
    }

    $script:ContentPanel.Controls.Add((New-AppLabel -Text "PDF ORIGINAL" -X 470 -Y 155 -Width 300 -Height 28 -FontSize 12 -Bold $true -ForeColor $script:ColorMenu))
    $script:ContentPanel.Controls.Add((New-AppLabel -Text "Arquivo: $([System.IO.Path]::GetFileName($document.arquivo_pdf))" -X 470 -Y 190 -Width 400 -Height 25 -FontSize 9 -ForeColor $script:ColorMuted))

    $openPdfButton = New-AppButton -Text "ABRIR PDF" -X 470 -Y 220 -Width 180 -Height 38 -BackColor $script:ColorPrimary
    $openPdfButton.Tag = $document.arquivo_pdf

    $openPdfButton.Add_Click({
        try {
            $pdfPath = [string]$this.Tag

            if (-not (Test-Path -Path $pdfPath)) {
                Show-AppMessage -Message "PDF não encontrado." -Icon Warning
                return
            }

            Start-Process -FilePath $pdfPath
        }
        catch {
            Show-AppMessage -Message $_.Exception.Message -Icon Error
        }
    })

    $script:ContentPanel.Controls.Add($openPdfButton)

    $canAct = $false

    if (
        $document.status -eq "EM_APROVACAO" -and
        $null -ne $currentStep -and
        $currentStep.status -eq "PENDENTE" -and
        [int]$currentStep.responsavel_id -eq [int]$script:CurrentUser.id
    ) {
        $canAct = $true
    }

    if ($canAct) {
        $action = "APROVAR"
        $actionText = "APROVAR"

        if ($currentStep.tipo -eq "VERIFICACAO") {
            $action = "CONFIRMAR_VERIFICACAO"
            $actionText = "CONFIRMAR VERIFICAÇÃO"
        }

        $approveButton = New-AppButton -Text $actionText -X 470 -Y 280 -Width 240 -Height 42 -BackColor $script:ColorSuccess

        $approveButton.Tag = [pscustomobject]@{
            DocumentId = $DocumentId
            Action = $action
            Display = $actionText
        }

        $approveButton.Add_Click({
            try {
                $context = $this.Tag

                if (-not (Confirm-AppMessage -Message "Confirma a ação: $($context.Display)?")) {
                    return
                }

                $message = Process-WorkflowAction `
                    -DocumentId ([int]$context.DocumentId) `
                    -UserId ([int]$script:CurrentUser.id) `
                    -Action ([string]$context.Action)

                Show-AppMessage -Message $message -Icon Information
                Show-DocumentDetails -DocumentId ([int]$context.DocumentId)
            }
            catch {
                Show-AppMessage -Message $_.Exception.Message -Icon Error
            }
        })

        $script:ContentPanel.Controls.Add($approveButton)

        if ($currentStep.tipo -eq "APROVACAO") {
            $rejectButton = New-AppButton -Text "RECUSAR" -X 725 -Y 280 -Width 160 -Height 42 -BackColor $script:ColorDanger
            $rejectButton.Tag = $DocumentId

            $rejectButton.Add_Click({
                Show-RejectDialog -DocumentId ([int]$this.Tag)
            })

            $script:ContentPanel.Controls.Add($rejectButton)
        }
    }

    $script:ContentPanel.Controls.Add((New-AppLabel -Text "HISTÓRICO" -X 470 -Y 360 -Width 300 -Height 28 -FontSize 12 -Bold $true -ForeColor $script:ColorMenu))

    $historyBox = New-Object System.Windows.Forms.TextBox
    $historyBox.Location = New-Object System.Drawing.Point(470, 395)
    $historyBox.Size = New-Object System.Drawing.Size(415, 180)
    $historyBox.Multiline = $true
    $historyBox.ReadOnly = $true
    $historyBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $historyBox.Font = New-Object System.Drawing.Font("Consolas", 8)
    $historyBox.BackColor = $script:ColorWhite

    foreach ($entry in (Get-History)) {
        if ([int]$entry.documento_id -eq $DocumentId) {
            $historyUser = Get-UserById -UserId ([int]$entry.usuario_id)
            $historyUserName = "Sistema"

            if ($null -ne $historyUser) {
                $historyUserName = $historyUser.nome
            }

            $historyBox.AppendText("$($entry.criado_em) | $historyUserName`r`n$($entry.descricao)`r`n`r`n")
        }
    }

    $backButton = New-AppButton -Text "VOLTAR" -X 25 -Y 580 -Width 150 -Height 36 -BackColor $script:ColorMuted
    $backButton.Add_Click({ Show-Dashboard })

    $script:ContentPanel.Controls.Add($historyBox)
    $script:ContentPanel.Controls.Add($backButton)
}

# =============================================================================
# RECUSA
# =============================================================================

function Show-RejectDialog {
    param([int]$DocumentId)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Recusar documento"
    $form.Size = New-Object System.Drawing.Size(470, 300)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $form.BackColor = $script:ColorBackground
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $form.Controls.Add((New-AppLabel -Text "MOTIVO DA RECUSA" -X 25 -Y 20 -Width 350 -Height 30 -FontSize 13 -Bold $true -ForeColor $script:ColorDanger))
    $form.Controls.Add((New-AppLabel -Text "O motivo é obrigatório." -X 25 -Y 55 -Width 350 -Height 25 -FontSize 9 -ForeColor $script:ColorMuted))

    $reasonBox = New-Object System.Windows.Forms.TextBox
    $reasonBox.Location = New-Object System.Drawing.Point(25, 85)
    $reasonBox.Size = New-Object System.Drawing.Size(400, 100)
    $reasonBox.Multiline = $true
    $reasonBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical

    $confirmButton = New-AppButton -Text "CONFIRMAR RECUSA" -X 25 -Y 210 -Width 190 -Height 38 -BackColor $script:ColorDanger

    $confirmButton.Tag = [pscustomobject]@{
        Form = $form
        DocumentId = $DocumentId
        ReasonBox = $reasonBox
    }

    $confirmButton.Add_Click({
        try {
            $context = $this.Tag
            $reason = $context.ReasonBox.Text.Trim()

            if ([string]::IsNullOrWhiteSpace($reason)) {
                Show-AppMessage -Message "Informe o motivo da recusa." -Icon Warning
                return
            }

            $message = Process-WorkflowAction `
                -DocumentId ([int]$context.DocumentId) `
                -UserId ([int]$script:CurrentUser.id) `
                -Action "RECUSAR" `
                -Reason $reason

            $context.Form.Close()
            Show-AppMessage -Message $message -Icon Information
            Show-DocumentDetails -DocumentId ([int]$context.DocumentId)
        }
        catch {
            Show-AppMessage -Message $_.Exception.Message -Icon Error
        }
    })

    $cancelButton = New-AppButton -Text "CANCELAR" -X 235 -Y 210 -Width 190 -Height 38 -BackColor $script:ColorMuted
    $cancelButton.Tag = $form

    $cancelButton.Add_Click({
        $this.Tag.Close()
    })

    $form.Controls.Add($reasonBox)
    $form.Controls.Add($confirmButton)
    $form.Controls.Add($cancelButton)

    [void]$form.ShowDialog()
}

# =============================================================================
# DOCUMENTOS
# =============================================================================

function Show-Documents {
    $script:CurrentPage = "DOCUMENTOS"
    $script:CurrentDocumentId = $null

    Clear-Content
    Update-NotificationCounter

    $page = New-Object System.Windows.Forms.Panel
    $page.Dock = [System.Windows.Forms.DockStyle]::Fill
    $page.BackColor = $script:ColorBackground

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = [System.Windows.Forms.DockStyle]::Top
    $header.Height = 64
    $header.BackColor = $script:ColorBackground
    $header.Controls.Add((New-AppLabel -Text "Documentos" -X 24 -Y 16 -Width 500 -Height 34 -FontSize 18 -Bold $true -ForeColor $script:ColorText))

    $footer = New-Object System.Windows.Forms.Panel
    $footer.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footer.Height = 64
    $footer.BackColor = $script:ColorBackground

    $gridHost = New-Object System.Windows.Forms.Panel
    $gridHost.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridHost.Padding = New-Object System.Windows.Forms.Padding(24, 0, 24, 8)
    $gridHost.BackColor = $script:ColorBackground

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $false
    Set-AppGridLook -Grid $grid

    foreach ($columnName in @("ID", "Número", "Pessoa", "Data", "Status", "Responsável")) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.Name = $columnName
        $column.HeaderText = $columnName
        [void]$grid.Columns.Add($column)
    }

    foreach ($document in (Get-Documents)) {
        $currentStep = Get-CurrentStep -DocumentId ([int]$document.id)
        $responsibleName = "-"

        if ($null -ne $currentStep) {
            $responsible = Get-UserById -UserId ([int]$currentStep.responsavel_id)

            if ($null -ne $responsible) {
                $responsibleName = $responsible.nome
            }
        }

        [void]$grid.Rows.Add(
            $document.id,
            $document.numero,
            $document.pessoa_relacionada,
            $document.data_documento,
            $document.status,
            $responsibleName
        )
    }

    $openButton = New-AppButton -Text "Abrir documento" -X 24 -Y 12 -Width 180 -Height 38 -BackColor $script:ColorPrimary
    $openButton.Tag = $grid

    $openButton.Add_Click({
        try {
            $selectedGrid = $this.Tag

            if ($selectedGrid.SelectedRows.Count -eq 0) {
                Show-AppMessage -Message "Selecione um documento." -Icon Warning
                return
            }

            Show-DocumentDetails -DocumentId ([int]$selectedGrid.SelectedRows[0].Cells["ID"].Value)
        }
        catch {
            Show-AppMessage -Message $_.Exception.Message -Icon Error
        }
    })

    $footer.Controls.Add($openButton)
    $gridHost.Controls.Add($grid)
    $page.Controls.Add($gridHost)
    $page.Controls.Add($footer)
    $page.Controls.Add($header)
    $script:ContentPanel.Controls.Add($page)
}

# =============================================================================
# NOTIFICAÇÕES
# =============================================================================

function Show-Notifications {
    $script:CurrentPage = "NOTIFICACOES"
    $script:CurrentDocumentId = $null

    Clear-Content
    Update-NotificationCounter

    $page = New-Object System.Windows.Forms.Panel
    $page.Dock = [System.Windows.Forms.DockStyle]::Fill
    $page.BackColor = $script:ColorBackground

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = [System.Windows.Forms.DockStyle]::Top
    $header.Height = 64
    $header.BackColor = $script:ColorBackground
    $header.Controls.Add((New-AppLabel -Text "Notificações" -X 24 -Y 16 -Width 500 -Height 34 -FontSize 18 -Bold $true -ForeColor $script:ColorText))

    $footer = New-Object System.Windows.Forms.Panel
    $footer.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footer.Height = 64
    $footer.BackColor = $script:ColorBackground

    $gridHost = New-Object System.Windows.Forms.Panel
    $gridHost.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridHost.Padding = New-Object System.Windows.Forms.Padding(24, 0, 24, 8)
    $gridHost.BackColor = $script:ColorBackground

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $false
    Set-AppGridLook -Grid $grid

    foreach ($columnName in @("ID", "Documento", "Mensagem", "Lida", "Data")) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.Name = $columnName
        $column.HeaderText = $columnName
        [void]$grid.Columns.Add($column)
    }

    foreach ($notification in (Get-Notifications)) {
        if ([int]$notification.usuario_id -eq [int]$script:CurrentUser.id) {
            $readText = "Não"

            if ((Convert-ToBoolean -Value $notification.lida) -eq $true) {
                $readText = "Sim"
            }

            [void]$grid.Rows.Add(
                $notification.id,
                $notification.documento_id,
                $notification.mensagem,
                $readText,
                $notification.criado_em
            )
        }
    }

    $openButton = New-AppButton -Text "Abrir documento" -X 24 -Y 12 -Width 180 -Height 38 -BackColor $script:ColorPrimary
    $openButton.Tag = $grid

    $openButton.Add_Click({
        try {
            $selectedGrid = $this.Tag

            if ($selectedGrid.SelectedRows.Count -eq 0) {
                Show-AppMessage -Message "Selecione uma notificação." -Icon Warning
                return
            }

            $notificationId = [int]$selectedGrid.SelectedRows[0].Cells["ID"].Value
            $documentId = [int]$selectedGrid.SelectedRows[0].Cells["Documento"].Value

            $notifications = Get-Notifications

            foreach ($notification in $notifications) {
                if ([int]$notification.id -eq $notificationId) {
                    $notification.lida = "true"
                }
            }

            Save-CsvRows -Rows $notifications -FilePath $script:NotificationsFile

            Update-NotificationCounter
            Show-DocumentDetails -DocumentId $documentId
        }
        catch {
            Show-AppMessage -Message $_.Exception.Message -Icon Error
        }
    })

    $footer.Controls.Add($openButton)
    $gridHost.Controls.Add($grid)
    $page.Controls.Add($gridHost)
    $page.Controls.Add($footer)
    $page.Controls.Add($header)
    $script:ContentPanel.Controls.Add($page)
}

# =============================================================================
# JANELA PRINCIPAL
# =============================================================================

function Refresh-CurrentPage {
    try {
        $currentPage = $script:CurrentPage

        if ([string]::IsNullOrWhiteSpace($currentPage)) {
            $currentPage = "DASHBOARD"
        }

        Update-NotificationCounter

        switch ($currentPage) {
            "DASHBOARD" {
                Show-Dashboard
            }

            "DOCUMENTOS" {
                Show-Documents
            }

            "NOTIFICACOES" {
                Show-Notifications
            }

            "CADASTRO_DOCUMENTO" {
                Show-RegisterDocument
            }

            "DETALHES_DOCUMENTO" {
                if ($null -ne $script:CurrentDocumentId) {
                    Show-DocumentDetails `
                        -DocumentId ([int]$script:CurrentDocumentId)
                }
                else {
                    Show-Dashboard
                }
            }

            # Novas abas administrativas
            "GERENCIAR_USUARIOS" {
                Show-UserManagement
            }

            "EXCLUSAO_ARQUIVOS" {
                Show-FileDeletion
            }

            "LIXEIRA" {
                Show-RecycleBin
            }

            default {
                Show-Dashboard
            }
        }
    }
    catch {
        Show-AppMessage `
            -Message "Não foi possível atualizar as informações.`r`n$($_.Exception.Message)" `
            -Icon Error
    }
}

# =============================================================================
# GERENCIAR UTILIZADORES (MÓDULO ADMIN)
# =============================================================================

function Show-UserManagement {
    $script:CurrentPage = "GERENCIAR_USUARIOS"
    $script:CurrentDocumentId = $null

    if ($script:CurrentUser.cargo -ne "ADMINISTRADOR") {
        Show-AppMessage -Message "Somente administradores podem gerenciar usuários." -Icon Warning
        return
    }

    Clear-Content

    $script:ContentPanel.Controls.Add((New-AppLabel -Text "GERENCIAR UTILIZADORES" -X 25 -Y 20 -Width 600 -Height 35 -FontSize 18 -Bold $true -ForeColor $script:ColorMenu))

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(25, 75)
    $grid.Size = New-Object System.Drawing.Size(850, 230)
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $false
    Set-AppGridLook -Grid $grid

    foreach ($colName in @("ID", "Nome", "E-mail", "Código", "Cargo", "Ativo")) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $colName
        $col.HeaderText = $colName
        [void]$grid.Columns.Add($col)
    }

    foreach ($u in (Get-Users)) {
        $ativoTxt = if ((Convert-ToBoolean -Value $u.ativo)) { "Sim" } else { "Não" }
        [void]$grid.Rows.Add($u.id, $u.nome, $u.email, $u.codigo, $u.cargo, $ativoTxt)
    }

    foreach ($u in (Get-Users)) {
    $ativoTxt = if ((Convert-ToBoolean -Value $u.ativo)) { "Sim" } else { "Não" }
    [void]$grid.Rows.Add($u.id, $u.nome, $u.email, $u.codigo, $u.cargo, $ativoTxt)
}

$script:SelectedUserId = $null

$grid.Add_CellClick({
    param($sender, $e)

    if ($e.RowIndex -lt 0) {
        return
    }

    $script:SelectedUserId = [int]$sender.Rows[$e.RowIndex].Cells["ID"].Value
})

$script:ContentPanel.Controls.Add($grid)

    $script:ContentPanel.Controls.Add($grid)

    $script:ContentPanel.Controls.Add((New-AppLabel -Text "ID:" -X 25 -Y 325 -Width 60 -Height 22 -FontSize 9 -Bold $true))
    $idBox = New-Object System.Windows.Forms.TextBox
    $idBox.Location = New-Object System.Drawing.Point(25, 348)
    $idBox.Size = New-Object System.Drawing.Size(60, 28)
    $idBox.ReadOnly = $true
    $script:ContentPanel.Controls.Add($idBox)

    $script:ContentPanel.Controls.Add((New-AppLabel -Text "Nome:" -X 100 -Y 325 -Width 160 -Height 22 -FontSize 9 -Bold $true))
    $nameBox = New-Object System.Windows.Forms.TextBox
    $nameBox.Location = New-Object System.Drawing.Point(100, 348)
    $nameBox.Size = New-Object System.Drawing.Size(160, 28)
    $script:ContentPanel.Controls.Add($nameBox)

    $script:ContentPanel.Controls.Add((New-AppLabel -Text "E-mail:" -X 275 -Y 325 -Width 220 -Height 22 -FontSize 9 -Bold $true))
    $emailBox = New-Object System.Windows.Forms.TextBox
    $emailBox.Location = New-Object System.Drawing.Point(275, 348)
    $emailBox.Size = New-Object System.Drawing.Size(220, 28)
    $script:ContentPanel.Controls.Add($emailBox)

    $script:ContentPanel.Controls.Add((New-AppLabel -Text "Código:" -X 510 -Y 325 -Width 120 -Height 22 -FontSize 9 -Bold $true))
    $codeBox = New-Object System.Windows.Forms.TextBox
    $codeBox.Location = New-Object System.Drawing.Point(510, 348)
    $codeBox.Size = New-Object System.Drawing.Size(120, 28)
    $script:ContentPanel.Controls.Add($codeBox)

    $script:ContentPanel.Controls.Add((New-AppLabel -Text "Cargo:" -X 25 -Y 395 -Width 150 -Height 22 -FontSize 9 -Bold $true))
    $roleCombo = New-Object System.Windows.Forms.ComboBox
    $roleCombo.Location = New-Object System.Drawing.Point(25, 418)
    $roleCombo.Size = New-Object System.Drawing.Size(180, 28)
    $roleCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($r in @("ADMINISTRADOR", "TERCEIRO", "LIDER", "SUPERVISOR", "COORDENADOR")) { [void]$roleCombo.Items.Add($r) }
    $script:ContentPanel.Controls.Add($roleCombo)

    $script:ContentPanel.Controls.Add((New-AppLabel -Text "Ativo:" -X 220 -Y 395 -Width 80 -Height 22 -FontSize 9 -Bold $true))
    $activeCheck = New-Object System.Windows.Forms.CheckBox
    $activeCheck.Location = New-Object System.Drawing.Point(220, 420)
    $activeCheck.Size = New-Object System.Drawing.Size(80, 24)
    $activeCheck.Checked = $true
    $activeCheck.Text = "Sim"
    $script:ContentPanel.Controls.Add($activeCheck)

    $grid.Add_SelectionChanged({
    if ($grid.SelectedRows.Count -gt 0) {
        $row = $grid.SelectedRows[0]

        $idBox.Text = [string]$row.Cells["ID"].Value
        $nameBox.Text = [string]$row.Cells["Nome"].Value
        $emailBox.Text = [string]$row.Cells["E-mail"].Value
        $codeBox.Text = [string]$row.Cells["Código"].Value
        $roleCombo.SelectedItem = [string]$row.Cells["Cargo"].Value
        $activeCheck.Checked = ([string]$row.Cells["Ativo"].Value -eq "Sim")
    }
})

   $clearBtn = New-AppButton -Text "LIMPAR" -X 315 -Y 410 -Width 110 -Height 38 -BackColor $script:ColorMuted

$clearBtn.Tag = [pscustomobject]@{
    IdBox       = $idBox
    NameBox     = $nameBox
    EmailBox    = $emailBox
    CodeBox     = $codeBox
    RoleCombo   = $roleCombo
    ActiveCheck = $activeCheck
    Grid        = $grid
}

$clearBtn.Add_Click({
    try {
        $script:SelectedUserId = $null
        $ctx = $this.Tag

        $ctx.IdBox.Clear()
        $ctx.NameBox.Clear()
        $ctx.EmailBox.Clear()
        $ctx.CodeBox.Clear()
        $ctx.RoleCombo.SelectedIndex = -1
        $ctx.ActiveCheck.Checked = $true
        $ctx.Grid.ClearSelection()
    }
    catch {
        Show-AppMessage -Message "Erro ao limpar campos: $($_.Exception.Message)" -Icon Error
    }
})

$script:ContentPanel.Controls.Add($clearBtn)
    $editUserBtn = New-AppButton `
    -Text "EDITAR SELECIONADO" `
    -X 25 -Y 465 -Width 180 -Height 38 `
    -BackColor $script:ColorPrimary

    $editUserBtn.Tag = [pscustomobject]@{
    IdBox       = $idBox
    NameBox     = $nameBox
    EmailBox    = $emailBox
    CodeBox     = $codeBox
    RoleCombo   = $roleCombo
    ActiveCheck = $activeCheck
}

    $editUserBtn.Add_Click({
    try {
        if ($null -eq $script:SelectedUserId) {
            Show-AppMessage -Message "Clique em um usuário da tabela antes de editar." -Icon Warning
            return
        }

        $ctx = $this.Tag
        if ($null -eq $ctx) {
            throw "O botão Editar está sem os controles no Tag."
        }

        $selectedUser = Get-UserById -UserId ([int]$script:SelectedUserId)
        if ($null -eq $selectedUser) {
            throw "Usuário não encontrado. Atualize a lista."
        }

        foreach ($campo in @('IdBox', 'NameBox', 'EmailBox', 'CodeBox')) {
            $controle = $ctx.$campo

            if ($null -eq $controle -or
                $controle -isnot [System.Windows.Forms.TextBox]) {
                $tipo = if ($null -eq $controle) {
                    'nulo'
                } else {
                    $controle.GetType().FullName
                }

                throw "Controle $campo inválido: $tipo"
            }
        }

        $ctx.IdBox.Text    = [string]$selectedUser.id
        $ctx.NameBox.Text  = [string]$selectedUser.nome
        $ctx.EmailBox.Text = [string]$selectedUser.email
        $ctx.CodeBox.Text  = [string]$selectedUser.codigo

        $ctx.RoleCombo.SelectedItem = [string]$selectedUser.cargo
        $ctx.ActiveCheck.Checked =
            (Convert-ToBoolean -Value $selectedUser.ativo)

        $ctx.NameBox.Focus()
    }
    catch {
        Show-AppMessage `
            -Message "Erro ao editar usuário: $($_.Exception.Message)" `
            -Icon Error
    }
})

$script:ContentPanel.Controls.Add($editUserBtn)


    $saveUserBtn = New-AppButton -Text "SALVAR / ATUALIZAR" -X 435 -Y 410 -Width 170 -Height 38 -BackColor $script:ColorSuccess
    $saveUserBtn.Tag = [pscustomobject]@{
    IdBox = $idBox
    NameBox = $nameBox
    EmailBox = $emailBox
    CodeBox = $codeBox
    RoleCombo = $roleCombo
    ActiveCheck = $activeCheck
}
    $saveUserBtn.Add_Click({
        try {
            $ctx = $this.Tag
            $n = $ctx.NameBox.Text.Trim()
            $e = $ctx.EmailBox.Text.Trim()
            $c = $ctx.CodeBox.Text.Trim()
            $cargoSel = $ctx.RoleCombo.SelectedItem
            $at = if ($ctx.ActiveCheck.Checked) { "true" } else { "false" }

            if ([string]::IsNullOrWhiteSpace($n) -or [string]::IsNullOrWhiteSpace($e) -or [string]::IsNullOrWhiteSpace($c) -or $null -eq $cargoSel) {
                Show-AppMessage -Message "Preencha todos os campos do usuário." -Icon Warning
                return
            }

            $usersList = Get-Users

            if ([string]::IsNullOrWhiteSpace($ctx.IdBox.Text)) {
                $newId = Get-NextId -Rows $usersList
                $newU = [pscustomobject]@{
                    id = $newId; nome = $n; email = $e; codigo = $c
                    cargo = $cargoSel; ativo = $at; criado_em = (Get-NowText)
                }
                $usersList = @($usersList) + @($newU)
                Show-AppMessage -Message "Usuário cadastrado com sucesso!" -Icon Information
            } else {
                $targetId = [int]$ctx.IdBox.Text
                foreach ( $userObj in $usersList ) {
                    if ([int]$userObj.id -eq $targetId) {
                        $userObj.nome = $n; $userObj.email = $e; $userObj.codigo = $c
                        $userObj.cargo = $cargoSel; $userObj.ativo = $at
                    }
                }
                Show-AppMessage -Message "Usuário atualizado com sucesso!" -Icon Information
            }

            Save-CsvRows -Rows $usersList -FilePath $script:UsersFile
            Show-UserManagement
        } catch {
            Show-AppMessage -Message $_.Exception.Message -Icon Error
        }
    })

    $script:ContentPanel.Controls.Add($saveUserBtn)

    $deleteUserBtn = New-AppButton -Text "EXCLUIR SELECIONADO" -X 615 -Y 410 -Width 180 -Height 38 -BackColor $script:ColorDanger
    $deleteUserBtn.Tag = $grid
    $deleteUserBtn.Add_Click({
        try {
            $targetGrid = $this.Tag
            if ($targetGrid.SelectedRows.Count -eq 0) {
                Show-AppMessage -Message "Selecione um usuário na tabela para excluir." -Icon Warning
                return
            }

            if (-not (Confirm-AppMessage -Message "Tem certeza que deseja excluir este usuário?")) {
                return
            }

            $targetId = [int]$targetGrid.SelectedRows[0].Cells["ID"].Value
            $usersList = Get-Users
            $newUsersList = @()

            foreach ($u in $usersList) {
                if ([int]$u.id -ne $targetId) {
                    $newUsersList += $u
                }
            }

            Save-CsvRows -Rows $newUsersList -FilePath $script:UsersFile
            Show-AppMessage -Message "Usuário excluído com sucesso!" -Icon Information
            Show-UserManagement
        } catch {
            Show-AppMessage -Message $_.Exception.Message -Icon Error
        }
    })
    $script:ContentPanel.Controls.Add($deleteUserBtn)
    }



    function Get-DeletedDocuments {
    return @(Get-CsvRows -FilePath $script:DeletedDocumentsFile)
}

function Get-RecycleBinItems {
    return @(Get-CsvRows -FilePath $script:RecycleBinFile)
}

function Convert-ToDateSafe {
    param([string]$Value)

    try {
        return [datetime]::Parse($Value)
    }
    catch {
        return $null
    }
}

function Get-RecycleBinDaysRemaining {
    param([string]$PermanentDeletionDate)

    $date = Convert-ToDateSafe -Value $PermanentDeletionDate

    if ($null -eq $date) {
        return -1
    }

    return ($date.Date - (Get-Date).Date).Days
}

function Show-UserRegistration {
    if ($script:CurrentUser.cargo -ne "ADMINISTRADOR") {
        Show-AppMessage -Message "Apenas administradores podem cadastrar usuários." -Icon Warning
        return
    }

    $script:CurrentPage = "USUARIOS"
    $script:CurrentDocumentId = $null

    Clear-Content
    Update-NotificationCounter

    $script:ContentPanel.Controls.Add(
        (New-AppLabel `
            -Text "CADASTRO DE USUÁRIOS" `
            -X 25 -Y 20 -Width 500 -Height 35 `
            -FontSize 18 -Bold $true `
            -ForeColor $script:ColorMenu)
    )

    $script:ContentPanel.Controls.Add(
        (New-AppLabel `
            -Text "Cadastre os responsáveis pelas etapas do workflow." `
            -X 25 -Y 60 -Width 700 -Height 25 `
            -FontSize 10 `
            -ForeColor $script:ColorMuted)
    )

    $fields = @(
        @{ Text = "Nome"; Y = 115 },
        @{ Text = "E-mail"; Y = 175 },
        @{ Text = "Código de acesso"; Y = 235 },
        @{ Text = "Cargo"; Y = 295 }
    )

    foreach ($field in $fields) {
        $script:ContentPanel.Controls.Add(
            (New-AppLabel `
                -Text $field.Text `
                -X 25 -Y $field.Y -Width 210 -Height 25 `
                -FontSize 10 -Bold $true)
        )
    }

    $nameBox = New-Object System.Windows.Forms.TextBox
    $nameBox.Location = New-Object System.Drawing.Point(245, 112)
    $nameBox.Size = New-Object System.Drawing.Size(350, 28)

    $emailBox = New-Object System.Windows.Forms.TextBox
    $emailBox.Location = New-Object System.Drawing.Point(245, 172)
    $emailBox.Size = New-Object System.Drawing.Size(350, 28)

    $codeBox = New-Object System.Windows.Forms.TextBox
    $codeBox.Location = New-Object System.Drawing.Point(245, 232)
    $codeBox.Size = New-Object System.Drawing.Size(350, 28)
    $codeBox.UseSystemPasswordChar = $true

    $roleCombo = New-Object System.Windows.Forms.ComboBox
    $roleCombo.Location = New-Object System.Drawing.Point(245, 292)
    $roleCombo.Size = New-Object System.Drawing.Size(350, 28)
    $roleCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    foreach ($role in @(
        "ADMINISTRADOR",
        "TERCEIRO",
        "LIDER",
        "SUPERVISOR",
        "COORDENADOR"
    )) {
        [void]$roleCombo.Items.Add($role)
    }

    $roleCombo.SelectedIndex = 1

    $saveButton = New-AppButton `
        -Text "CADASTRAR USUÁRIO" `
        -X 245 -Y 355 -Width 220 -Height 40 `
        -BackColor $script:ColorSuccess

    $saveButton.Tag = [pscustomobject]@{
        NameBox   = $nameBox
        EmailBox  = $emailBox
        CodeBox   = $codeBox
        RoleCombo = $roleCombo
    }

    $saveButton.Add_Click({
        try {
            $context = $this.Tag

            $name = $context.NameBox.Text.Trim()
            $email = $context.EmailBox.Text.Trim().ToLowerInvariant()
            $code = $context.CodeBox.Text
            $role = [string]$context.RoleCombo.SelectedItem

            if ([string]::IsNullOrWhiteSpace($name)) {
                Show-AppMessage -Message "Informe o nome do usuário." -Icon Warning
                return
            }

            if ([string]::IsNullOrWhiteSpace($email)) {
                Show-AppMessage -Message "Informe o e-mail do usuário." -Icon Warning
                return
            }

            if ([string]::IsNullOrWhiteSpace($code)) {
                Show-AppMessage -Message "Informe o código de acesso." -Icon Warning
                return
            }

            $users = Get-Users

            foreach ($existingUser in $users) {
                if ($existingUser.email.ToLowerInvariant() -eq $email) {
                    Show-AppMessage -Message "Já existe usuário com este e-mail." -Icon Warning
                    return
                }
            }

            $newUser = [pscustomobject]@{
                id        = Get-NextId -Rows $users
                nome      = $name
                email     = $email
                codigo    = $code
                cargo     = $role
                ativo     = "true"
                criado_em = Get-NowText
            }

            $users = @($users) + @($newUser)

            Save-CsvRows `
                -Rows $users `
                -FilePath $script:UsersFile

            Show-AppMessage -Message "Usuário cadastrado com sucesso." -Icon Information

            Show-UserRegistration
        }
        catch {
            Show-AppMessage -Message $_.Exception.Message -Icon Error
        }
    })

    $usersGrid = New-Object System.Windows.Forms.DataGridView
    $usersGrid.Location = New-Object System.Drawing.Point(25, 435)
    $usersGrid.Size = New-Object System.Drawing.Size(850, 180)
    $usersGrid.ReadOnly = $true
    $usersGrid.AllowUserToAddRows = $false
    $usersGrid.AllowUserToDeleteRows = $false
    $usersGrid.RowHeadersVisible = $false
    $usersGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $usersGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $usersGrid.MultiSelect = $false
    Set-AppGridLook -Grid $usersGrid

    foreach ($columnName in @(
        "ID",
        "Nome",
        "E-mail",
        "Cargo",
        "Ativo",
        "Criado em"
    )) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.Name = $columnName
        $column.HeaderText = $columnName
        [void]$usersGrid.Columns.Add($column)
    }

    foreach ($user in (Get-Users)) {
        [void]$usersGrid.Rows.Add(
            $user.id,
            $user.nome,
            $user.email,
            $user.cargo,
            $user.ativo,
            $user.criado_em
        )
    }

    $script:ContentPanel.Controls.Add($nameBox)
    $script:ContentPanel.Controls.Add($emailBox)
    $script:ContentPanel.Controls.Add($codeBox)
    $script:ContentPanel.Controls.Add($roleCombo)
    $script:ContentPanel.Controls.Add($saveButton)
    $script:ContentPanel.Controls.Add($usersGrid)
}

function Move-DocumentToRecycleBin {
    param(
        [int]$DocumentId,
        [string]$Reason
    )

    $document = Get-DocumentById -DocumentId $DocumentId

    if ($null -eq $document) {
        throw "Documento não encontrado."
    }

    $documents = Get-Documents
    $deletedDocuments = Get-DeletedDocuments
    $recycleItems = Get-RecycleBinItems

    $now = Get-NowText
    $permanentDeletion = (Get-Date).AddDays(30).ToString("yyyy-MM-dd HH:mm:ss")

    $originalPdf = $document.arquivo_pdf
    $recyclePdf = ""

    if (
        -not [string]::IsNullOrWhiteSpace($originalPdf) -and
        (Test-Path -Path $originalPdf)
    ) {
        $fileName = "{0}_{1}_{2}" -f `
            $document.numero, `
            $DocumentId, `
            [System.IO.Path]::GetFileName($originalPdf)

        $recyclePdf = Join-Path $script:RecycleBinPath $fileName

        Move-Item `
            -Path $originalPdf `
            -Destination $recyclePdf `
            -Force
    }

    $deletedDocument = [pscustomobject]@{
        id                 = $document.id
        numero             = $document.numero
        pessoa_relacionada = $document.pessoa_relacionada
        data_documento     = $document.data_documento
        arquivo_pdf        = $document.arquivo_pdf
        status             = $document.status
        etapa_atual_id     = $document.etapa_atual_id
        criado_por         = $document.criado_por
        criado_em          = $document.criado_em
        finalizado_em      = $document.finalizado_em
        excluido_por       = $script:CurrentUser.nome
        excluido_em        = $now
        motivo_exclusao    = $Reason
    }

    $recycleEntry = [pscustomobject]@{
        id                         = Get-NextId -Rows $recycleItems
        documento_id               = $DocumentId
        numero                     = $document.numero
        arquivo_original           = $originalPdf
        arquivo_lixeira            = $recyclePdf
        excluido_por               = $script:CurrentUser.nome
        excluido_em                = $now
        excluir_definitivamente_em = $permanentDeletion
        motivo                     = $Reason
        status                     = "NA_LIXEIRA"
    }

    $remainingDocuments = @()

    foreach ($item in $documents) {
        if ([int]$item.id -ne $DocumentId) {
            $remainingDocuments += $item
        }
    }

    $deletedDocuments = @($deletedDocuments) + @($deletedDocument)
    $recycleItems = @($recycleItems) + @($recycleEntry)

    Save-CsvRows -Rows $remainingDocuments -FilePath $script:DocumentsFile
    Save-CsvRows -Rows $deletedDocuments -FilePath $script:DeletedDocumentsFile
    Save-CsvRows -Rows $recycleItems -FilePath $script:RecycleBinFile

    Add-History `
        -DocumentId $DocumentId `
        -UserId ([int]$script:CurrentUser.id) `
        -EventType "DOCUMENTO_ENVIADO_LIXEIRA" `
        -Description "Documento enviado para lixeira por $($script:CurrentUser.nome). Motivo: $Reason"
}

function Show-FileDeletion {
    $script:CurrentPage = "EXCLUSAO_ARQUIVOS"
    $script:CurrentDocumentId = $null

    if ($script:CurrentUser.cargo -ne "ADMINISTRADOR") {
        Show-AppMessage -Message "Apenas administradores podem excluir documentos." -Icon Warning
        return
    }

    $script:CurrentPage = "EXCLUSAO_ARQUIVOS"
    $script:CurrentDocumentId = $null

    Clear-Content
    Update-NotificationCounter

    $script:ContentPanel.Controls.Add(
        (New-AppLabel `
            -Text "EXCLUSÃO DE ARQUIVOS" `
            -X 25 -Y 20 -Width 550 -Height 35 `
            -FontSize 18 -Bold $true `
            -ForeColor $script:ColorDanger)
    )

    $script:ContentPanel.Controls.Add(
        (New-AppLabel `
            -Text "Os documentos enviados à lixeira podem ser restaurados por até 30 dias." `
            -X 25 -Y 60 -Width 800 -Height 25 `
            -FontSize 10 `
            -ForeColor $script:ColorMuted)
    )

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(25, 105)
    $grid.Size = New-Object System.Drawing.Size(850, 360)
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $false
    Set-AppGridLook -Grid $grid

    foreach ($columnName in @(
        "ID",
        "Número",
        "Pessoa",
        "Status",
        "Criado em",
        "Alerta"
    )) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.Name = $columnName
        $column.HeaderText = $columnName
        [void]$grid.Columns.Add($column)
    }

    foreach ($document in (Get-Documents)) {
        $alert = "Disponível para exclusão"

        $createdDate = Convert-ToDateSafe -Value $document.criado_em

        if ($null -ne $createdDate) {
            $daysOld = ((Get-Date) - $createdDate).Days

            if ($daysOld -ge 25) {
                $alert = "ATENÇÃO: documento com mais de 25 dias"
            }
        }

        [void]$grid.Rows.Add(
            $document.id,
            $document.numero,
            $document.pessoa_relacionada,
            $document.status,
            $document.criado_em,
            $alert
        )
    }

    $deleteButton = New-AppButton `
        -Text "MOVER PARA LIXEIRA" `
        -X 25 -Y 490 -Width 240 -Height 40 `
        -BackColor $script:ColorDanger

    $deleteButton.Tag = $grid

    $deleteButton.Add_Click({
        try {
            $selectedGrid = $this.Tag

            if ($selectedGrid.SelectedRows.Count -eq 0) {
                Show-AppMessage -Message "Selecione um documento." -Icon Warning
                return
            }

            $documentId = [int]$selectedGrid.SelectedRows[0].Cells["ID"].Value

            if (-not (Confirm-AppMessage -Message "Confirma mover este documento para a lixeira?")) {
                return
            }

            Move-DocumentToRecycleBin `
                -DocumentId $documentId `
                -Reason "Exclusão administrativa"

            Show-AppMessage -Message "Documento movido para a lixeira." -Icon Information

            Show-FileDeletion
        }
        catch {
            Show-AppMessage -Message $_.Exception.Message -Icon Error
        }
    })

    $recycleButton = New-AppButton `
        -Text "ABRIR LIXEIRA" `
        -X 280 -Y 490 -Width 190 -Height 40 `
        -BackColor $script:ColorMuted

    $recycleButton.Add_Click({
        Show-RecycleBin
    })

    $script:ContentPanel.Controls.Add($grid)
    $script:ContentPanel.Controls.Add($deleteButton)
    $script:ContentPanel.Controls.Add($recycleButton)
}

function Test-DocumentRestoreReady {
    param([int]$RecycleId)

    if (
        $null -eq $script:CurrentUser -or
        $script:CurrentUser.cargo -ne "ADMINISTRADOR"
    ) {
        throw "Apenas administradores podem verificar a restauração."
    }

    $entries = @(
        Get-RecycleBinItems |
            Where-Object { [int]$_.id -eq $RecycleId }
    )

    if ($entries.Count -ne 1) {
        throw "Registro ausente ou duplicado em lixeira.csv."
    }

    $entry = $entries[0]

    if ($entry.status -ne "NA_LIXEIRA") {
        throw "O item não está marcado como NA_LIXEIRA."
    }

    $deletedRecords = @(
        Get-DeletedDocuments |
            Where-Object { [int]$_.id -eq [int]$entry.documento_id }
    )

    if ($deletedRecords.Count -ne 1) {
        throw "Registro original ausente ou duplicado em documentos_excluidos.csv."
    }

    $deleted = $deletedRecords[0]

    if ([string]$deleted.numero -ne [string]$entry.numero) {
        throw "O número diverge entre os dois CSVs."
    }

    foreach ($active in (Get-Documents)) {
        if (
            [int]$active.id -eq [int]$deleted.id -or
            [string]$active.numero -eq [string]$deleted.numero
        ) {
            throw "Já existe documento ativo com o mesmo ID ou número."
        }
    }

    $originalPdf = [string]$entry.arquivo_original
    $recyclePdf = [string]$entry.arquivo_lixeira

    if (
        [string]::IsNullOrWhiteSpace($originalPdf) -or
        [string]::IsNullOrWhiteSpace($recyclePdf)
    ) {
        throw "Caminho do PDF ausente."
    }

    if (
        -not [string]::Equals(
            [string]$deleted.arquivo_pdf,
            $originalPdf,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        throw "Caminho do PDF diverge entre os CSVs."
    }

    $originalPdf = [System.IO.Path]::GetFullPath($originalPdf)
    $recyclePdf = [System.IO.Path]::GetFullPath($recyclePdf)

    $pdfRoot = [System.IO.Path]::GetFullPath($script:PdfPath).TrimEnd('\') + '\'
    $recycleRoot = [System.IO.Path]::GetFullPath($script:RecycleBinPath).TrimEnd('\') + '\'

    if (-not $originalPdf.StartsWith($pdfRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Caminho original fora da pasta PDFs."
    }

    if (-not $recyclePdf.StartsWith($recycleRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Caminho da lixeira fora da pasta Lixeira."
    }

    if (Test-Path -LiteralPath $originalPdf) {
        throw "Já existe arquivo no caminho original."
    }

    if (-not (Test-Path -LiteralPath $recyclePdf -PathType Leaf)) {
        throw "PDF não encontrado na lixeira."
    }

    return "Pronto para restaurar o documento $($entry.numero). Nenhum arquivo foi alterado."
}

function Restore-DocumentFromRecycleBin {
    param([int]$RecycleId)

    if (
        $null -eq $script:CurrentUser -or
        $script:CurrentUser.cargo -ne "ADMINISTRADOR"
    ) {
        throw "Somente administradores podem restaurar documentos."
    }

    # Lê novamente os três cadastros antes de agir.
    $recycleItems = @(Get-RecycleBinItems)
    $deletedDocuments = @(Get-DeletedDocuments)
    $activeDocuments = @(Get-Documents)

    $entry = @($recycleItems | Where-Object { [int]$_.id -eq $RecycleId })

    if ($entry.Count -ne 1) {
        throw "Registro da lixeira não encontrado ou duplicado."
    }

    $entry = $entry[0]

    if ($entry.status -ne "NA_LIXEIRA") {
        throw "O item não está disponível para restauração."
    }

    $deleted = @(
        $deletedDocuments |
            Where-Object { [int]$_.id -eq [int]$entry.documento_id }
    )

    if ($deleted.Count -ne 1) {
        throw "Registro original do documento não encontrado ou duplicado."
    }

    $deleted = $deleted[0]

    if ([string]$deleted.numero -ne [string]$entry.numero) {
        throw "Número do documento diverge entre os dois cadastros."
    }

    if (
        -not [string]::Equals(
            [string]$deleted.arquivo_pdf,
            [string]$entry.arquivo_original,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        throw "O caminho original do PDF diverge entre os dois cadastros."
    }

    foreach ($document in $activeDocuments) {
        if (
            [int]$document.id -eq [int]$deleted.id -or
            [string]$document.numero -eq [string]$deleted.numero
        ) {
            throw "Já existe documento ativo com o mesmo ID ou número. Restauração bloqueada."
        }
    }

    $originalPdf = [string]$entry.arquivo_original
    $recyclePdf = [string]$entry.arquivo_lixeira

    if (
        [string]::IsNullOrWhiteSpace($originalPdf) -or
        [string]::IsNullOrWhiteSpace($recyclePdf)
    ) {
        throw "Caminho do PDF ausente. Restauração bloqueada."
    }

    $originalPdf = [System.IO.Path]::GetFullPath($originalPdf)
    $recyclePdf = [System.IO.Path]::GetFullPath($recyclePdf)

    $pdfRoot = [System.IO.Path]::GetFullPath($script:PdfPath).TrimEnd('\') + '\'
    $recycleRoot = [System.IO.Path]::GetFullPath($script:RecycleBinPath).TrimEnd('\') + '\'

    if (
        -not $originalPdf.StartsWith(
            $pdfRoot,
            [System.StringComparison]::OrdinalIgnoreCase
        ) -or
        -not $recyclePdf.StartsWith(
            $recycleRoot,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        throw "Um dos caminhos está fora das pastas esperadas. Restauração bloqueada."
    }

    if (-not (Test-Path -LiteralPath $recyclePdf -PathType Leaf)) {
        throw "O PDF não foi encontrado na lixeira."
    }

    if (Test-Path -LiteralPath $originalPdf) {
        throw "Já existe arquivo no caminho original. Nenhum PDF foi sobrescrito."
    }

    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Restaurar o documento $($entry.numero) e seu PDF? As etapas anteriores serão preservadas.",
        "Confirmar restauração",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    # Monta um objeto com SOMENTE as colunas de documentos.csv.
    $restoredDocument = [pscustomobject]@{
        id                 = $deleted.id
        numero             = $deleted.numero
        pessoa_relacionada = $deleted.pessoa_relacionada
        data_documento     = $deleted.data_documento
        arquivo_pdf        = $deleted.arquivo_pdf
        status             = $deleted.status
        etapa_atual_id     = $deleted.etapa_atual_id
        criado_por         = $deleted.criado_por
        criado_em          = $deleted.criado_em
        finalizado_em      = $deleted.finalizado_em
    }

    # A cópia é verificada ANTES de retirar algo da lixeira.
    Copy-Item -LiteralPath $recyclePdf -Destination $originalPdf -ErrorAction Stop

    $sourceHash = Get-FileHash -LiteralPath $recyclePdf -Algorithm SHA256 -ErrorAction Stop
    $targetHash = Get-FileHash -LiteralPath $originalPdf -Algorithm SHA256 -ErrorAction Stop

    if ($sourceHash.Hash -ne $targetHash.Hash) {
        Remove-Item -LiteralPath $originalPdf -Force -ErrorAction SilentlyContinue
        throw "A cópia restaurada não corresponde ao PDF da lixeira. Os cadastros foram preservados."
    }

    $remainingDeleted = @(
        $deletedDocuments |
            Where-Object { [int]$_.id -ne [int]$deleted.id }
    )

    $updatedRecycle = @(
    foreach ($item in $recycleItems) {
        if ([int]$item.id -eq $RecycleId) {
            $item.status = "RESTAURADO"
        }

        $item
    }
)
    $newActive = @($activeDocuments) + @($restoredDocument)
    $lastConfirmed = "PDF copiado e verificado; nenhum CSV confirmado"

    try {
        # 1. Recoloca o documento no cadastro ativo.
        $newActive | Export-Csv -LiteralPath $script:DocumentsFile `
            -Delimiter ";" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop

        $activeCheck = @(
            Import-Csv -LiteralPath $script:DocumentsFile -Delimiter ";" -ErrorAction Stop |
                Where-Object {
                    [string]$_.id -eq [string]$deleted.id -and
                    [string]$_.numero -eq [string]$deleted.numero
                }
        )

        if ($activeCheck.Count -ne 1) {
            throw "A gravação de documentos.csv não foi confirmada por releitura."
        }

        $lastConfirmed = "documentos.csv confirmado"

        # 2. Retira o registro de documentos_excluidos.csv.
        if ($remainingDeleted.Count -gt 0) {
            $remainingDeleted | Export-Csv -LiteralPath $script:DeletedDocumentsFile `
                -Delimiter ";" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        }
        else {
            Set-Content -LiteralPath $script:DeletedDocumentsFile `
                -Value "id;numero;pessoa_relacionada;data_documento;arquivo_pdf;status;etapa_atual_id;criado_por;criado_em;finalizado_em;excluido_por;excluido_em;motivo_exclusao" `
                -Encoding UTF8 -ErrorAction Stop
        }

        $deletedCheck = @(
            Import-Csv -LiteralPath $script:DeletedDocumentsFile -Delimiter ";" -ErrorAction Stop |
                Where-Object { [string]$_.id -eq [string]$deleted.id }
        )

        if ($deletedCheck.Count -ne 0) {
            throw "A remoção do registro em documentos_excluidos.csv não foi confirmada."
        }

        $lastConfirmed = "documentos.csv e documentos_excluidos.csv confirmados"

        # 3. Retira o registro da lixeira.
       $updatedRecycle | Export-Csv `
    -LiteralPath $script:RecycleBinFile `
    -Delimiter ";" `
    -NoTypeInformation `
    -Encoding UTF8 `
    -ErrorAction Stop

$recycleCheck = @(
    Import-Csv -LiteralPath $script:RecycleBinFile -Delimiter ";" -ErrorAction Stop |
        Where-Object {
            [string]$_.id -eq [string]$RecycleId -and
            [string]$_.status -eq "RESTAURADO"
        }
)

if ($recycleCheck.Count -ne 1) {
    throw "O status RESTAURADO não foi confirmado em lixeira.csv."
}

        $lastConfirmed = "os três CSVs confirmados"

        if (-not (Test-Path -LiteralPath $originalPdf -PathType Leaf)) {
            throw "PDF restaurado não encontrado no caminho original."
        }

        # Preserva a cópia da lixeira: não há exclusão irreversível nesta rotina.
        return "Documento $($entry.numero) restaurado. Os três CSVs foram conferidos; a cópia de segurança permanece na pasta Lixeira."
    }
    catch {
        throw "Restauração interrompida. Última etapa confirmada: $lastConfirmed. Erro: $($_.Exception.Message). NÃO clique novamente; confira os três CSVs e os dois caminhos do PDF."
        }
    }

function Show-RecycleBin {
    $script:CurrentPage = "LIXEIRA"
    $script:CurrentDocumentId = $null

    if ($script:CurrentUser.cargo -ne "ADMINISTRADOR") {
        Show-AppMessage -Message "Apenas administradores podem acessar a lixeira." -Icon Warning
        return
    }

    $script:CurrentPage = "LIXEIRA"
    $script:CurrentDocumentId = $null

    Clear-Content
    Update-NotificationCounter

    $script:ContentPanel.Controls.Add(
        (New-AppLabel `
            -Text "LIXEIRA" `
            -X 25 -Y 20 -Width 500 -Height 35 `
            -FontSize 18 -Bold $true `
            -ForeColor $script:ColorDanger)
    )

    $script:ContentPanel.Controls.Add(
        (New-AppLabel `
            -Text "Itens podem ser restaurados enquanto estiverem dentro do prazo de 30 dias." `
            -X 25 -Y 60 -Width 800 -Height 25 `
            -FontSize 10 `
            -ForeColor $script:ColorMuted)
    )

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(25, 105)
    $grid.Size = New-Object System.Drawing.Size(850, 360)
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $false
    $grid.ReadOnly = $true
    $grid.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditProgrammatically

    Set-AppGridLook -Grid $grid

    foreach ($columnName in @(
        "ID",
        "Documento",
        "Excluído em",
        "Exclusão definitiva",
        "Dias restantes",
        "Status"
    )) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.Name = $columnName
        $column.HeaderText = $columnName
        [void]$grid.Columns.Add($column)
    }

    foreach ($item in (Get-RecycleBinItems)) {
        $remainingDays = Get-RecycleBinDaysRemaining `
            -PermanentDeletionDate $item.excluir_definitivamente_em

        [void]$grid.Rows.Add(
            $item.id,
            $item.numero,
            $item.excluido_em,
            $item.excluir_definitivamente_em,
            $remainingDays,
            $item.status
        )
    }
    $restoreButton = New-AppButton -Text "RESTAURAR" -X 25 -Y 490 -Width 170 -Height 40 -BackColor $script:ColorSuccess
$restoreButton.Tag = $grid

    $restoreButton.Add_Click({
    try {
        $selectedGrid = $this.Tag

        if ($selectedGrid.SelectedRows.Count -ne 1) {
            Show-AppMessage -Message "Selecione um item da lixeira." -Icon Warning
            return
        }

        $recycleId = [int]$selectedGrid.SelectedRows[0].Cells["ID"].Value

        # Reconfere o estado antes de iniciar qualquer gravação.
        [void](Test-DocumentRestoreReady -RecycleId $recycleId)

        # A função de restauração faz a confirmação e executa a operação.
        $message = Restore-DocumentFromRecycleBin -RecycleId $recycleId

        if ([string]::IsNullOrWhiteSpace([string]$message)) {
            return  # Usuário cancelou na confirmação; nada deve ser anunciado como sucesso.
        }

        Show-AppMessage -Message $message -Icon Information
        Show-RecycleBin
    }
    catch {
        Show-AppMessage -Message (
            "Restauração não confirmada. NÃO clique novamente.`r`n`r`n" +
            $_.Exception.Message +
            "`r`n`r`nConfira os três CSVs e os caminhos dos dois PDFs antes de outra ação."
        ) -Icon Error
    }
})

    $script:ContentPanel.Controls.Add($grid)
    $script:ContentPanel.Controls.Add($restoreButton)

    
}

function Remove-RecycleBinItemIfExpired {
    param([int]$RecycleId)
 
    if ($script:CurrentUser.cargo -ne "ADMINISTRADOR") {
        throw "Apenas administradores podem excluir itens da lixeira."
    }
 
    $recycleItems = @(Get-RecycleBinItems)
    $recycleItem = $recycleItems |
        Where-Object { [int]$_.id -eq $RecycleId } |
        Select-Object -First 1
 
    if ($null -eq $recycleItem) {
        throw "Item não encontrado na lixeira."
    }
 
    $deleteAfter = Convert-ToDateSafe `
        -Value $recycleItem.excluir_definitivamente_em
 
    if ($null -eq $deleteAfter) {
        throw "A data de exclusão definitiva do item é inválida."
    }
 
    if ((Get-Date) -lt $deleteAfter) {
        return "NOT_EXPIRED"
    }
 
    # Remove o PDF da lixeira somente após o prazo.
    if (
        -not [string]::IsNullOrWhiteSpace($recycleItem.arquivo_lixeira) -and
        (Test-Path -LiteralPath $recycleItem.arquivo_lixeira)
    ) {
        Remove-Item `
            -LiteralPath $recycleItem.arquivo_lixeira `
            -Force `
            -ErrorAction Stop
    }
 
    # Remove somente o registro da lixeira selecionado.
    $remainingRecycleItems = @(
        $recycleItems | Where-Object {
            [int]$_.id -ne $RecycleId
        }
    )
 
    Save-CsvRows `
        -Rows $remainingRecycleItems `
        -FilePath $script:RecycleBinFile
 
    # Remove o registro arquivado correspondente, se houver.
    $deletedDocuments = @(Get-DeletedDocuments)
    $remainingDeletedDocuments = @(
        $deletedDocuments | Where-Object {
            [int]$_.id -ne [int]$recycleItem.documento_id
        }
    )
 
    Save-CsvRows `
        -Rows $remainingDeletedDocuments `
        -FilePath $script:DeletedDocumentsFile
 
    return "DELETED"
}

function New-NavSection {
    param([string]$Text)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text.ToUpperInvariant()
    $label.AutoSize = $false
    $label.Height = 28
    $label.Width = 200
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Regular)
    $label.ForeColor = $script:ColorMuted
    $label.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
    $label.Margin = New-Object System.Windows.Forms.Padding(8, 10, 8, 0)
    $label.BackColor = $script:ColorSurface
    [void]$script:NavSections.Add($label)
    [void]$script:NavHost.Controls.Add($label)
}

function New-NavItem {
    param(
        [string]$Text,
        [int]$IconCode,
        [string]$Key,
        [scriptblock]$Action
    )

    $item = New-Object System.Windows.Forms.Panel
    $item.Height = 40
    $item.Width = 210
    $item.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
    $item.BackColor = $script:ColorSurface
    $item.Cursor = [System.Windows.Forms.Cursors]::Hand
    $item.Tag = [pscustomobject]@{
        Key    = $Key
        Action = $Action
        Active = $false
        Text   = $Text
    }

    $icon = New-Object System.Windows.Forms.Label
    $icon.Name = "icon"
    $icon.Text = [char]$IconCode
    $icon.Font = $script:FontIcon
    $icon.ForeColor = $script:ColorMuted
    $icon.BackColor = [System.Drawing.Color]::Transparent
    $icon.Location = New-Object System.Drawing.Point(12, 8)
    $icon.Size = New-Object System.Drawing.Size(22, 22)
    $icon.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

    $caption = New-Object System.Windows.Forms.Label
    $caption.Name = "caption"
    $caption.Text = $Text
    $caption.Font = $script:FontUi
    $caption.ForeColor = $script:ColorMuted
    $caption.BackColor = [System.Drawing.Color]::Transparent
    $caption.Location = New-Object System.Drawing.Point(42, 8)
    $caption.AutoEllipsis = $true
    $caption.Size = New-Object System.Drawing.Size(160, 22)
    $caption.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

    $item.Controls.Add($icon)
    $item.Controls.Add($caption)

    $onClick = {
        $nav = $this
        if (-not ($nav -is [System.Windows.Forms.Panel])) {
            $nav = $nav.Parent
        }
        if ($null -ne $nav.Tag -and $null -ne $nav.Tag.Action) {
            & $nav.Tag.Action
        }
    }

    $onEnter = {
        $nav = $this
        if (-not ($nav -is [System.Windows.Forms.Panel])) {
            $nav = $nav.Parent
        }
        if (-not $nav.Tag.Active) {
            $nav.BackColor = $script:ColorHover
        }
    }

    $onLeave = {
        $nav = $this
        if (-not ($nav -is [System.Windows.Forms.Panel])) {
            $nav = $nav.Parent
        }
        $local = $nav.PointToClient([System.Windows.Forms.Cursor]::Position)
        if (-not $nav.ClientRectangle.Contains($local)) {
            if ($nav.Tag.Active) {
                $nav.BackColor = $script:ColorActive
            }
            else {
                $nav.BackColor = $script:ColorSurface
            }
        }
    }

    foreach ($control in @($item, $icon, $caption)) {
        $control.Add_Click($onClick)
        $control.Add_MouseEnter($onEnter)
        $control.Add_MouseLeave($onLeave)
    }

    [void]$script:NavItems.Add($item)
    [void]$script:NavHost.Controls.Add($item)
}

function Sync-SidebarItems {
    if ($script:SidebarSyncing) {
        return
    }
    if ($null -eq $script:Sidebar -or $null -eq $script:NavHost) {
        return
    }

    $script:SidebarSyncing = $true
    try {
        $collapsed = -not $script:SidebarExpanded
        $script:Sidebar.SuspendLayout()
        $script:NavHost.SuspendLayout()

        $inner = $script:NavHost.ClientSize.Width - $script:NavHost.Padding.Horizontal
        if ($inner -lt 44) {
            $inner = $script:Sidebar.ClientSize.Width - 24
        }
        if ($inner -lt 44) {
            $inner = $(if ($collapsed) { 48 } else { 210 })
        }

        for ($sectionIndex = 0; $sectionIndex -lt $script:NavSections.Count; $sectionIndex++) {
            $section = $script:NavSections[$sectionIndex]
            $section.Visible = -not $collapsed
            $section.Width = [Math]::Max(48, $inner)
        }

        for ($navIndex = 0; $navIndex -lt $script:NavItems.Count; $navIndex++) {
            $navItem = $script:NavItems[$navIndex]
            $navItem.Region = $null
            if ($collapsed) {
                $navItem.Width = 48
            }
            else {
                $navItem.Width = [Math]::Max(48, $inner)
            }

            $caption = $navItem.Controls["caption"]
            $icon = $navItem.Controls["icon"]
            if ($null -ne $caption) {
                $caption.Visible = -not $collapsed
            }
            if ($null -ne $icon) {
                if ($collapsed) {
                    $icon.Left = [Math]::Max(0, [int](($navItem.Width - $icon.Width) / 2))
                }
                else {
                    $icon.Left = 12
                }
            }

            if ($null -ne $script:NavTip) {
                $tip = ""
                if ($collapsed) {
                    $tip = [string]$navItem.Tag.Text
                }
                [void]$script:NavTip.SetToolTip($navItem, $tip)
                if ($null -ne $icon) {
                    [void]$script:NavTip.SetToolTip($icon, $tip)
                }
                if ($null -ne $caption) {
                    [void]$script:NavTip.SetToolTip($caption, $tip)
                }
            }
        }

        if ($null -ne $script:BrandName) {
            $script:BrandName.Visible = -not $collapsed
        }
        if ($null -ne $script:BrandSubtitle) {
            $script:BrandSubtitle.Visible = -not $collapsed
        }

        if ($null -ne $script:BrandPanel -and $null -ne $script:BrandMark -and $null -ne $script:CollapseButton) {
            $script:BrandPanel.Height = 86
            if ($collapsed) {
                $script:BrandMark.Left = [Math]::Max(8, [int](($script:BrandPanel.ClientSize.Width - $script:BrandMark.Width) / 2))
                $script:BrandMark.Top = 10
                $script:CollapseButton.Left = [Math]::Max(8, [int](($script:BrandPanel.ClientSize.Width - $script:CollapseButton.Width) / 2))
                $script:CollapseButton.Top = 50
                $script:CollapseButton.Text = [char]0xE76C
                if ($null -ne $script:NavTip) {
                    [void]$script:NavTip.SetToolTip($script:CollapseButton, "Expandir menu")
                }
            }
            else {
                $script:BrandMark.Left = 16
                $script:BrandMark.Top = 22
                $script:CollapseButton.Left = $script:BrandPanel.ClientSize.Width - 44
                $script:CollapseButton.Top = 26
                $script:CollapseButton.Text = [char]0xE76B
                if ($null -ne $script:NavTip) {
                    [void]$script:NavTip.SetToolTip($script:CollapseButton, "Recolher menu")
                }
            }
        }

        if ($null -ne $script:AccountChip) {
            $script:AccountChip.Region = $null
            $script:AccountChip.BackColor = $script:ColorSurface
            if ($null -ne $script:AccountName) {
                $script:AccountName.Visible = -not $collapsed
            }
            if ($null -ne $script:AccountRole) {
                $script:AccountRole.Visible = -not $collapsed
            }
            if ($null -ne $script:AccountChevron) {
                $script:AccountChevron.Visible = -not $collapsed
            }
            if ($null -ne $script:AccountAvatar) {
                if ($collapsed) {
                    $script:AccountAvatar.Left = [Math]::Max(4, [int](($script:AccountChip.ClientSize.Width - $script:AccountAvatar.Width) / 2))
                }
                else {
                    $script:AccountAvatar.Left = 10
                }
            }
            if (-not $collapsed -and $null -ne $script:AccountChevron) {
                $script:AccountChevron.Left = [Math]::Max(120, $script:AccountChip.ClientSize.Width - 28)
                $script:AccountChevron.Top = 16
            }
        }

        $script:NavHost.ResumeLayout($true)
        $script:Sidebar.ResumeLayout($true)

        for ($navIndex = 0; $navIndex -lt $script:NavItems.Count; $navIndex++) {
            Set-RoundRegion -Control $script:NavItems[$navIndex] -Radius 10
        }
        if ($null -ne $script:BrandMark) {
            Set-RoundRegion -Control $script:BrandMark -Radius 8
        }
        if ($null -ne $script:AccountChip) {
            Set-RoundRegion -Control $script:AccountChip -Radius 12
        }
        if ($null -ne $script:AccountAvatar) {
            Set-RoundRegion -Control $script:AccountAvatar -Radius 8
        }

        Set-ActiveNav -Key $script:CurrentPage
    }
    finally {
        if ($null -ne $script:NavHost) {
            $script:NavHost.ResumeLayout($true)
        }
        if ($null -ne $script:Sidebar) {
            $script:Sidebar.ResumeLayout($true)
        }
        $script:SidebarSyncing = $false
    }
}

function Toggle-Sidebar {
    $script:SidebarExpanded = -not $script:SidebarExpanded
    if ($script:SidebarExpanded) {
        $script:SidebarTargetWidth = $script:SidebarExpandedWidth
    }
    else {
        $script:SidebarTargetWidth = $script:SidebarCollapsedWidth
        if ($null -ne $script:AccountMenu -and -not $script:AccountMenu.IsDisposed) {
            $script:AccountMenu.Close()
        }
    }

    if ($null -ne $script:NavItems) {
        for ($navIndex = 0; $navIndex -lt $script:NavItems.Count; $navIndex++) {
            $script:NavItems[$navIndex].Region = $null
        }
    }
    if ($null -ne $script:AccountChip) {
        $script:AccountChip.Region = $null
        $script:AccountChip.BackColor = $script:ColorSurface
    }

    if ($null -ne $script:SidebarTimer) {
        $script:SidebarTimer.Start()
    }
}

function Update-HeaderActions {
    if ($null -eq $script:HeaderBar -or $null -eq $script:NotificationButton -or $null -eq $script:RefreshButton) {
        return
    }

    $script:NotificationButton.Top = 14
    $script:NotificationButton.Left = $script:HeaderBar.ClientSize.Width - $script:NotificationButton.Width - 20
    $script:RefreshButton.Top = 14
    $script:RefreshButton.Left = $script:NotificationButton.Left - $script:RefreshButton.Width - 8
    if ($null -ne $script:HeaderTitle) {
        $script:HeaderTitle.Width = [Math]::Max(140, $script:RefreshButton.Left - 32)
    }
}

function Register-AccountClick {
    param([System.Windows.Forms.Control]$Control)

    $Control.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Control.Add_Click({ Show-AccountMenu })
    foreach ($child in @($Control.Controls)) {
        Register-AccountClick -Control $child
    }
}

function Register-AccountHover {
    param([System.Windows.Forms.Control]$Control)

    $Control.Add_MouseEnter({
        if ($null -ne $script:AccountChip) {
            $script:AccountChip.BackColor = $script:ColorHover
        }
    })
    $Control.Add_MouseLeave({
        if ($null -eq $script:AccountChip) {
            return
        }
        $local = $script:AccountChip.PointToClient([System.Windows.Forms.Cursor]::Position)
        if (-not $script:AccountChip.ClientRectangle.Contains($local)) {
            $script:AccountChip.BackColor = $script:ColorSurface
        }
    })

    foreach ($child in @($Control.Controls)) {
        Register-AccountHover -Control $child
    }
}

function Show-AccountMenu {
    if (((Get-Date) - $script:AccountMenuClosedAt).TotalMilliseconds -lt 250) {
        return
    }

    if ($null -ne $script:AccountMenu -and -not $script:AccountMenu.IsDisposed) {
        $script:AccountMenu.Close()
        return
    }

    $menu = New-Object System.Windows.Forms.Form
    $menu.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $menu.ShowInTaskbar = $false
    $menu.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $menu.Size = New-Object System.Drawing.Size(230, 118)
    $menu.BackColor = $script:ColorSurface
    $menu.Opacity = 0
    $menu.Font = $script:FontUi

    $menu.Controls.Add((New-AppLabel -Text $script:CurrentUser.nome -X 16 -Y 12 -Width 196 -Height 22 -FontSize 10 -Bold $true -ForeColor $script:ColorText))
    $menu.Controls.Add((New-AppLabel -Text $script:CurrentUser.cargo -X 16 -Y 34 -Width 196 -Height 18 -FontSize 8 -ForeColor $script:ColorMuted))

    $line = New-Object System.Windows.Forms.Panel
    $line.Location = New-Object System.Drawing.Point(12, 60)
    $line.Size = New-Object System.Drawing.Size(206, 1)
    $line.BackColor = $script:ColorBorder
    $menu.Controls.Add($line)

    $signOut = New-Object System.Windows.Forms.Button
    $signOut.Text = "Sair"
    $signOut.Font = $script:FontUi
    $signOut.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $signOut.FlatAppearance.BorderSize = 0
    $signOut.FlatAppearance.MouseOverBackColor = $script:ColorHover
    $signOut.BackColor = $script:ColorSurface
    $signOut.ForeColor = $script:ColorDanger
    $signOut.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $signOut.Padding = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
    $signOut.Location = New-Object System.Drawing.Point(8, 68)
    $signOut.Size = New-Object System.Drawing.Size(214, 40)
    $signOut.Cursor = [System.Windows.Forms.Cursors]::Hand
    $signOut.Add_Click({
        if ($null -ne $script:AccountMenu -and -not $script:AccountMenu.IsDisposed) {
            $script:AccountMenu.Close()
        }
        if ($null -ne $script:MainForm) {
            $script:MainForm.Close()
        }
    })
    $menu.Controls.Add($signOut)

    $menu.Add_Paint({
        $graphics = $_.Graphics
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $bounds = New-Object System.Drawing.Rectangle(0, 0, ($this.Width - 1), ($this.Height - 1))
        if ($bounds.Width -lt 24 -or $bounds.Height -lt 24) {
            return
        }
        $pen = New-Object System.Drawing.Pen($script:ColorBorder)
        $path = Get-RoundPath -Bounds $bounds -Radius 12
        $graphics.DrawPath($pen, $path)
        $path.Dispose()
        $pen.Dispose()
    })

    $origin = $script:AccountChip.PointToScreen((New-Object System.Drawing.Point(0, 0)))
    $menuY = $origin.Y - $menu.Height - 8
    if ($menuY -lt 8) {
        $menuY = 8
    }
    $menu.Location = New-Object System.Drawing.Point($origin.X, $menuY)

    $menu.Add_Shown({ Set-RoundRegion -Control $this -Radius 12 })
    $menu.Add_FormClosed({
        $script:AccountMenuClosedAt = Get-Date
        $script:AccountMenu = $null
    })
    $menu.Add_Deactivate({
        try {
            $this.BeginInvoke([action]{
                try {
                    if ($null -ne $script:AccountMenu -and -not $script:AccountMenu.IsDisposed) {
                        $script:AccountMenu.Close()
                    }
                }
                catch { }
            })
        }
        catch { }
    })

    $script:AccountMenu = $menu
    [void]$menu.Show($script:MainForm)

    $fade = New-Object System.Windows.Forms.Timer
    $fade.Interval = 16
    $fade.Add_Tick({
        $timer = $this
        $target = $script:AccountMenu
        if ($null -eq $target -or $target.IsDisposed) {
            $timer.Stop()
            $timer.Dispose()
            return
        }
        $next = [Math]::Min(1.0, ($target.Opacity + 0.22))
        $target.Opacity = $next
        if ($next -ge 1) {
            $timer.Stop()
            $timer.Dispose()
        }
    })
    $fade.Start()
}

function Show-MainWindow {
    if ($null -ne $script:AccountMenu -and -not $script:AccountMenu.IsDisposed) {
        $script:AccountMenu.Close()
    }

    $script:NavItems = New-Object System.Collections.Generic.List[object]
    $script:NavSections = New-Object System.Collections.Generic.List[object]
    $script:SidebarExpanded = $true
    $script:SidebarTargetWidth = $script:SidebarExpandedWidth
    $script:SidebarSyncing = $false

    if ($null -eq $script:NavTip) {
        $script:NavTip = New-Object System.Windows.Forms.ToolTip
        $script:NavTip.InitialDelay = 350
        $script:NavTip.ShowAlways = $true
    }

    if ($null -eq $script:SidebarTimer) {
        $script:SidebarTimer = New-Object System.Windows.Forms.Timer
        $script:SidebarTimer.Interval = 12
        $script:SidebarTimer.Add_Tick({
            if ($null -eq $script:Sidebar) {
                $this.Stop()
                return
            }

            $current = $script:Sidebar.Width
            $target = [int]$script:SidebarTargetWidth
            $delta = $target - $current
            if ([Math]::Abs($delta) -le 4) {
                $script:Sidebar.Width = $target
                $this.Stop()
                Sync-SidebarItems
                if ($null -ne $script:MainForm) {
                    $script:MainForm.Invalidate($true)
                    $script:MainForm.Update()
                }
                return
            }

            $step = [int]([Math]::Sign($delta) * [Math]::Max(4, [Math]::Round([Math]::Abs($delta) * 0.22)))
            $script:Sidebar.Width = $current + $step
        })
    }

    $script:MainForm = New-Object System.Windows.Forms.Form
    $script:MainForm.Text = $script:AppName
    $script:MainForm.Font = $script:FontUi
    $script:MainForm.Size = New-Object System.Drawing.Size(1280, 820)
    $script:MainForm.MinimumSize = New-Object System.Drawing.Size(1100, 700)
    $script:MainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $script:MainForm.BackColor = $script:ColorBackground
    Enable-DoubleBuffer -Control $script:MainForm
    $script:MainForm.Add_FormClosed({
        if ($null -ne $script:SidebarTimer) {
            $script:SidebarTimer.Stop()
        }
        if ($null -ne $script:AccountMenu -and -not $script:AccountMenu.IsDisposed) {
            $script:AccountMenu.Close()
        }
    })

    $script:Sidebar = New-Object System.Windows.Forms.Panel
    $script:Sidebar.Dock = [System.Windows.Forms.DockStyle]::Left
    $script:Sidebar.Width = $script:SidebarExpandedWidth
    $script:Sidebar.BackColor = $script:ColorSurface
    Enable-DoubleBuffer -Control $script:Sidebar
    $script:Sidebar.Add_Paint({
        $pen = New-Object System.Drawing.Pen($script:ColorBorder)
        $_.Graphics.DrawLine($pen, ($this.Width - 1), 0, ($this.Width - 1), $this.Height)
        $pen.Dispose()
    })

    $script:BrandPanel = New-Object System.Windows.Forms.Panel
    $script:BrandPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $script:BrandPanel.Height = 86
    $script:BrandPanel.BackColor = $script:ColorSurface

    $script:BrandMark = New-Object System.Windows.Forms.Panel
    $script:BrandMark.Size = New-Object System.Drawing.Size(34, 34)
    $script:BrandMark.Location = New-Object System.Drawing.Point(16, 20)
    $script:BrandMark.BackColor = $script:ColorPrimary
    $markText = New-Object System.Windows.Forms.Label
    $markText.Text = "W"
    $markText.Dock = [System.Windows.Forms.DockStyle]::Fill
    $markText.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $markText.ForeColor = $script:ColorWhite
    $markText.Font = $script:FontUiBold
    $markText.BackColor = $script:ColorPrimary
    $script:BrandMark.Controls.Add($markText)

    $script:BrandName = New-AppLabel -Text "Workflow" -X 58 -Y 16 -Width 130 -Height 22 -FontSize 12 -Bold $true -ForeColor $script:ColorText
    $script:BrandSubtitle = New-AppLabel -Text "Documental" -X 58 -Y 38 -Width 140 -Height 22 -FontSize 9 -ForeColor $script:ColorMuted

    $script:CollapseButton = New-Object System.Windows.Forms.Button
    $script:CollapseButton.Size = New-Object System.Drawing.Size(28, 28)
    $script:CollapseButton.Location = New-Object System.Drawing.Point(208, 24)
    $script:CollapseButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $script:CollapseButton.FlatAppearance.BorderSize = 0
    $script:CollapseButton.FlatAppearance.MouseOverBackColor = $script:ColorHover
    $script:CollapseButton.BackColor = $script:ColorSurface
    $script:CollapseButton.ForeColor = $script:ColorMuted
    $script:CollapseButton.Font = $script:FontIcon
    $script:CollapseButton.Text = [char]0xE76B
    $script:CollapseButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $script:CollapseButton.TabStop = $false
    $script:CollapseButton.Add_Click({ Toggle-Sidebar })
    $script:NavTip.SetToolTip($script:CollapseButton, "Recolher menu")

    $script:BrandPanel.Controls.Add($script:CollapseButton)
    $script:BrandPanel.Controls.Add($script:BrandSubtitle)
    $script:BrandPanel.Controls.Add($script:BrandName)
    $script:BrandPanel.Controls.Add($script:BrandMark)

    $script:NavHost = New-Object System.Windows.Forms.FlowLayoutPanel
    $script:NavHost.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:NavHost.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $script:NavHost.WrapContents = $false
    $script:NavHost.AutoScroll = $true
    $script:NavHost.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 8)
    $script:NavHost.BackColor = $script:ColorSurface
    Enable-DoubleBuffer -Control $script:NavHost

    New-NavSection -Text "Fluxo"
    [void](New-NavItem -Text "Dashboard" -IconCode 0xE80F -Key "DASHBOARD" -Action { Show-Dashboard })
    [void](New-NavItem -Text "Documentos" -IconCode 0xE7C3 -Key "DOCUMENTOS" -Action { Show-Documents })

    if ($script:CurrentUser.cargo -eq "ADMINISTRADOR") {
        New-NavSection -Text "Administração"
        [void](New-NavItem -Text "Cadastrar documento" -IconCode 0xE710 -Key "CADASTRO_DOCUMENTO" -Action { Show-RegisterDocument })
        [void](New-NavItem -Text "Usuários" -IconCode 0xE716 -Key "GERENCIAR_USUARIOS" -Action { Show-UserManagement })
        [void](New-NavItem -Text "Exclusão de arquivos" -IconCode 0xE74D -Key "EXCLUSAO_ARQUIVOS" -Action { Show-FileDeletion })
        [void](New-NavItem -Text "Lixeira" -IconCode 0xE777 -Key "LIXEIRA" -Action { Show-RecycleBin })
    }

    $script:AccountPanel = New-Object System.Windows.Forms.Panel
    $script:AccountPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $script:AccountPanel.Height = 78
    $script:AccountPanel.BackColor = $script:ColorSurface
    $script:AccountPanel.Padding = New-Object System.Windows.Forms.Padding(12, 4, 12, 14)

    $script:AccountChip = New-Object System.Windows.Forms.Panel
    $script:AccountChip.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:AccountChip.BackColor = $script:ColorSurface
    $script:AccountChip.Cursor = [System.Windows.Forms.Cursors]::Hand
    Enable-ResizeRedraw -Control $script:AccountChip
    Enable-ResizeRedraw -Control $script:AccountPanel
    $script:AccountChip.Add_Resize({
        $this.Invalidate()
        if ($null -ne $this.Parent) {
            $this.Parent.Invalidate()
        }
    })
    $script:AccountChip.Add_Paint({
        $graphics = $_.Graphics
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $bounds = New-Object System.Drawing.Rectangle(0, 0, ($this.Width - 1), ($this.Height - 1))
        if ($bounds.Width -lt 20 -or $bounds.Height -lt 20) {
            return
        }
        $pen = New-Object System.Drawing.Pen($script:ColorBorder)
        $path = Get-RoundPath -Bounds $bounds -Radius 12
        $graphics.DrawPath($pen, $path)
        $path.Dispose()
        $pen.Dispose()
    })

    $script:AccountAvatar = New-Object System.Windows.Forms.Panel
    $script:AccountAvatar.Size = New-Object System.Drawing.Size(34, 34)
    $script:AccountAvatar.Location = New-Object System.Drawing.Point(10, 10)
    $script:AccountAvatar.BackColor = $script:ColorPrimary
    $avatarIcon = New-Object System.Windows.Forms.Label
    $avatarIcon.Text = [char]0xE77B
    $avatarIcon.Font = $script:FontIcon
    $avatarIcon.ForeColor = $script:ColorWhite
    $avatarIcon.Dock = [System.Windows.Forms.DockStyle]::Fill
    $avatarIcon.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $avatarIcon.BackColor = $script:ColorPrimary
    $script:AccountAvatar.Controls.Add($avatarIcon)

    $script:AccountName = New-AppLabel -Text $script:CurrentUser.nome -X 52 -Y 8 -Width 120 -Height 20 -FontSize 9 -Bold $true -ForeColor $script:ColorText
    $script:AccountRole = New-AppLabel -Text $script:CurrentUser.cargo -X 52 -Y 26 -Width 120 -Height 18 -FontSize 8 -ForeColor $script:ColorMuted
    $script:AccountName.BackColor = [System.Drawing.Color]::Transparent
    $script:AccountRole.BackColor = [System.Drawing.Color]::Transparent

    $script:AccountChevron = New-Object System.Windows.Forms.Label
    $script:AccountChevron.Text = [char]0xE76C
    $script:AccountChevron.Font = $script:FontIcon
    $script:AccountChevron.ForeColor = $script:ColorMuted
    $script:AccountChevron.BackColor = [System.Drawing.Color]::Transparent
    $script:AccountChevron.Size = New-Object System.Drawing.Size(18, 18)
    $script:AccountChevron.Location = New-Object System.Drawing.Point(186, 16)
    $script:AccountChevron.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

    $script:AccountChip.Controls.Add($script:AccountChevron)
    $script:AccountChip.Controls.Add($script:AccountRole)
    $script:AccountChip.Controls.Add($script:AccountName)
    $script:AccountChip.Controls.Add($script:AccountAvatar)
    $script:AccountPanel.Controls.Add($script:AccountChip)
    Register-AccountClick -Control $script:AccountChip
    Register-AccountHover -Control $script:AccountChip

    $script:Sidebar.Controls.Add($script:NavHost)
    $script:Sidebar.Controls.Add($script:AccountPanel)
    $script:Sidebar.Controls.Add($script:BrandPanel)

    $script:HeaderBar = New-Object System.Windows.Forms.Panel
    $script:HeaderBar.Dock = [System.Windows.Forms.DockStyle]::Top
    $script:HeaderBar.Height = 64
    $script:HeaderBar.BackColor = $script:ColorBackground
    Enable-DoubleBuffer -Control $script:HeaderBar

    $script:HeaderTitle = New-AppLabel -Text "Dashboard" -X 24 -Y 14 -Width 360 -Height 36 -FontSize 18 -Bold $true -ForeColor $script:ColorText
    $script:HeaderBar.Controls.Add($script:HeaderTitle)

    $script:RefreshButton = New-AppButton -Text "Atualizar" -X 0 -Y 14 -Width 118 -Height 36 -BackColor $script:ColorSurface
    $script:RefreshButton.Add_Click({ Refresh-CurrentPage })
    $script:HeaderBar.Controls.Add($script:RefreshButton)

    $script:NotificationButton = New-AppButton -Text "Notificações (0)" -X 0 -Y 14 -Width 168 -Height 36 -BackColor $script:ColorSurface
    $script:NotificationButton.Add_Click({ Show-Notifications })
    $script:HeaderBar.Controls.Add($script:NotificationButton)
    $script:HeaderBar.Add_Resize({ Update-HeaderActions })

    $script:ContentPanel = New-Object System.Windows.Forms.Panel
    $script:ContentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:ContentPanel.BackColor = $script:ColorBackground
    $script:ContentPanel.AutoScroll = $true
    Enable-DoubleBuffer -Control $script:ContentPanel

    $script:MainForm.Controls.Add($script:ContentPanel)
    $script:MainForm.Controls.Add($script:HeaderBar)
    $script:MainForm.Controls.Add($script:Sidebar)

    Sync-SidebarItems
    Update-HeaderActions
    Update-NotificationCounter
    Show-Dashboard

    [void]$script:MainForm.ShowDialog()
}

function Get-DeletedDocuments {
    return @(Get-CsvRows -FilePath $script:DeletedDocumentsFile)
}

function Get-RecycleBinItems {
    return @(Get-CsvRows -FilePath $script:RecycleBinFile)
}

function Convert-ToDateSafe {
    param([string]$Value)

    try {
        return [datetime]::Parse($Value)
    }
    catch {
        return $null
    }
}

function Get-RecycleBinDaysRemaining {
    param([string]$PermanentDeletionDate)

    $date = Convert-ToDateSafe -Value $PermanentDeletionDate

    if ($null -eq $date) {
        return -1
    }

    return ($date.Date - (Get-Date).Date).Days
}

# =============================================================================
# INICIALIZAÇÃO
# =============================================================================

try {
    Initialize-DemoUsers

    do {
        $loginResult = Show-Login

        if (
            $loginResult -eq [System.Windows.Forms.DialogResult]::OK -and
            $null -ne $script:CurrentUser
        ) {
            Show-MainWindow
            $script:CurrentUser = $null
        }
        else {
            break
        }
    }
    while ($true)
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Erro crítico:`r`n$($_.Exception.Message)",
        $script:AppName,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}
