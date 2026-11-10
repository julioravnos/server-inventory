<#
.SYNOPSIS
    Interface gráfica para verificar Hyper-V e inventariar VMs em servidores locais e remotos.
.NOTES
    Execute em PowerShell 5.1 (Windows) como administrador:
        powershell.exe -STA -ExecutionPolicy Bypass -File .\HyperV-Inventario-GUI.ps1
    Servidores remotos precisam de WinRM habilitado (Enable-PSRemoting -Force)
    e a conta usada deve ser administradora ou do grupo Hyper-V Administrators.
#>
#Requires -Version 5.1
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Cred = $null
$script:Results = @()

# ----------------------------------------------------------------------
# Coleta executada no servidor de destino
# ----------------------------------------------------------------------
$collect = {
    $ErrorActionPreference = 'SilentlyContinue'
    function G($b) { if ($null -eq $b) { return $null }; [math]::Round([double]$b / 1GB, 2) }
    $os = Get-CimInstance Win32_OperatingSystem; $cs = Get-CimInstance Win32_ComputerSystem
    $cpu = @(Get-CimInstance Win32_Processor)
    $feat = $null
    if (Get-Command Get-WindowsFeature) { $f = Get-WindowsFeature Hyper-V; if ($f) { $feat = [bool]$f.Installed } }
    elseif (Get-Command Get-WindowsOptionalFeature) { $f = Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-All; if ($f) { $feat = ($f.State -eq 'Enabled') } }
    $svc = Get-Service vmms
    $active = [bool]($svc -and $svc.Status -eq 'Running' -and (Get-Command Get-VM))
    $h = [pscustomobject]@{
        Computer = $env:COMPUTERNAME; Reachable = $true; OS = $os.Caption; Build = $os.BuildNumber
        Model = "$($cs.Manufacturer) $($cs.Model)"; CPU = ($cpu | Select-Object -First 1).Name
        Cores = ($cpu | Measure-Object NumberOfCores -Sum).Sum; Logical = ($cpu | Measure-Object NumberOfLogicalProcessors -Sum).Sum
        RAM_GB = G $cs.TotalPhysicalMemory; RAM_FreeGB = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
        FeatureInstalled = $feat; HypervisorPresent = $cs.HypervisorPresent
        VMMS = $(if ($svc) { [string]$svc.Status } else { 'Não encontrado' })
        HyperV_Active = $active; VMCount = 0; Running = 0; Switches = @(); Volumes = @(); VMs = @(); Error = $null
    }
    $h.Volumes = @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | ForEach-Object {
        [pscustomobject]@{ Drive = $_.DeviceID; SizeGB = G $_.Size; FreeGB = G $_.FreeSpace } })
    if (-not $active) { return $h }
    try {
        $ErrorActionPreference = 'Stop'
        $h.Switches = @(Get-VMSwitch | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Type = [string]$_.SwitchType; Adapter = $_.NetAdapterInterfaceDescription } })
        $ErrorActionPreference = 'SilentlyContinue'
        $list = foreach ($vm in @(Get-VM)) {
            $p = Get-VMProcessor -VM $vm; $m = Get-VMMemory -VM $vm
            $disks = @(Get-VMHardDiskDrive -VM $vm | ForEach-Object {
                $vhd = $null; if ($_.Path) { $vhd = Get-VHD -Path $_.Path }
                [pscustomobject]@{ Controller = "$($_.ControllerType) $($_.ControllerNumber):$($_.ControllerLocation)"; Path = $_.Path
                    Type = $(if ($vhd) { [string]$vhd.VhdType }); Format = $(if ($vhd) { [string]$vhd.VhdFormat })
                    MaxGB = $(if ($vhd) { G $vhd.Size }); UsedGB = $(if ($vhd) { G $vhd.FileSize }); Parent = $(if ($vhd) { $vhd.ParentPath }) } })
            $nics = @(Get-VMNetworkAdapter -VM $vm | ForEach-Object {
                $vl = Get-VMNetworkAdapterVlan -VMNetworkAdapter $_
                [pscustomobject]@{ Name = $_.Name; Switch = $_.SwitchName; Mac = $_.MacAddress; IPs = ($_.IPAddresses -join ', ')
                    Vlan = $(if ($vl) { "$($vl.OperationMode) $($vl.AccessVlanId)" }); Status = ($_.Status -join ', ') } })
            $snaps = @(Get-VMSnapshot -VM $vm | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Type = [string]$_.SnapshotType; Created = $_.CreationTime } })
            $integ = @(Get-VMIntegrationService -VM $vm | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Enabled = $_.Enabled; Status = [string]$_.PrimaryStatusDescription } })
            $dvd = @(Get-VMDvdDrive -VM $vm | ForEach-Object { $_.Path } | Where-Object { $_ })
            $rep = Get-VMReplication -VMName $vm.Name
            $sb = $null; $tpm = $null
            if ($vm.Generation -eq 2) { $sb = [string](Get-VMFirmware -VM $vm).SecureBoot; $tpm = (Get-VMSecurity -VM $vm).TpmEnabled }
            [pscustomobject]@{
                Name = $vm.Name; State = [string]$vm.State; Generation = $vm.Generation; Version = $vm.Version
                Uptime = $vm.Uptime.ToString(); Created = $vm.CreationTime; Clustered = $vm.IsClustered; Heartbeat = [string]$vm.Heartbeat
                ConfigPath = $vm.ConfigurationLocation; CheckpointPath = $vm.SnapshotFileLocation
                AutoStart = [string]$vm.AutomaticStartAction; AutoStop = [string]$vm.AutomaticStopAction; Notes = $vm.Notes
                vCPU = $p.Count; CPU_Usage = $vm.CPUUsage; CPU_Reserve = $p.Reserve; CPU_Limit = $p.Maximum; CPU_Weight = $p.RelativeWeight; Nested = $p.ExposeVirtualizationExtensions
                Mem_Dynamic = $m.DynamicMemoryEnabled; Mem_StartupGB = G $m.Startup; Mem_MinGB = G $m.Minimum; Mem_MaxGB = G $m.Maximum
                Mem_AssignedGB = G $vm.MemoryAssigned; Mem_DemandGB = G $vm.MemoryDemand
                Disk_MaxGB = [math]::Round(($disks | Measure-Object MaxGB -Sum).Sum, 2); Disk_UsedGB = [math]::Round(($disks | Measure-Object UsedGB -Sum).Sum, 2)
                SecureBoot = $sb; TPM = $tpm; Disks = $disks; Nics = $nics; Checkpoints = $snaps; Integration = $integ; ISOs = $dvd
                Replication = $(if ($rep) { "$($rep.State) / $($rep.Health) -> $($rep.CurrentReplicaServerName)" })
            } }
        $h.VMs = @($list); $h.VMCount = @($list).Count; $h.Running = @($list | Where-Object State -eq 'Running').Count
    } catch { $h.Error = $_.Exception.Message }
    return $h
}

# ----------------------------------------------------------------------
# Funções auxiliares
# ----------------------------------------------------------------------
function Test-Port($name, $port, $ms = 4000) {
    $t = New-Object System.Net.Sockets.TcpClient
    try { $a = $t.BeginConnect($name, $port, $null, $null); return ($a.AsyncWaitHandle.WaitOne($ms, $false) -and $t.Connected) }
    catch { return $false } finally { $t.Close() }
}

function Get-HostData($name) {
    $local = $name -in @('.', 'localhost', '127.0.0.1', $env:COMPUTERNAME)
    $fail = { param($msg) [pscustomobject]@{ Computer = $name; Reachable = $false; HyperV_Active = $false; VMs = @(); VMCount = 0; Running = 0; Error = $msg } }
    try {
        if ($local) { $job = Start-Job -ScriptBlock $collect }
        else {
            try { $null = [System.Net.Dns]::GetHostAddresses($name) } catch { return (& $fail "Falha de DNS ao resolver '$name'.") }
            $port = if ($chkSSL.Checked) { 5986 } else { 5985 }
            if (-not (Test-Port $name $port)) { return (& $fail "Porta $port (WinRM) inacessível. Verifique firewall/WinRM/servidor ligado.") }
            $p = @{ ComputerName = $name; ScriptBlock = $collect; AsJob = $true; ErrorAction = 'Stop' }
            if ($script:Cred) { $p.Credential = $script:Cred }
            if ($chkSSL.Checked) { $p.UseSSL = $true }
            if ($cmbAuth.Text -ne 'Default') { $p.Authentication = $cmbAuth.Text }
            $job = Invoke-Command @p
        }
        while ($job.State -in 'Running', 'NotStarted') { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100 }
        $r = Receive-Job $job -ErrorAction Stop; Remove-Job $job -Force
        return ($r | Select-Object * -ExcludeProperty PSComputerName, RunspaceId, PSShowComputerName)
    } catch { return (& $fail $_.Exception.Message) }
}

function Format-VM($v) {
    $s = New-Object System.Text.StringBuilder
    $n = { param($t) [void]$s.AppendLine($t) }
    & $n "VM: $($v.Name)   [$($v.State)]"
    & $n ('=' * 70)
    & $n "Geração: $($v.Generation) | Versão: $($v.Version) | Cluster: $($v.Clustered) | Heartbeat: $($v.Heartbeat)"
    & $n "Criada: $($v.Created) | Uptime: $($v.Uptime)"
    & $n "Config: $($v.ConfigPath)"
    & $n "Checkpoints em: $($v.CheckpointPath)"
    & $n "Início auto: $($v.AutoStart) | Parada auto: $($v.AutoStop)"
    if ($v.Generation -eq 2) { & $n "Secure Boot: $($v.SecureBoot) | TPM: $($v.TPM)" }
    if ($v.Notes) { & $n "Notas: $($v.Notes)" }
    & $n ''; & $n '--- PROCESSADOR ---'
    & $n "vCPU: $($v.vCPU) | Uso: $($v.CPU_Usage)% | Reserva: $($v.CPU_Reserve)% | Limite: $($v.CPU_Limit)% | Peso: $($v.CPU_Weight) | Nested: $($v.Nested)"
    & $n ''; & $n '--- MEMÓRIA ---'
    & $n "Dinâmica: $($v.Mem_Dynamic) | Inicial: $($v.Mem_StartupGB) GB | Mín: $($v.Mem_MinGB) GB | Máx: $($v.Mem_MaxGB) GB"
    & $n "Atribuída: $($v.Mem_AssignedGB) GB | Demanda: $($v.Mem_DemandGB) GB"
    & $n ''; & $n '--- DISCOS ---'
    foreach ($d in $v.Disks) { & $n "[$($d.Controller)] $($d.Type) $($d.Format) | máx $($d.MaxGB) GB | usado $($d.UsedGB) GB"; & $n "    $($d.Path)"; if ($d.Parent) { & $n "    Pai: $($d.Parent)" } }
    foreach ($i in $v.ISOs) { & $n "[DVD] $i" }
    & $n ''; & $n '--- REDE ---'
    foreach ($x in $v.Nics) { & $n "$($x.Name) | Switch: $($x.Switch) | MAC: $($x.Mac) | IPs: $($x.IPs) | VLAN: $($x.Vlan) | $($x.Status)" }
    & $n ''; & $n '--- CHECKPOINTS ---'
    if ($v.Checkpoints) { foreach ($c in $v.Checkpoints) { & $n "$($c.Name) | $($c.Type) | $($c.Created)" } } else { & $n '(nenhum)' }
    & $n ''; & $n '--- SERVIÇOS DE INTEGRAÇÃO ---'
    foreach ($i in $v.Integration) { & $n "$($i.Name): $(if ($i.Enabled) { 'Ativo' } else { 'Desativado' }) ($($i.Status))" }
    if ($v.Replication) { & $n ''; & $n "--- REPLICAÇÃO ---"; & $n $v.Replication }
    $s.ToString()
}

function Update-Grids {
    $hosts = New-Object System.Collections.ArrayList
    $vms = New-Object System.Collections.ArrayList
    foreach ($r in $script:Results) {
        [void]$hosts.Add([pscustomobject]@{
            Servidor = $r.Computer; Acessível = $r.Reachable; 'Hyper-V ativo' = $r.HyperV_Active
            Recurso = $r.FeatureInstalled; VMMS = $r.VMMS; SO = $r.OS; CPU = $r.CPU; Núcleos = $r.Cores
            'RAM (GB)' = $r.RAM_GB; VMs = $r.VMCount; 'Em execução' = $r.Running; Erro = $r.Error })
        foreach ($v in $r.VMs) {
            [void]$vms.Add([pscustomobject]@{
                Servidor = $r.Computer; VM = $v.Name; Estado = $v.State; Ger = $v.Generation; vCPU = $v.vCPU
                'RAM atrib. (GB)' = $v.Mem_AssignedGB; 'RAM dinâmica' = $v.Mem_Dynamic
                'Disco máx (GB)' = $v.Disk_MaxGB; 'Disco usado (GB)' = $v.Disk_UsedGB
                IPs = (($v.Nics | ForEach-Object { $_.IPs }) -join ' | '); Checkpoints = @($v.Checkpoints).Count; Uptime = $v.Uptime })
        }
    }
    $gridHosts.DataSource = $hosts; $gridVMs.DataSource = $vms
    foreach ($g in $gridHosts, $gridVMs) { $g.AutoResizeColumns('AllCells') }
}

# ----------------------------------------------------------------------
# Interface
# ----------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Inventário Hyper-V'; $form.Size = '1200,780'; $form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$top = New-Object System.Windows.Forms.Panel; $top.Dock = 'Top'; $top.Height = 170
$lbl = New-Object System.Windows.Forms.Label; $lbl.Text = 'Servidores (um por linha; "localhost" para a máquina local):'; $lbl.Location = '10,8'; $lbl.AutoSize = $true
$txt = New-Object System.Windows.Forms.TextBox; $txt.Multiline = $true; $txt.ScrollBars = 'Vertical'; $txt.Location = '10,30'; $txt.Size = '420,130'
$txt.Text = $env:COMPUTERNAME

function New-Btn($text, $x, $y, $w = 150) {
    $b = New-Object System.Windows.Forms.Button; $b.Text = $text; $b.Location = "$x,$y"; $b.Size = "$w,30"; $b
}
$btnFile = New-Btn 'Carregar lista (.txt)' 445 30
$btnCred = New-Btn 'Definir credenciais...' 445 65
$btnRun = New-Btn 'COLETAR' 445 100; $btnRun.BackColor = [System.Drawing.Color]::FromArgb(11, 83, 148); $btnRun.ForeColor = 'White'; $btnRun.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$btnExp = New-Btn 'Exportar relatório...' 445 135
$lblCred = New-Object System.Windows.Forms.Label; $lblCred.Text = 'Credenciais: usuário atual'; $lblCred.Location = '610,38'; $lblCred.AutoSize = $true
$chkSSL = New-Object System.Windows.Forms.CheckBox; $chkSSL.Text = 'Usar WinRM sobre HTTPS (5986)'; $chkSSL.Location = '610,72'; $chkSSL.AutoSize = $true
$lblAuth = New-Object System.Windows.Forms.Label; $lblAuth.Text = 'Autenticação:'; $lblAuth.Location = '610,108'; $lblAuth.AutoSize = $true
$cmbAuth = New-Object System.Windows.Forms.ComboBox; $cmbAuth.DropDownStyle = 'DropDownList'; $cmbAuth.Location = '700,104'; $cmbAuth.Width = 120
[void]$cmbAuth.Items.AddRange(@('Default', 'Negotiate', 'Kerberos', 'Credssp', 'Basic')); $cmbAuth.SelectedIndex = 0
$top.Controls.AddRange(@($lbl, $txt, $btnFile, $btnCred, $btnRun, $btnExp, $lblCred, $chkSSL, $lblAuth, $cmbAuth))

$bottom = New-Object System.Windows.Forms.Panel; $bottom.Dock = 'Bottom'; $bottom.Height = 28
$status = New-Object System.Windows.Forms.Label; $status.Text = 'Pronto.'; $status.Dock = 'Fill'; $status.TextAlign = 'MiddleLeft'
$bar = New-Object System.Windows.Forms.ProgressBar; $bar.Dock = 'Right'; $bar.Width = 250
$bottom.Controls.AddRange(@($status, $bar))

function New-Grid { $g = New-Object System.Windows.Forms.DataGridView; $g.Dock = 'Fill'; $g.ReadOnly = $true; $g.AllowUserToAddRows = $false
    $g.SelectionMode = 'FullRowSelect'; $g.AutoSizeColumnsMode = 'None'; $g.BackgroundColor = 'White'; $g.RowHeadersVisible = $false; $g }
$gridHosts = New-Grid; $gridVMs = New-Grid
$detail = New-Object System.Windows.Forms.TextBox; $detail.Multiline = $true; $detail.ReadOnly = $true; $detail.ScrollBars = 'Both'
$detail.WordWrap = $false; $detail.Dock = 'Fill'; $detail.Font = New-Object System.Drawing.Font('Consolas', 9.5)

$split = New-Object System.Windows.Forms.SplitContainer; $split.Dock = 'Fill'; $split.Orientation = 'Horizontal'; $split.SplitterDistance = 220
$split.Panel1.Controls.Add($gridVMs); $split.Panel2.Controls.Add($detail)

$tabs = New-Object System.Windows.Forms.TabControl; $tabs.Dock = 'Fill'
foreach ($t in @(@('Servidores', $gridHosts), @('Máquinas virtuais (clique em uma VM para ver detalhes)', $split))) {
    $tp = New-Object System.Windows.Forms.TabPage; $tp.Text = $t[0]; $tp.Controls.Add($t[1]); $tabs.TabPages.Add($tp)
}
$form.Controls.AddRange(@($tabs, $top, $bottom))

# ----------------------------------------------------------------------
# Eventos
# ----------------------------------------------------------------------
$btnFile.Add_Click({
    $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = 'Texto (*.txt)|*.txt|Todos|*.*'
    if ($d.ShowDialog() -eq 'OK') { $txt.Text = (Get-Content $d.FileName) -join "`r`n" }
})

$btnCred.Add_Click({
    $c = Get-Credential -Message 'Credenciais para os servidores remotos'
    if ($c) { $script:Cred = $c; $lblCred.Text = "Credenciais: $($c.UserName)" }
})

$gridVMs.Add_SelectionChanged({
    if ($gridVMs.CurrentRow) {
        $srv = $gridVMs.CurrentRow.Cells['Servidor'].Value; $vmn = $gridVMs.CurrentRow.Cells['VM'].Value
        $r = $script:Results | Where-Object Computer -eq $srv | Select-Object -First 1
        $v = $r.VMs | Where-Object Name -eq $vmn | Select-Object -First 1
        if ($v) { $detail.Text = Format-VM $v }
    }
})

$btnRun.Add_Click({
    $targets = @($txt.Lines | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') } | Select-Object -Unique)
    if ($targets.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show('Informe ao menos um servidor.'); return }
    $btnRun.Enabled = $false; $script:Results = @(); $bar.Maximum = $targets.Count; $bar.Value = 0; $i = 0
    foreach ($t in $targets) {
        $status.Text = "Consultando $t ($($i + 1) de $($targets.Count))..."
        $script:Results += Get-HostData $t
        $i++; $bar.Value = $i; Update-Grids; [System.Windows.Forms.Application]::DoEvents()
    }
    $ok = @($script:Results | Where-Object HyperV_Active).Count
    $status.Text = "Concluído: $($targets.Count) servidor(es) consultado(s), $ok com Hyper-V ativo."
    $btnRun.Enabled = $true
})

$btnExp.Add_Click({
    if (-not $script:Results) { [void][System.Windows.Forms.MessageBox]::Show('Execute a coleta primeiro.'); return }
    $d = New-Object System.Windows.Forms.FolderBrowserDialog; $d.Description = 'Escolha a pasta para salvar o relatório'
    if ($d.ShowDialog() -ne 'OK') { return }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'; $base = Join-Path $d.SelectedPath "HyperV_$stamp"
    $script:Results | ConvertTo-Json -Depth 8 | Out-File "$base.json" -Encoding UTF8
    $gridHosts.DataSource | Export-Csv "$base`_servidores.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    $gridVMs.DataSource | Export-Csv "$base`_vms.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    $css = '<style>body{font-family:Segoe UI,Arial;margin:20px}h2{background:#0b5394;color:#fff;padding:6px}table{border-collapse:collapse;font-size:12px}th{background:#d9e2f3}th,td{border:1px solid #bbb;padding:3px 8px}pre{background:#f4f4f4;padding:8px}</style>'
    $html = "<html><head><meta charset='utf-8'>$css</head><body><h1>Inventário Hyper-V - $(Get-Date -Format 'dd/MM/yyyy HH:mm')</h1>"
    $html += ($gridHosts.DataSource | ConvertTo-Html -Fragment | Out-String)
    foreach ($r in $script:Results) {
        $html += "<h2>$($r.Computer)</h2>"
        if ($r.Error) { $html += "<p style='color:#c00'>$($r.Error)</p>" }
        if ($r.Volumes) { $html += '<h3>Volumes do host</h3>' + ($r.Volumes | ConvertTo-Html -Fragment | Out-String) }
        if ($r.Switches) { $html += '<h3>Switches virtuais</h3>' + ($r.Switches | ConvertTo-Html -Fragment | Out-String) }
        foreach ($v in $r.VMs) { $html += "<pre>$([System.Net.WebUtility]::HtmlEncode((Format-VM $v)))</pre>" }
    }
    $html + '</body></html>' | Out-File "$base.html" -Encoding UTF8
    $status.Text = "Relatórios salvos em $($d.SelectedPath)"
    Start-Process explorer.exe $d.SelectedPath
})

[void]$form.ShowDialog()
