<#
.SYNOPSIS
    Verifica se o Hyper-V está ativo em um ou mais servidores e levanta o inventário completo das VMs.

.DESCRIPTION
    Para cada servidor informado, o script:
      1. Verifica conectividade (WinRM) e coleta dados do sistema operacional e hardware do host.
      2. Verifica se o recurso/função Hyper-V está instalado, se o serviço VMMS está em execução
         e se o hypervisor está carregado.
      3. Se o Hyper-V estiver ativo, lista TODAS as VMs (ligadas, desligadas, pausadas, salvas...)
         com: geral, processador, memória, discos (VHD/VHDX), rede, VLAN, checkpoints,
         serviços de integração, replicação, firmware/segurança (Gen2) e cluster.
      4. Exporta os resultados em JSON, CSV e HTML.

.PARAMETER ComputerName
    Um ou mais servidores (nome ou IP). Padrão: máquina local.

.PARAMETER Credential
    Credencial para conexão remota (opcional).

.PARAMETER OutputPath
    Pasta onde os relatórios serão gravados. Padrão: .\HyperV-Report

.PARAMETER NoExport
    Apenas exibe no console, sem gerar arquivos.

.EXAMPLE
    .\Get-HyperVInventory.ps1

.EXAMPLE
    .\Get-HyperVInventory.ps1 -ComputerName SRV-HV01,SRV-HV02 -Credential (Get-Credential)

.EXAMPLE
    "SRV-HV01","10.0.0.20" | .\Get-HyperVInventory.ps1 -OutputPath C:\Relatorios

.NOTES
    Requisitos:
      - Executar como administrador (ou com conta com permissão de Hyper-V Administrators nos hosts).
      - WinRM habilitado nos servidores remotos (Enable-PSRemoting -Force).
      - O módulo Hyper-V só é necessário no host de destino (onde as VMs rodam).
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(ValueFromPipeline = $true)]
    [string[]]$ComputerName = $env:COMPUTERNAME,

    [pscredential]$Credential,

    # Arquivo .txt com um servidor por linha (linhas vazias e iniciadas com # são ignoradas)
    [string]$ComputerListPath,

    # Usar WinRM sobre HTTPS (porta 5986)
    [switch]$UseSSL,

    # Tipo de autenticação (use Negotiate para IP/workgroup, Kerberos para domínio)
    [ValidateSet('Default', 'Negotiate', 'Kerberos', 'Credssp', 'Basic')]
    [string]$Authentication = 'Default',

    [int]$TimeoutSeconds = 5,

    [string]$OutputPath = (Join-Path (Get-Location) 'HyperV-Report'),

    [switch]$NoExport
)

begin {
    $targets = New-Object System.Collections.Generic.List[string]

    # ------------------------------------------------------------------
    # Bloco executado NO SERVIDOR DE DESTINO (local ou remoto)
    # ------------------------------------------------------------------
    $collect = {
        $ErrorActionPreference = 'SilentlyContinue'

        function ToGB($bytes) {
            if ($null -eq $bytes) { return $null }
            [math]::Round(([double]$bytes) / 1GB, 2)
        }

        $os  = Get-CimInstance Win32_OperatingSystem
        $cs  = Get-CimInstance Win32_ComputerSystem
        $cpu = @(Get-CimInstance Win32_Processor)
        $bios = Get-CimInstance Win32_BIOS

        # --- Hyper-V instalado? ---
        $featureInstalled = $null
        if (Get-Command Get-WindowsFeature) {
            $f = Get-WindowsFeature -Name 'Hyper-V'
            if ($f) { $featureInstalled = [bool]$f.Installed }
        }
        elseif (Get-Command Get-WindowsOptionalFeature) {
            $f = Get-WindowsOptionalFeature -Online -FeatureName 'Microsoft-Hyper-V-All'
            if ($f) { $featureInstalled = ($f.State -eq 'Enabled') }
        }

        $vmms     = Get-Service -Name vmms
        $vmcomp   = Get-Service -Name vmcompute
        $moduleOk = [bool](Get-Command Get-VM)
        $active   = [bool]($vmms -and $vmms.Status -eq 'Running' -and $moduleOk)

        $hostInfo = [pscustomobject]@{
            Computer               = $env:COMPUTERNAME
            Domain                 = $cs.Domain
            Reachable              = $true
            OS                     = $os.Caption
            OSVersion              = $os.Version
            OSBuild                = $os.BuildNumber
            LastBoot               = $os.LastBootUpTime
            Manufacturer           = $cs.Manufacturer
            Model                  = $cs.Model
            SerialNumber           = $bios.SerialNumber
            CPUModel               = ($cpu | Select-Object -First 1).Name
            CPUSockets             = $cpu.Count
            CPUCoresTotal          = ($cpu | Measure-Object NumberOfCores -Sum).Sum
            CPULogicalTotal        = ($cpu | Measure-Object NumberOfLogicalProcessors -Sum).Sum
            RAM_TotalGB            = ToGB $cs.TotalPhysicalMemory
            RAM_FreeGB             = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
            HyperV_FeatureInstalled = $featureInstalled
            HyperV_HypervisorPresent = $cs.HypervisorPresent
            HyperV_VMMS_Status     = if ($vmms) { [string]$vmms.Status } else { 'Não encontrado' }
            HyperV_VMMS_StartType  = if ($vmms) { [string]$vmms.StartType } else { $null }
            HyperV_VMCompute_Status = if ($vmcomp) { [string]$vmcomp.Status } else { 'Não encontrado' }
            HyperV_ModuleAvailable = $moduleOk
            HyperV_Active          = $active
            VMHost                 = $null
            Switches               = @()
            Volumes                = @()
            VMCount                = 0
            VMsRunning             = 0
            VMs                    = @()
            Error                  = $null
        }

        # Volumes do host (útil para análise de capacidade)
        $hostInfo.Volumes = @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | ForEach-Object {
            [pscustomobject]@{
                Drive     = $_.DeviceID
                Label     = $_.VolumeName
                FileSystem = $_.FileSystem
                SizeGB    = ToGB $_.Size
                FreeGB    = ToGB $_.FreeSpace
                FreePct   = if ($_.Size) { [math]::Round(($_.FreeSpace / $_.Size) * 100, 1) } else { $null }
            }
        })

        if (-not $active) { return $hostInfo }

        try {
            $ErrorActionPreference = 'Stop'

            # --- Configuração do host Hyper-V ---
            $vh = Get-VMHost
            $hostInfo.VMHost = [pscustomobject]@{
                VirtualMachinePath       = $vh.VirtualMachinePath
                VirtualHardDiskPath      = $vh.VirtualHardDiskPath
                EnableEnhancedSessionMode = $vh.EnableEnhancedSessionMode
                NumaSpanningEnabled      = $vh.NumaSpanningEnabled
                LiveMigrationEnabled     = $vh.VirtualMachineMigrationEnabled
                MaxConcurrentMigrations  = $vh.MaximumVirtualMachineMigrations
                MacAddressMinimum        = $vh.MacAddressMinimum
                MacAddressMaximum        = $vh.MacAddressMaximum
            }

            # --- Switches virtuais ---
            $hostInfo.Switches = @(Get-VMSwitch | ForEach-Object {
                [pscustomobject]@{
                    Name              = $_.Name
                    Type              = [string]$_.SwitchType
                    NetAdapter        = $_.NetAdapterInterfaceDescription
                    AllowManagementOS = $_.AllowManagementOS
                    IovEnabled        = $_.IovEnabled
                    BandwidthMode     = [string]$_.BandwidthReservationMode
                }
            })

            # --- Máquinas virtuais ---
            $ErrorActionPreference = 'SilentlyContinue'
            $vms = @(Get-VM)
            $vmList = foreach ($vm in $vms) {

                $proc = Get-VMProcessor -VM $vm
                $mem  = Get-VMMemory -VM $vm

                # Discos
                $disks = @(Get-VMHardDiskDrive -VM $vm | ForEach-Object {
                    $d = $_
                    $vhd = $null
                    if ($d.Path) { $vhd = Get-VHD -Path $d.Path }
                    [pscustomobject]@{
                        Controller     = "$($d.ControllerType) $($d.ControllerNumber):$($d.ControllerLocation)"
                        Path           = $d.Path
                        PassThroughDisk = if ($d.DiskNumber -ne $null) { $d.DiskNumber } else { $null }
                        Format         = if ($vhd) { [string]$vhd.VhdFormat } else { $null }
                        Type           = if ($vhd) { [string]$vhd.VhdType } else { $null }
                        SizeMaxGB      = if ($vhd) { ToGB $vhd.Size } else { $null }
                        SizeUsedGB     = if ($vhd) { ToGB $vhd.FileSize } else { $null }
                        MinimumSizeGB  = if ($vhd) { ToGB $vhd.MinimumSize } else { $null }
                        ParentPath     = if ($vhd) { $vhd.ParentPath } else { $null }
                        BlockSizeMB    = if ($vhd) { [math]::Round($vhd.BlockSize / 1MB, 2) } else { $null }
                        LogicalSectorSize  = if ($vhd) { $vhd.LogicalSectorSize } else { $null }
                        PhysicalSectorSize = if ($vhd) { $vhd.PhysicalSectorSize } else { $null }
                        QoS_MinIOPS    = $d.MinimumIOPS
                        QoS_MaxIOPS    = $d.MaximumIOPS
                    }
                })

                # DVD
                $dvds = @(Get-VMDvdDrive -VM $vm | ForEach-Object {
                    [pscustomobject]@{
                        Controller = "$($_.ControllerType) $($_.ControllerNumber):$($_.ControllerLocation)"
                        ISO        = $_.Path
                    }
                })

                # Rede
                $nics = @(Get-VMNetworkAdapter -VM $vm | ForEach-Object {
                    $n = $_
                    $vlan = Get-VMNetworkAdapterVlan -VMNetworkAdapter $n
                    [pscustomobject]@{
                        Name           = $n.Name
                        Switch         = $n.SwitchName
                        Connected      = $n.Connected
                        Status         = ($n.Status -join ', ')
                        MacAddress     = $n.MacAddress
                        DynamicMac     = $n.DynamicMacAddressEnabled
                        IPAddresses    = ($n.IPAddresses -join ', ')
                        VlanMode       = if ($vlan) { [string]$vlan.OperationMode } else { $null }
                        VlanId         = if ($vlan) { $vlan.AccessVlanId } else { $null }
                        MacSpoofing    = [string]$n.MacAddressSpoofing
                        DhcpGuard      = [string]$n.DhcpGuard
                        RouterGuard    = [string]$n.RouterGuard
                        BandwidthMaxMbps = if ($n.BandwidthSetting) { [math]::Round($n.BandwidthSetting.MaximumBandwidth / 1MB, 2) } else { $null }
                        VmqWeight      = $n.VmqWeight
                        IovWeight      = $n.IovWeight
                    }
                })

                # Checkpoints
                $snaps = @(Get-VMSnapshot -VM $vm | ForEach-Object {
                    [pscustomobject]@{
                        Name         = $_.Name
                        Type         = [string]$_.SnapshotType
                        Created      = $_.CreationTime
                        ParentName   = $_.ParentSnapshotName
                    }
                })

                # Serviços de integração
                $integ = @(Get-VMIntegrationService -VM $vm | ForEach-Object {
                    [pscustomobject]@{
                        Name    = $_.Name
                        Enabled = $_.Enabled
                        Status  = [string]$_.PrimaryStatusDescription
                    }
                })

                # Replicação
                $repl = Get-VMReplication -VMName $vm.Name
                $replication = if ($repl) {
                    [pscustomobject]@{
                        State           = [string]$repl.State
                        Health          = [string]$repl.Health
                        Mode            = [string]$repl.Mode
                        Relationship    = [string]$repl.ReplicationRelationshipType
                        ReplicaServer   = $repl.CurrentReplicaServerName
                        LastReplication = $repl.LastReplicationTime
                    }
                } else { $null }

                # Firmware / segurança (Gen2)
                $secureBoot = $null; $tpm = $null; $bootOrder = $null
                if ($vm.Generation -eq 2) {
                    $fw = Get-VMFirmware -VM $vm
                    if ($fw) {
                        $secureBoot = [string]$fw.SecureBoot
                        $bootOrder  = ($fw.BootOrder | ForEach-Object { $_.BootType }) -join ' > '
                    }
                    $sec = Get-VMSecurity -VM $vm
                    if ($sec) { $tpm = $sec.TpmEnabled }
                }

                [pscustomobject]@{
                    # Geral
                    Name              = $vm.Name
                    Id                = [string]$vm.Id
                    State             = [string]$vm.State
                    Status            = $vm.Status
                    Generation        = $vm.Generation
                    ConfigVersion     = $vm.Version
                    Uptime            = $vm.Uptime.ToString()
                    Created           = $vm.CreationTime
                    IsClustered       = $vm.IsClustered
                    Heartbeat         = [string]$vm.Heartbeat
                    Notes             = $vm.Notes
                    # Caminhos
                    ConfigPath        = $vm.ConfigurationLocation
                    CheckpointPath    = $vm.SnapshotFileLocation
                    SmartPagingPath   = $vm.SmartPagingFilePath
                    # Inicialização
                    AutoStartAction   = [string]$vm.AutomaticStartAction
                    AutoStartDelaySec = $vm.AutomaticStartDelay
                    AutoStopAction    = [string]$vm.AutomaticStopAction
                    CheckpointType    = [string]$vm.CheckpointType
                    # Processador
                    vCPU              = $proc.Count
                    CPU_UsagePct      = $vm.CPUUsage
                    CPU_ReservePct    = $proc.Reserve
                    CPU_LimitPct      = $proc.Maximum
                    CPU_Weight        = $proc.RelativeWeight
                    CPU_NestedVirt    = $proc.ExposeVirtualizationExtensions
                    CPU_MigrationCompat = $proc.CompatibilityForMigrationEnabled
                    # Memória
                    Mem_Dynamic       = $mem.DynamicMemoryEnabled
                    Mem_StartupGB     = ToGB $mem.Startup
                    Mem_MinGB         = ToGB $mem.Minimum
                    Mem_MaxGB         = ToGB $mem.Maximum
                    Mem_AssignedGB    = ToGB $vm.MemoryAssigned
                    Mem_DemandGB      = ToGB $vm.MemoryDemand
                    Mem_BufferPct     = $mem.Buffer
                    Mem_Priority      = $mem.Priority
                    # Discos (resumo)
                    Disk_Count        = $disks.Count
                    Disk_TotalMaxGB   = [math]::Round((($disks | Measure-Object SizeMaxGB -Sum).Sum), 2)
                    Disk_TotalUsedGB  = [math]::Round((($disks | Measure-Object SizeUsedGB -Sum).Sum), 2)
                    # Firmware / segurança
                    SecureBoot        = $secureBoot
                    TPM               = $tpm
                    BootOrder         = $bootOrder
                    # Detalhes
                    Disks             = $disks
                    DvdDrives         = $dvds
                    NetworkAdapters   = $nics
                    Checkpoints       = $snaps
                    IntegrationServices = $integ
                    Replication       = $replication
                }
            }

            $hostInfo.VMs        = @($vmList)
            $hostInfo.VMCount    = @($vmList).Count
            $hostInfo.VMsRunning = @($vmList | Where-Object { $_.State -eq 'Running' }).Count
        }
        catch {
            $hostInfo.Error = $_.Exception.Message
        }

        return $hostInfo
    }

    function Get-HostData {
        param([string]$Name)

        $localNames = @('.', 'localhost', '127.0.0.1', $env:COMPUTERNAME)
        if ($Name -in $localNames) {
            return (& $collect)
        }

        # --- Diagnóstico de conectividade antes de tentar a coleta ---
        $port = if ($UseSSL) { 5986 } else { 5985 }

        try { $null = [System.Net.Dns]::GetHostAddresses($Name) }
        catch {
            return [pscustomobject]@{
                Computer = $Name; Reachable = $false; HyperV_Active = $false; VMs = @()
                Error = "Falha de DNS: não foi possível resolver '$Name'."
            }
        }

        $tcp = New-Object System.Net.Sockets.TcpClient
        try {
            $iar = $tcp.BeginConnect($Name, $port, $null, $null)
            $okPort = $iar.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($TimeoutSeconds), $false) -and $tcp.Connected
        } catch { $okPort = $false } finally { $tcp.Close() }

        if (-not $okPort) {
            return [pscustomobject]@{
                Computer = $Name; Reachable = $false; HyperV_Active = $false; VMs = @()
                Error = "Porta $port (WinRM) inacessível. Verifique firewall, se o WinRM está habilitado (Enable-PSRemoting -Force) e se o servidor está ligado."
            }
        }

        $params = @{
            ComputerName = $Name
            ScriptBlock  = $collect
            ErrorAction  = 'Stop'
        }
        if ($Credential)                    { $params.Credential = $Credential }
        if ($UseSSL)                        { $params.UseSSL = $true }
        if ($Authentication -ne 'Default')  { $params.Authentication = $Authentication }

        try {
            Invoke-Command @params |
                Select-Object * -ExcludeProperty PSComputerName, RunspaceId, PSShowComputerName
        }
        catch {
            [pscustomobject]@{
                Computer  = $Name
                Reachable = $false
                HyperV_Active = $false
                VMs       = @()
                Error     = "Falha ao conectar (WinRM/permissões): $($_.Exception.Message)"
            }
        }
    }

    function ConvertTo-Fragment($data, $title) {
        $rows = @($data)
        if ($rows.Count -eq 0) { return "<h3>$title</h3><p><i>Nenhum item.</i></p>" }
        "<h3>$title</h3>" + ($rows | ConvertTo-Html -Fragment | Out-String)
    }
}

process {
    foreach ($c in $ComputerName) { if ($c) { $targets.Add($c) } }
}

end {
    if ($ComputerListPath) {
        if (-not (Test-Path $ComputerListPath)) { throw "Arquivo de lista não encontrado: $ComputerListPath" }
        Get-Content $ComputerListPath |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith('#') } |
            ForEach-Object { $targets.Add($_) }
    }
    # Se foi informado lista/servidores explícitos, não inclui a máquina local por padrão
    if ($ComputerListPath -and -not $PSBoundParameters.ContainsKey('ComputerName')) {
        $targets.Remove($env:COMPUTERNAME) | Out-Null
        if ($targets.Count -eq 0) { $targets.Add($env:COMPUTERNAME) }
    }

    $results = @()
    $i = 0
    foreach ($t in ($targets | Select-Object -Unique)) {
        $i++
        Write-Progress -Activity 'Inventário Hyper-V' -Status "Consultando $t" -PercentComplete (($i / $targets.Count) * 100)
        $results += Get-HostData -Name $t
    }
    Write-Progress -Activity 'Inventário Hyper-V' -Completed

    # ------------------------------------------------------------------
    # Saída no console
    # ------------------------------------------------------------------
    foreach ($r in $results) {
        Write-Host ""
        Write-Host ('=' * 90) -ForegroundColor Cyan
        Write-Host " SERVIDOR: $($r.Computer)" -ForegroundColor Cyan
        Write-Host ('=' * 90) -ForegroundColor Cyan

        if (-not $r.Reachable) {
            Write-Host " [X] $($r.Error)" -ForegroundColor Red
            continue
        }

        Write-Host " SO..............: $($r.OS) (build $($r.OSBuild))"
        Write-Host " Hardware........: $($r.Manufacturer) $($r.Model)  |  Serial: $($r.SerialNumber)"
        Write-Host " CPU.............: $($r.CPUModel) | $($r.CPUSockets) socket(s), $($r.CPUCoresTotal) núcleos, $($r.CPULogicalTotal) lógicos"
        Write-Host " Memória.........: $($r.RAM_TotalGB) GB total / $($r.RAM_FreeGB) GB livre"
        Write-Host " Último boot.....: $($r.LastBoot)"
        Write-Host ""
        Write-Host " Recurso Hyper-V instalado : $($r.HyperV_FeatureInstalled)"
        Write-Host " Hypervisor carregado      : $($r.HyperV_HypervisorPresent)"
        Write-Host " Serviço VMMS              : $($r.HyperV_VMMS_Status) ($($r.HyperV_VMMS_StartType))"
        Write-Host " Módulo PowerShell Hyper-V : $($r.HyperV_ModuleAvailable)"

        if ($r.HyperV_Active) {
            Write-Host " >>> HYPER-V ATIVO <<<" -ForegroundColor Green
        } else {
            Write-Host " >>> HYPER-V NÃO ATIVO neste servidor <<<" -ForegroundColor Yellow
            continue
        }

        if ($r.Error) { Write-Host " [!] Erro parcial: $($r.Error)" -ForegroundColor Red }

        if ($r.VMHost) {
            Write-Host ""
            Write-Host " Caminho padrão das VMs : $($r.VMHost.VirtualMachinePath)"
            Write-Host " Caminho padrão dos VHDs: $($r.VMHost.VirtualHardDiskPath)"
            Write-Host " Live Migration         : $($r.VMHost.LiveMigrationEnabled)"
        }

        Write-Host "`n Volumes do host:" -ForegroundColor Yellow
        $r.Volumes | Format-Table Drive, Label, FileSystem, SizeGB, FreeGB, FreePct -AutoSize | Out-String | Write-Host

        Write-Host " Switches virtuais:" -ForegroundColor Yellow
        $r.Switches | Format-Table Name, Type, NetAdapter, AllowManagementOS -AutoSize | Out-String | Write-Host

        Write-Host " VMs: $($r.VMCount) existente(s), $($r.VMsRunning) em execução" -ForegroundColor Yellow
        $r.VMs | Sort-Object State, Name | Format-Table Name, State, Generation, vCPU,
            @{N='RAM Atrib.(GB)'; E={$_.Mem_AssignedGB}},
            @{N='RAM Dinâm.'; E={$_.Mem_Dynamic}},
            @{N='Disco Máx(GB)'; E={$_.Disk_TotalMaxGB}},
            @{N='Disco Usado(GB)'; E={$_.Disk_TotalUsedGB}},
            Uptime -AutoSize | Out-String | Write-Host

        foreach ($vm in ($r.VMs | Sort-Object Name)) {
            Write-Host (" --- VM: {0}  [{1}] ---" -f $vm.Name, $vm.State) -ForegroundColor Green
            Write-Host ("   Geração {0} | Config v{1} | Cluster: {2} | Criada: {3}" -f $vm.Generation, $vm.ConfigVersion, $vm.IsClustered, $vm.Created)
            Write-Host ("   vCPU: {0} (uso {1}%, reserva {2}%, limite {3}%, peso {4}, nested: {5})" -f $vm.vCPU, $vm.CPU_UsagePct, $vm.CPU_ReservePct, $vm.CPU_LimitPct, $vm.CPU_Weight, $vm.CPU_NestedVirt)
            Write-Host ("   Memória: dinâmica={0} | inicial {1} GB | min {2} GB | máx {3} GB | atribuída {4} GB | demanda {5} GB" -f $vm.Mem_Dynamic, $vm.Mem_StartupGB, $vm.Mem_MinGB, $vm.Mem_MaxGB, $vm.Mem_AssignedGB, $vm.Mem_DemandGB)
            Write-Host ("   Config: {0}" -f $vm.ConfigPath)
            if ($vm.Generation -eq 2) { Write-Host ("   SecureBoot: {0} | TPM: {1}" -f $vm.SecureBoot, $vm.TPM) }
            Write-Host ("   Início auto: {0} | Parada auto: {1} | Checkpoint: {2}" -f $vm.AutoStartAction, $vm.AutoStopAction, $vm.CheckpointType)

            if ($vm.Disks) {
                Write-Host "   Discos:"
                $vm.Disks | Format-Table Controller, Type, Format, SizeMaxGB, SizeUsedGB, Path -AutoSize | Out-String -Width 250 | Write-Host
            }
            if ($vm.NetworkAdapters) {
                Write-Host "   Rede:"
                $vm.NetworkAdapters | Format-Table Name, Switch, MacAddress, IPAddresses, VlanMode, VlanId, Status -AutoSize | Out-String -Width 250 | Write-Host
            }
            if ($vm.Checkpoints) {
                Write-Host "   Checkpoints:"
                $vm.Checkpoints | Format-Table Name, Type, Created -AutoSize | Out-String | Write-Host
            }
            if ($vm.Replication) {
                Write-Host ("   Replicação: {0} / {1} -> {2}" -f $vm.Replication.State, $vm.Replication.Health, $vm.Replication.ReplicaServer)
            }
        }
    }

    # ------------------------------------------------------------------
    # Exportação
    # ------------------------------------------------------------------
    if (-not $NoExport) {
        if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

        # JSON completo
        $jsonFile = Join-Path $OutputPath "HyperV_Inventario_$stamp.json"
        $results | ConvertTo-Json -Depth 8 | Out-File -FilePath $jsonFile -Encoding UTF8

        # CSV - resumo por servidor
        $csvHosts = Join-Path $OutputPath "HyperV_Hosts_$stamp.csv"
        $results | Select-Object Computer, Reachable, OS, OSBuild, Manufacturer, Model, CPUModel,
            CPUSockets, CPUCoresTotal, CPULogicalTotal, RAM_TotalGB, RAM_FreeGB,
            HyperV_FeatureInstalled, HyperV_HypervisorPresent, HyperV_VMMS_Status,
            HyperV_Active, VMCount, VMsRunning, Error |
            Export-Csv -Path $csvHosts -NoTypeInformation -Encoding UTF8 -Delimiter ';'

        # CSV - resumo por VM
        $csvVMs = Join-Path $OutputPath "HyperV_VMs_$stamp.csv"
        $flatVMs = foreach ($r in $results) {
            foreach ($vm in $r.VMs) {
                [pscustomobject]@{
                    Host          = $r.Computer
                    VM            = $vm.Name
                    Estado        = $vm.State
                    Geracao       = $vm.Generation
                    vCPU          = $vm.vCPU
                    CPU_UsoPct    = $vm.CPU_UsagePct
                    MemDinamica   = $vm.Mem_Dynamic
                    Mem_InicialGB = $vm.Mem_StartupGB
                    Mem_MinGB     = $vm.Mem_MinGB
                    Mem_MaxGB     = $vm.Mem_MaxGB
                    Mem_AtribGB   = $vm.Mem_AssignedGB
                    Disco_MaxGB   = $vm.Disk_TotalMaxGB
                    Disco_UsadoGB = $vm.Disk_TotalUsedGB
                    Discos        = ($vm.Disks | ForEach-Object { $_.Path }) -join ' | '
                    Redes         = ($vm.NetworkAdapters | ForEach-Object { "$($_.Name)[$($_.Switch)] $($_.MacAddress) $($_.IPAddresses)" }) -join ' | '
                    Checkpoints   = @($vm.Checkpoints).Count
                    Uptime        = $vm.Uptime
                    Cluster       = $vm.IsClustered
                    Criada        = $vm.Created
                    PastaConfig   = $vm.ConfigPath
                }
            }
        }
        if ($flatVMs) { $flatVMs | Export-Csv -Path $csvVMs -NoTypeInformation -Encoding UTF8 -Delimiter ';' }

        # HTML
        $htmlFile = Join-Path $OutputPath "HyperV_Inventario_$stamp.html"
        $css = @'
<style>
 body{font-family:Segoe UI,Arial,sans-serif;margin:20px;color:#222}
 h1{color:#0b5394} h2{background:#0b5394;color:#fff;padding:6px 10px;border-radius:4px}
 h3{color:#0b5394;margin-bottom:4px} h4{margin:14px 0 4px;color:#38761d}
 table{border-collapse:collapse;margin-bottom:10px;font-size:12px}
 th{background:#d9e2f3;text-align:left} th,td{border:1px solid #bbb;padding:3px 8px}
 .ok{color:#38761d;font-weight:bold} .bad{color:#cc0000;font-weight:bold}
 .vm{border:1px solid #ccc;border-radius:6px;padding:8px 12px;margin:10px 0}
</style>
'@
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("<html><head><meta charset='utf-8'><title>Inventário Hyper-V</title>$css</head><body>")
        [void]$sb.AppendLine("<h1>Inventário Hyper-V</h1><p>Gerado em $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')</p>")

        foreach ($r in $results) {
            [void]$sb.AppendLine("<h2>$($r.Computer)</h2>")
            if (-not $r.Reachable) {
                [void]$sb.AppendLine("<p class='bad'>$($r.Error)</p>"); continue
            }
            $status = if ($r.HyperV_Active) { "<span class='ok'>Hyper-V ATIVO</span>" } else { "<span class='bad'>Hyper-V NÃO ativo</span>" }
            [void]$sb.AppendLine("<p>$status &mdash; $($r.OS) | $($r.CPUModel) | RAM $($r.RAM_TotalGB) GB | VMs: $($r.VMCount) ($($r.VMsRunning) em execução)</p>")
            [void]$sb.AppendLine((ConvertTo-Fragment ($r | Select-Object Computer, OS, OSBuild, Manufacturer, Model, SerialNumber, CPUModel, CPUSockets, CPUCoresTotal, CPULogicalTotal, RAM_TotalGB, RAM_FreeGB, LastBoot, HyperV_FeatureInstalled, HyperV_HypervisorPresent, HyperV_VMMS_Status, HyperV_VMCompute_Status) 'Host'))
            [void]$sb.AppendLine((ConvertTo-Fragment $r.Volumes 'Volumes'))
            if ($r.VMHost)    { [void]$sb.AppendLine((ConvertTo-Fragment $r.VMHost 'Configuração do Hyper-V')) }
            [void]$sb.AppendLine((ConvertTo-Fragment $r.Switches 'Switches virtuais'))

            if ($r.VMs.Count -gt 0) {
                [void]$sb.AppendLine((ConvertTo-Fragment ($r.VMs | Select-Object Name, State, Generation, vCPU, Mem_Dynamic, Mem_StartupGB, Mem_AssignedGB, Mem_MaxGB, Disk_TotalMaxGB, Disk_TotalUsedGB, Uptime, IsClustered) 'Resumo das VMs'))
                foreach ($vm in ($r.VMs | Sort-Object Name)) {
                    [void]$sb.AppendLine("<div class='vm'><h4>VM: $($vm.Name) [$($vm.State)]</h4>")
                    [void]$sb.AppendLine((ConvertTo-Fragment ($vm | Select-Object Generation, ConfigVersion, Created, Heartbeat, ConfigPath, CheckpointPath, AutoStartAction, AutoStopAction, CheckpointType, SecureBoot, TPM, BootOrder, Notes) 'Geral'))
                    [void]$sb.AppendLine((ConvertTo-Fragment ($vm | Select-Object vCPU, CPU_UsagePct, CPU_ReservePct, CPU_LimitPct, CPU_Weight, CPU_NestedVirt) 'Processador'))
                    [void]$sb.AppendLine((ConvertTo-Fragment ($vm | Select-Object Mem_Dynamic, Mem_StartupGB, Mem_MinGB, Mem_MaxGB, Mem_AssignedGB, Mem_DemandGB, Mem_BufferPct, Mem_Priority) 'Memória'))
                    [void]$sb.AppendLine((ConvertTo-Fragment $vm.Disks 'Discos'))
                    [void]$sb.AppendLine((ConvertTo-Fragment $vm.DvdDrives 'DVD/ISO'))
                    [void]$sb.AppendLine((ConvertTo-Fragment $vm.NetworkAdapters 'Rede'))
                    [void]$sb.AppendLine((ConvertTo-Fragment $vm.Checkpoints 'Checkpoints'))
                    [void]$sb.AppendLine((ConvertTo-Fragment $vm.IntegrationServices 'Serviços de integração'))
                    if ($vm.Replication) { [void]$sb.AppendLine((ConvertTo-Fragment $vm.Replication 'Replicação')) }
                    [void]$sb.AppendLine("</div>")
                }
            }
        }
        [void]$sb.AppendLine("</body></html>")
        $sb.ToString() | Out-File -FilePath $htmlFile -Encoding UTF8

        Write-Host "`nRelatórios gerados em: $OutputPath" -ForegroundColor Cyan
        Write-Host "  - $jsonFile"
        Write-Host "  - $csvHosts"
        if ($flatVMs) { Write-Host "  - $csvVMs" }
        Write-Host "  - $htmlFile"
    }

    # Retorna os objetos para uso em pipeline
    return $results
}
