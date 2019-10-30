Import-Module VMware.VimAutomation.Core
Import-Module ActiveDirectory

$VerbosePreference="Continue"
$WarningPreference = "SilentlyContinue"

function Test-Ping {
    param(
        # ComputerName
        [Parameter(Mandatory = $true)]
        [string]
        $ComputerName
    )

    # Try pinging and return status.
    $ping_status = $false
    & ipconfig /flushdns | Out-Null
    Start-Sleep -Seconds 1
    $ping = Get-WmiObject -Query "select * from Win32_PingStatus where Address='$ComputerName'"
    if ($ping.StatusCode -eq 0) {
        $ping_status = $true
    }
    return $ping_status
}
function Get-NextAvailableIP {
    param(
        # Port Group name
        [Parameter(Mandatory = $true)]
        [string]
        $PortGroupName,

        # Switch Type, Standard or Distributed
        [Parameter(Mandatory = $true)]
        [ValidateSet('Standard', 'Distributed')]
        [string]
        $SwitchType
    )

    # $ip_regex = '^(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9])' +
    # '\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])' +
    # '\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])' +
    # '\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])'

     $netmask_table = @{
         '255.255.0.0'     = 65534;
         '255.255.128.0'   = 32766;
         '255.255.192.0'   = 16382;
         '255.255.224.0'   = 8190;
         '255.255.240.0'   = 4094;
         '255.255.248.0'   = 2046;
         '255.255.252.0'   = 1022;
         '255.255.254.0'   = 510;
         '255.255.255.0'   = 254;
         '255.255.255.128' = 126;
         '255.255.255.192' = 62;
         '255.255.255.224' = 30;
         '255.255.255.240' = 14;
         '255.255.255.248' = 6;
         '255.255.255.252' = 2;
     }

    if ($SwitchType -match 'distributed') {
        $pg = Get-VDPortGroup -Name $PortGroupName
    }
    else {
        $pg = Get-VirtualPortGroup -Name $PortGroupName
    }

    # Retrieve the Custom Value data from PG
    $network = $pg.ExtensionData.AvailableField | Where-Object {$_.Name -eq 'Network'}
    $network = $pg.ExtensionData.CustomValue | Where-Object {$_.Key -eq $network.Key}
    $network = $network.Value -replace '\.0', ''

    $netmask = $pg.ExtensionData.AvailableField | Where-Object {$_.Name -eq 'Netmask'}
    $netmask = $pg.ExtensionData.CustomValue | Where-Object {$_.Key -eq $netmask.Key}
    $netmask = $netmask.Value

    # Reserve 1-9 for network team.
    $reserved_range = 1..9
    # Reserve broadcast address.

    # Get existing IP list from Guest Tools and sort
    $vm_list = $pg | Get-VM
    $ip_list = $vm_list.ExtensionData.Guest.IpAddress | Where-Object {$_ -match $network}
    $ip_list = $ip_list | ForEach-Object {[System.Net.IPAddress]::Parse($_)} | Sort-Object Address | ForEach-Object {$_.IPAddressToString}

    for ($x = 1; $x -le $netmask_table[$netmask]; $x++) {
        $ip = [System.Net.IPAddress]::Parse($network + "." + $x)
        $ip_str = $ip.IPAddressToString
        $test = ($ip_str -split "\.")[3]
        if ($reserved_range -notmatch $test) {
            if ($ip_list -notcontains $ip_str) {
                if (-not (Test-Ping -ComputerName $ip_str)) {
                    $new_ip = $ip_str
                    break
                }
            }
        }
    }

    $new_ip
}
function Get-PGNetworkInfo {
    # Map custom attributes to network information
    # exit object with network information stored
    Param(
        # PortGroup Name
        [Parameter(Mandatory = $true)]
        [string]
        $PortGroupName
    )
    $network_obj = New-Object PSObject
    $pg = Get-VDPortGroup -Name $PortGroupName

    foreach ($field in $pg.ExtensionData.AvailableField) {
        switch ($field.Name) {
            "Netmask" { $network_obj | Add-Member NoteProperty netmask ($pg.ExtensionData.CustomValue | Where-Object {$_.Key -eq $field.Key}).Value }
            "Gateway" { $network_obj | Add-Member NoteProperty gateway ($pg.ExtensionData.CustomValue | Where-Object {$_.Key -eq $field.Key}).Value }
            "DNS1" { $network_obj | Add-Member NoteProperty dns1 ($pg.ExtensionData.CustomValue | Where-Object {$_.Key -eq $field.Key}).Value }
            "DNS2" { $network_obj | Add-Member NoteProperty dns2 ($pg.ExtensionData.CustomValue | Where-Object {$_.Key -eq $field.Key}).Value }
        }
    }

    return $network_obj
}
function vm_deploy {
    Param([string]$current_vcenter,
        [string]$current_template,
        [string]$current_hostname,
        [string]$current_cpu_count,
        [string]$current_memory,
        [string]$current_ip_address,
        [string]$current_role,
        [string]$current_datacenter,
        [string]$current_cluster,
        [string]$current_portgroup,
        [string]$current_datastore,
        [string]$current_notes

    )
    # Start transcript logging
    $TranscriptFile = "\\networkshare\PSLogs\vmdeploy_$(get-date -f MMddyyyyHHmmss).txt"
    Start-Transcript -Path $TranscriptFile
                $os_table = @{
                    'windows9_64Guest'      = 'WorkstationWindows';
                    'windows7_64Guest'      = 'WorkstationWindows';
                    'windows9Server64Guest' = 'ServerWindows';
                    'windows8Server64Guest' = 'ServerWindows';
                    'centos7_64Guest'       = 'ServerLinux';
                    'rhel7_64Guest'         = 'ServerLinux';
                }


                # Create credential objects for all layers
                $vsphere_user = "Administrator@local"
                $vsphere_pwd = ConvertTo-SecureString -String 'somepassword' -AsPlainText -Force
                $vsphere_creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $vsphere_user,$vsphere_pwd -ErrorAction Stop
                 # Connect to specified vcenter server
                Connect-VIServer -Server $current_vcenter -Credential $vsphere_creds -ErrorAction Stop

                Set-PowerCLIConfiguration -Scope Session -WebOperationTimeoutSeconds 180 -Confirm:$false

                $os_specification_windows = "Win_Domain"
                $os_specification_linux = "Lin_Domain"

                $Requestor = $env:USERNAME

                ####### Get template info #######
                Write-Host "Retrieving template information..."
    try {
                $vm_template = Get-Template -Name $current_template
                $vm_template_id = $vm_template.ExtensionData.Config.GuestId
                $os_type = $os_table[$vm_template_id]
                if ($os_type -imatch 'Linux') {
                    $os_specification = $os_specification_linux
                }
                else {
                    $os_specification = $os_specification_windows
                }
                $os_spec_source = Get-OSCustomizationSpec -Name $os_specification
                $os_spec_clone = New-OSCustomizationSpec -Name ("VMD_" + (Get-Date).Ticks) -OSCustomizationSpec $os_spec_source -Type NonPersistent
        }
        catch {
                Write-Host "Unable to retrieve template info!"
                exit
        }
                ####### Configure network #######
        try {
                Write-Host "Configuring the network..."

                # Get extra network config info
                $network_info = Get-PGNetworkInfo $current_portgroup
                # Get IP if AUTO
                if ($current_ip_address -imatch 'AUTO' -and $os_type -notmatch "Workstation") {
                    $current_ip_address = Get-NextAvailableIP -PortGroupName $current_portgroup -SwitchType 'Distributed'
                }
                # Create NIC profile(spec)
                # Remove initial NIC profile from OS profile
                Write-Host "Creating NIC profile..."
                
                Remove-OSCustomizationNicMapping -OSCustomizationNicMapping (Get-OSCustomizationNicMapping -OSCustomizationSpec $os_spec_clone) -Confirm:$false
                if ($current_role -match "Workstation" -or $current_role -match "Standard") {
                    New-OSCustomizationNicMapping -OSCustomizationSpec $os_spec_clone -IpMode UseDhcp -Confirm:$false
                }
                else {
                    New-OSCustomizationNicMapping -OSCustomizationSpec $os_spec_clone -IpMode UseStaticIP -IpAddress $current_ip_address -SubnetMask $network_info.netmask -DefaultGateway $network_info.gateway -Dns @($network_info.dns1, $network_info.dns2) -Confirm:$false
                }
           }
            catch {
               Write-Host "Unable to configure network!"
               exit
           }

                ####### Get VMHost Info #######
                Write-Host "Retrieving VMHost information..."
        try {
                $vm_host = $null
                if (Get-Cluster -name $current_cluster -erroraction SilentlyContinue) {
                    Write-Host "Finding low utilization host in cluster: $current_cluster"
                    $vm_host = (Get-Cluster -Name $current_cluster | Get-VMHost | Where-Object {$_.ConnectionState -eq 'Connected'} | Sort-Object -Property MemoryUsageGB)[0]
                    Write-Host ("Found host: " + $vm_host.Name)
                }
                else {
                    $vm_host = Get-VMHost -Name $current_cluster
                }
        }
        catch {
                Write-Host "Failed to retrieve VMHost information!"
                exit
        }
                ##### Get Datastore info ######

                Write-Host "Retrieving datastore information..."
        try {
                $vm_datastore = $null
                if (Get-DatastoreCluster -Name $current_datastore -ErrorAction SilentlyContinue) {
                    $vm_datastore = Get-DatastoreCluster -Name $current_datastore
                }
                else {
                    $vm_datastore = Get-Datastore -Name $current_datastore
                }
        }
        catch {
               Write-Host "Failed to retrieve datastore information!"
               exit
        }
                ##### Clone VM #####
                Write-Host "Cloning VM from template..."
        try {

                New-VM -VMHost $vm_host -Name $current_hostname -Datastore $vm_datastore -Template $vm_template -OSCustomizationSpec $os_spec_clone -Notes "$current_notes" -Confirm:$false 
        }
        catch {
               Write-Host "Failed to clone VM!"
               exit
        }

                ##### Configure VM #####
                Write-Host "Configuring VM..."
        try {
                $vm = Get-VM $current_hostname

                $vm.ExtensionData.setCustomValue("CreatedBy", $Requestor)
                $vm.ExtensionData.setCustomValue("CreatedDate", (Get-Date -Format 'yyy-MM-dd HH:mm:ss'))

                # Set CPU and Memory
                # Calculate Cores Per Socket, default = vm_num_cpu
                Write-Host "Setting CPU and RAM..."
                $vm_num_cpu = $current_cpu_count
                $vm_cores_ps = $vm_num_cpu
                if ($vm_num_cpu -gt 4 -and ($vm_num_cpu % 4) -eq 0) {
                    $vm_cores_ps = 4
                }
                Set-VM -VM $vm -NumCpu $vm_num_cpu -CoresPerSocket $vm_cores_ps -MemoryGB $current_memory -Confirm:$false | Out-Null
        }
        catch {
                Write-Host "Failed to set CPU and RAM!"
                exit
        }
                # Get portgroup to attach
                # Connect VM to portgroup
                Write-Host "Connecting to configured PortGroup..."

        try {

                $vm_pg = Get-VDPortgroup -Name $current_portgroup
                $vm_nic = Get-NetworkAdapter -VM $vm
                Set-NetworkAdapter -NetworkAdapter $vm_nic -Portgroup $vm_pg -Confirm:$false
        }
        catch {
                Write-Host "Failed to set network adapter!"
                exit
        }
                ##### Boot VM #####
                Write-Host "Powering on VM..."
        try {
               
                Start-VM -VM $vm -Confirm:$false | Out-Null
        }
        catch {
                Write-Host "Failed to power on VM!"
                exit
        }
        # Provision DNS
        try {
             Add-DnsServerResourceRecordA -ZoneName 'example.int' -Name $current_hostname -IPv4Address "$current_ip_address" -AllowUpdateAny -CreatePtr -ComputerName 'enterprise.example.int' -ErrorAction SilentlyContinue
             Add-DnsServerResourceRecordA -ZoneName 'example.int' -Name $current_hostname -IPv4Address "$current_ip_address" -AllowUpdateAny -CreatePtr -ComputerName 'example-dal-dc01.example.int' -ErrorAction SilentlyContinue
             }
             catch {
                        Write-Host "Failed to create DNS! Continuing..."
             }


                # Wait for VM to finish initial config
                Write-Host "Waiting for image scripts to complete..."

                # Flush DNS and wait
                & ipconfig.exe /flushdns
                Start-Sleep -Seconds 15

                # Initial Customization
                
                # Test VM is up. Wait a few seconds, flush dns, ping
                $counter = 900
                do {
                    Start-Sleep -Seconds 10
                    $counter = $counter - 10
                    Write-Host "$counter seconds left..."
                }
                while (-not (Test-Ping -ComputerName $current_hostname) -and $counter -gt 0)

                #### VM Post-boot configuration #####
                Write-Host "Running default configuration script..."

                # Windows Post-Config
         if ($os_type -imatch 'Windows') {
         Write-Host "Entering Windows script..."

                if ($os_type -imatch 'Server') {
                    Write-Host "Entering Server script..."
                    $counter = 0
                    do {
                        Start-Sleep -Seconds 10
                        $ad_obj = Get-ADComputer $current_hostname -ErrorAction SilentlyContinue
                        if ($ad_obj) {
                             Write-Host "AD Object Found"
                            if ($ad_obj.DistinguishedName -notmatch "OU=Servers,DC=example,DC=com") {
                               Write-Host "Moving AD Object"
                               Move-ADObject -Identity $ad_obj -TargetPath "OU=Servers,DC=example,DC=com" -Confirm:$false
                             }
                        }
                        $counter++
                    }
                    while (-not $ad_obj.SID -and $counter -lt 12)

             ##### ROLE SPECIFIC TASKS #####
             if ($current_role -match "SQL") {
                   Write-Host "No SQL tasks currently defined, this is a placeholder for later."
             }

             if ($current_role -match "Workstation") {
                Write-Host "Entering workstation script..."
                do {
                    Start-Sleep -Seconds 10
                    $ad_obj = Get-ADComputer $current_hostname -ErrorAction SilentlyContinue
                    $ad_ou = "OU=Computers,OU=Corporate,DC=example,DC=com"
                    if ($ad_obj) {
                        if ($ad_obj.DistinguishedName -notmatch $ad_ou) {
                                Write-Host "Moving AD Object"
                                Move-ADObject -Identity $ad_obj -TargetPath $ad_ou -Confirm:$false
                         }
                    }
                    $counter++
                }
                while (-not $ad_obj.SID -and $counter -lt 12)

             
                Write-Host "Waiting for workstation to boot..."
                $counter = 0
                do {
                    Start-Sleep -Seconds 10
                    $result = Test-Path -Path "\\$current_hostname\c`$" -ErrorAction SilentlyContinue
                    $counter++
                }
                while (-not $result -and $counter -lt 6)

                $session = New-PSSession -ComputerName $current_hostname -EnableNetworkAccess
                $script_block = {
   

                    & net use Z: "\\example-sccm02\SMS_example\Client" /USER:example\is.sccmservice somepassword
                    $install_string = 'Z:\ccmsetup.exe'
                    
                    Invoke-Expression $install_string

                    # Wait for install to finish
                    
                    Start-Sleep -Seconds 10
                    $counter = 0
                    do { Start-Sleep -Seconds 10; $counter++ }
                    while ((Get-Process | Where-Object { $_.ProcessName -like "ccmsetup" }) -and $counter -lt 18)

                    
                    & net use Z: /DELETE
                }
                Write-Host "Installing CCM..."
                
                Invoke-Command -Session $session -ScriptBlock $script_block

                Remove-PSSession $session
             }
        }
    }
    Write-Host "Completed!"

        # Generate email report
    $email_list=@("Someteam@example.com")
    $subject = "VM Deploy Report for: $current_hostname"

    $body = "VM Deploy report attached!`nVM Deployed: $current_hostname`nRole: $current_role`nCPUs: $current_cpu_count`nRAM (GB): $current_memory`nRequestor: $Requestor"

    Stop-Transcript

    $MailMessage = @{
        To = $email_list
        From = "VMDeployReport<Donotreply@example.com>"
        Subject = $subject
        Body = $body
        SmtpServer = "smtp.mhd.com"

        Attachment = $TranscriptFile
    }
    Send-MailMessage @MailMessage
}