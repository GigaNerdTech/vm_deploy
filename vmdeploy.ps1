# VM Deploy Interface
# Written by Joshua Woleben
# Started 7/24/19
# Version 1.0

$VerbosePreference="Continue"
# Load PowerCLI

if (Get-Module -ListAvailable -Name VMware.VimAutomation.Core) {
    Import-Module VMware.VimAutomation.Core
}
else {
    if (Get-PSRepository | Where-Object { $_ -match "internal-repo" }) {
        Install-Module -Name VMware.PowerCLI -Repository "internal-repo" -AllowClobber
        Import-Module VMware.VimAutomation.Core
    }
    else {
        Register-PSRepository -Name 'internal-repo' -SourceLocation 'internal-repo' -InstallationPolicy Trusted
        Install-Module -Name VMware.PowerCLI -Repository "internal-repo" -AllowClobber
        Import-Module VMware.VimAutomation.Core
    }
}
Import-Module ActiveDirectory
Import-Module "\\networkshare\Powershell\vmdeploy_functions.psm1"

# Variables
$vcenter_options = @("vcenter1","vcenter2")
$cpu_options = @("1","2","4","8")
$memory_options = @("2","4","8","16","24","32","48","64")
$role_options = @("Standard","SQL 2014 Standard","SQL 2014 Enterprise","SQL 2016 Standard","SQL 2016 Enterprise","Workstation")
$script:tag_category = "Provision"
$script:tag_name = "AllowProvisioning"
$script:current_vcenter_host = ""
$script:current_datacenter = ""
$script:current_datastore = ""
$script:current_portgroup = ""
$script:current_template = ""
$script:current_role = ""
$script:current_cpu_count = ""
$script:current_memory = ""

# Create credential objects for all layers
$vsphere_user = Read-Host -Prompt "Enter username for vcenter"
$vsphere_pwd = Read-Host -Prompt "Enter password for vcenter" -AsSecureString
$script:vsphere_creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $vsphere_user,$vsphere_pwd -ErrorAction Stop

# Individual functions
function clear_form {
    $script:current_vcenter_host = ""
    $script:current_datacenter = ""
    $script:current_datastore = ""
    $script:current_portgroup = ""
    $script:current_template = ""
    $script:current_role = ""
    $script:current_cpu_count = ""
    $script:current_memory = ""

    $vCenterSelect.Items.Clear()
    $TemplateSelect.Items.Clear()
    $HostnameBox.Text = ""
    $IPAddressBox.Text = ""
    $DatacenterSelect.Items.Clear()
    $NotesBox.Text = ""
    $ClusterSelect.Items.Clear()
    $PortGroupSelect.Items.Clear()
    $DatastoreSelect.Items.Clear()
    $global:Form.InvalidateVisual()

    Write-Host "Items Cleared!"
}

# GUI Code
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="[VM Deploy Tool]" Height="500" Width="450" MinHeight="768" MinWidth="400" ResizeMode="CanResizeWithGrip">
    <StackPanel>
        <Label x:Name="vCenterLabel" Content="vCenter"/>
        <ComboBox x:Name="vCenterSelect"/>
        <Label x:Name="TemplateLabel" Content="Template"/>
        <ComboBox x:Name="TemplateSelect"/>
        <Label x:Name="HostnameLabel" Content="Hostname"/>
        <TextBox x:Name="HostnameBox"/>
        <Label x:Name="NoCPULabel" Content="# CPU"/>
        <ComboBox x:Name="CPUSelect"/>
        <Label x:Name="MemoryLabel" Content="Memory (GB)"/>
        <ComboBox x:Name="MemorySelect"/>
        <Label x:Name="IPAddressLabel" Content="IP Address"/>
        <TextBox x:Name="IPAddressBox"/>
        <Label x:Name="RoleLabel" Content="Role"/>
        <ComboBox x:Name="RoleSelect"/>
        <Label x:Name="DatacenterLabel" Content="Datacenter"/>
        <ComboBox x:Name="DatacenterSelect"/>
        <Label x:Name="ClusterLabel" Content="Cluster"/>
        <ComboBox x:Name="ClusterSelect"/>
        <Label x:Name="PortGroupLabel" Content="PortGroup"/>
        <ComboBox x:Name="PortGroupSelect"/>
        <Label x:Name="DatastoreLabel" Content="Datastore"/>
        <ComboBox x:Name="DatastoreSelect"/>
        <Label x:Name="NotesLabel" Content="Notes"/>
        <TextBox x:Name="NotesBox" Height="100"/>
        <Button x:Name="DeployButton" Content="Deploy VM" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
    </StackPanel>
</Window>
'@
 
$global:Form = ""
# XAML Launcher
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$global:Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; break}
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $global:Form.FindName($_.Name)}

# Set up controls

$vCenterSelect = $global:Form.FindName('vCenterSelect')
$TemplateSelect = $global:Form.FindName('TemplateSelect')
$HostnameBox = $global:Form.FindName('HostnameBox')
$CPUSelect = $global:Form.FindName('CPUSelect')
$MemorySelect = $global:Form.FindName('MemorySelect')
$IPAddressBox = $global:Form.FindName('IPAddressBox')
$RoleSelect = $global:Form.FindName('RoleSelect')
$DatacenterSelect = $global:Form.FindName('DatacenterSelect')
$ClusterSelect = $global:Form.FindName('ClusterSelect')
$PortGroupSelect = $global:Form.FindName('PortGroupSelect')
$DatastoreSelect = $global:Form.FindName('DatastoreSelect')
$NotesBox = $global:Form.FindName('NotesBox')
$DeployButton = $global:Form.FindName('DeployButton')

# Add CPU options
foreach ($cpu_num in $cpu_options) {
    $CPUSelect.Items.Add($cpu_num) | out-null
}

# Add Memory Options
foreach ($mem_num in $memory_options) {
    $MemorySelect.Items.Add($mem_num) | out-null
}

# Add vCenter options
foreach ($vc_opt in $vcenter_options) {
    $vCenterSelect.Items.Add($vc_opt) | out-null
}
foreach ($role_opt in $role_options) {
    $RoleSelect.Items.Add($role_opt) | out-null
}

# Event controls

# When vcenter selection changes!
$vCenterSelect.Add_SelectionChanged({
    
    # Disconnect existing connections
    Disconnect-VIServer -Server * -Confirm:$false -Force -ErrorAction SilentlyContinue

    # Set current vcenter server
    $script:current_vcenter_host = $vCenterSelect.SelectedItem.ToString()

    # Connect to specified vcenter server
    Connect-VIServer -Server $script:current_vcenter_host -Credential $script:vsphere_creds -ErrorAction Stop

    Write-Host "Tag name: $script:tag_name`nTag cat: $script:tag_category"
    # Get tag for provisioning
    $script:tag = Get-Tag -Category $script:tag_category -Name $script:tag_name -ErrorAction SilentlyContinue

    # Populate datacenters
    $DatacenterSelect.Items.Clear()
    Get-Datacenter -Server $script:current_vcenter_host | Where { ($_ | Get-TagAssignment).Tag.Name -contains $script:tag_name }  | Select -ExpandProperty Name | ForEach-Object { $DatacenterSelect.Items.Add($_) }

    # Populate templates
    $TemplateSelect.Items.Clear()
    Get-Template -Server $script:current_vcenter_host | Where { ($_ | Get-TagAssignment).Tag.Name -contains $script:tag_name } | Select -ExpandProperty Name | ForEach-Object { $TemplateSelect.Items.Add($_) }

    # Clear other lists
    $ClusterSelect.Items.Clear()
    $PortGroupSelect.Items.Clear()
    $DatastoreSelect.Items.Clear()

    # Clear variables
    $script:current_cluster = ""
    $script:current_datastore = ""
    $script:current_portgroup = ""
    $script:current_datacenter= ""
    $script:current_template = ""

    # Repaint form
    $global:Form.InvalidateVisual()

})
# When template selection changes!
$TemplateSelect.Add_SelectionChanged({
    $script:current_template = $TemplateSelect.SelectedItem.ToString()

})

# When Role selection changes!
$RoleSelect.Add_SelectionChanged({

    $script:current_role = $RoleSelect.SelectedItem.ToString()
})
# When Datacenter selection changes!
$DatacenterSelect.Add_SelectionChanged({
    # Set current datacenter
    $script:current_datacenter = $DatacenterSelect.SelectedItem.ToString()

    # Get tag for provisioning
    $script:tag = Get-Tag -Category $script:tag_category -Name $script:tag_name -ErrorAction SilentlyContinue
    # Populate clusters
    $ClusterSelect.Items.Clear()
    Get-Cluster -Server $script:current_vcenter_host -Location $script:current_datacenter | Where { ($_ | Get-TagAssignment).Tag.Name -contains $script:tag_name } | Select -ExpandProperty Name | ForEach-Object { $ClusterSelect.Items.Add($_) }

    # Clear other lists

    $PortGroupSelect.Items.Clear()
    $DatastoreSelect.Items.Clear()

    # Clear variables
    $script:current_cluster = ""
    $script:current_datastore = ""
    $script:current_portgroup = ""

    # Repaint form
    $global:Form.InvalidateVisual()

})

# When CPU selection changes!
$CPUSelect.Add_SelectionChanged({
    $script:current_cpu_count = $CPUSelect.SelectedItem.ToString()
})

# When memory selection changes!
$MemorySelect.Add_SelectionChanged({
    $script:current_memory = $MemorySelect.SelectedItem.ToString()
})

# When cluster selection changes!
$ClusterSelect.Add_SelectionChanged({
    # Set current cluster
    $script:current_cluster = $ClusterSelect.SelectedItem.ToString()
    
    # Clear appropriate lists
    $PortGroupSelect.Items.Clear()
    $DatastoreSelect.Items.Clear()

    # Clear variables
    $script:current_datastore = ""
    $script:current_portgroup = ""

    # Get tag for provisioning
    $script:tag = Get-Tag -Category $script:tag_category -Name $script:tag_name -ErrorAction SilentlyContinue

    # Populate Portgroups and datastores
    $cluster = ""
    $cluster = Get-Cluster -Name $script:current_cluster

    $datacenter = ""
    $datacenter = Get-Datacenter -Name $script:current_datacenter

    $cluster | Get-VMHost | Get-VDSwitch -Server $script:current_vcenter_host | Get-VDPortgroup -Server $script:current_vcenter_host |  Where { ($_ | Get-TagAssignment).Tag.Name -contains $script:tag_name } | Select -ExpandProperty Name | ForEach-Object { $PortGroupSelect.Items.Add($_) }
    $cluster | Get-VMHost | Get-VirtualSwitch -Standard -Server $script:current_vcenter_host | Get-VirtualPortgroup -Standard -Server $script:current_vcenter_host |  Where { ($_ | Get-TagAssignment).Tag.Name -contains $script:tag_name } | Select -ExpandProperty Name | ForEach-Object { $PortGroupSelect.Items.Add($_) }
    $cluster | Get-Datastore -Server $script:current_vcenter_host | Get-DatastoreCluster -Server $script:current_vcenter_host |  Where { ($_ | Get-TagAssignment).Tag.Name -contains $script:tag_name } | Select -ExpandProperty Name | ForEach-Object { $DatastoreSelect.Items.Add($_) }
    $cluster | Get-Datastore -Server $script:current_vcenter_host | Where { ($_ | Get-TagAssignment).Tag.Name -contains $script:tag_name } | Select -ExpandProperty Name | ForEach-Object { $DatastoreSelect.Items.Add($_) }
    # Repaint form
    $global:Form.InvalidateVisual()
})

# When Portgroup selection changes!
$PortGroupSelect.Add_SelectionChanged({
    $script:current_portgroup = $PortGroupSelect.SelectedItem.ToString()
})

# When Datastore selection changes!
$DatastoreSelect.Add_SelectionChanged({
    $script:current_datastore = $DatastoreSelect.SelectedItem.ToString()
})

# Button clicked!

$DeployButton.Add_Click({
    

    # Get variables not set by dropdown boxes
    $script:current_hostname = $HostnameBox.Text
    $script:current_ip_address = $IPAddressBox.Text
    $script:current_notes = $NotesBox.Text

    Write-Host "Deploy commencing..."

    Write-Host "Options Selected:`nvCenter: $script:current_vcenter_host`nTemplate: $script:current_template`nHostname: $script:current_hostname`nCPU count: $script:current_cpu_count"
    Write-Host "Current Memory: $script:current_memory GB`nIP Address: $script:current_ip_address`nRole: $script:current_role`nDatacenter: $script:current_datacenter"
    Write-Host "Cluster: $script:current_cluster`nPortGroup: $script:current_portgroup`nDatastore: $script:current_datastore`nNotes: $script:current_notes"

    # Error checking
    if ($script:current_cpu_count -eq "") {
        [System.Windows.MessageBox]::Show('CPU count required!')
        exit
    }
    if ($script:current_memory -eq "") {
        [System.Windows.MessageBox]::Show('Memory required!')
        exit
    }
    if ($script:current_vcenter_host -eq "") {
        [System.Windows.MessageBox]::Show('vCenter host required!')
        exit
    }
    if ($script:current_datacenter -eq "") {
        [System.Windows.MessageBox]::Show('Datacenter required!')
        exit
    }
    if ($script:current_hostname -eq "") {
        [System.Windows.MessageBox]::Show('Hostname required!')
        exit
    }
    if ($script:current_ip_address -eq "") {
        $script:current_ip_address = "AUTO"
    }
    if ($script:current_role -eq "") {
        [System.Windows.MessageBox]::Show('Role required!')
        exit
    }
    if ($script:current_cluster -eq "") {
        [System.Windows.MessageBox]::Show('Cluster required!')
        exit
    }
    if ($script:current_portgroup -eq "") {
        [System.Windows.MessageBox]::Show('Portgroup required!')
        exit
    }
    if ($script:current_datastore -eq "") {
        [System.Windows.MessageBox]::Show('Datastore required!')
        exit
    }
    # Notes are optional, so this is commented out
    #if ($script:current_notes -eq "") {
    #    [System.Windows.MessageBox]::Show('Notes required!')
    #    exit
    #}
    if(($script:current_ip_address -notmatch "\d+\.\d+\.\d+\.\d+") -and ($script:current_ip_address -ne "AUTO")) {
        [System.Windows.MessageBox]::Show('IP Address not formatted properly!')
        exit
    }
    # Deploy code
    Set-PowerCLIConfiguration -Scope Session -WebOperationTimeoutSeconds 180 -Confirm:$false
    $script_block = { Param([string]$current_vcenter,
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

    Import-Module VMware.VimAutomation.Core
    Import-Module ActiveDirectory
    Import-Module "\\funzone\team\POSH\Powershell\vmdeploy_functions.psm1"
    vm_deploy $current_vcenter $current_template $current_hostname $current_cpu_count $current_memory $current_ip_address $current_role $current_datacenter $current_cluster $current_portgroup $current_datastore $current_notes 
    }
    
    $job = Start-Job -ArgumentList $script:current_vcenter_host, $script:current_template, $script:current_hostname, $script:current_cpu_count,
    $script:current_memory, $script:current_ip_address, $script:current_role, $script:current_datacenter, $script:current_cluster, $script:current_portgroup,
    $script:current_datastore, $script:current_notes -ScriptBlock $script_block

    [System.Windows.Forms.Application]::DoEvents() 
    clear_form
    Write-Host "Completed! See email for logs."
    # Extra tasks
})

# Show GUI
$global:Form.ShowDialog() | out-null

$global:Form.Close()