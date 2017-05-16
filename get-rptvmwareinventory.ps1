<#
=======================================================================================
File Name: get-rptvmwareinventory.ps1
Created on: 2017-05-16
Created with VSCode
Version 1.0
Last Updated: 
Last Updated by: John Shelton | c: 260-410-1200 | e: john.shelton@lucky13solutions.com

Purpose: Generate a report of all VMWare VMs, their datastores, and test connectivity.
         

Notes: By default the script will not test connectivity unless you pass the -TestConnectivity
       parameter.

Change Log:


=======================================================================================
#>
#
# Define Parameter(s)
#
param (
  [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
  [string[]] $VCenterServers = $(throw "-VCenterServers is required.  Pass as array."),
  [switch]$TestConnectivity
)
#
Clear-Host
#
# Load VMWare PSSnapin
#
Add-PSSnapin VMWare.VimAutomation.Core
Set-PowerCLIConfiguration -InvalidCertificateAction "Ignore" -DisplayDeprecationWarnings $False -Confirm $False
#
# Define Output Variables
#
$ExecutionStamp = Get-Date -Format yyyyMMdd_hh-mm-ss
$path = "c:\temp\"
$FilenamePrepend = 'rpt_'
$FullFilename = "get-rptvmwareinventory.ps1"
$FileName = $FullFilename.Substring(0, $FullFilename.LastIndexOf('.'))
$FileExt = '.xlsx'
$OutputFile = $path + $FilenamePrePend + '_' + $FileName + '_' + $ExecutionStamp + $FileExt
$PathExists = Test-Path $path
IF($PathExists -eq $False)
  {
  New-Item -Path $path -ItemType  Directory
  }
#
$VMDiskInfo = @()
$CountVCenterServers = $VCenterServers.Count
# $PercentVCenterServers = 0
# $VCenterServersProcessed = 0
$VMsProcessed = 0
$CompleteVMInfo = @()
$AllVMInfo = @()
connect-viserver $VCenterServers
$VMHosts = Get-VMHost | Where-Object {$_.ConnectionState -eq "Connected"}
$VMs = Get-VM
$CountVMs = $VMs.Count
Clear-Host
ForEach ($VM in $VMs){
  $VMsProcessed++
  $PercenetVMsProcessed = ($VMsProcessed/$CountVMs*100)
  Write-Progress -Activity "Processing through all VMs" -PercentComplete $PercenetVMsProcessed -CurrentOperation "Processing $VM"
  $VMDNS = Resolve-DnsName $VM.Name -ErrorAction SilentlyContinue
  IF(!$VMDNS.Name) {$VMConnected = "No DNS Name Found"}
  Else {IF($TestConnectivity -eq $true){$VMConnected = Test-Connection $VMDNS.Name -Count 1 -Quiet}Else {$VMConnected = "Test Connectivy Skipped"}}
  $VMDatastores = Get-Datastore -RelatedObject $VM
  $VMDatastoresNames = [system.String]::Join(" | ",$VMDatastores.Name)
  $TempDataStoreCluster = Get-VM $VM.Name | Get-DatastoreCluster
  IF(!$TempDataStoreCluster) {$TempDataStoreCluster = "Datastore is not defined in a DataStore Cluster on this Host"}
  $TempVMHost = $VM.VMHost.Name
  $TempVCenter = $VM.VMHost.Client.ServerUri.Split('@')[1]
  $results = New-Object psobject
  $results | Add-Member -MemberType NoteProperty -Name "VCenterServer" -Value $TempVCenter
  $results | Add-Member -MemberType NoteProperty -Name "Name" -Value $VM.Name
  $results | Add-Member -MemberType NoteProperty -Name "VMHost" -Value $TempVMHost
  $results | Add-Member -MemberType NoteProperty -Name "DNS Name" -Value $VMDNS.Name
  $results | Add-Member -MemberType NoteProperty -Name "IP" -Value $VMDNS.IPAddress
  $results | Add-Member -MemberType NoteProperty -Name "RepliedToPing" -Value $VMConnected
  $results | Add-Member -MemberType NoteProperty -Name "VMWareFolder" -Value $VM.Folder
  $results | Add-Member -MemberType NoteProperty -Name "VMPowerState" -Value $VM.PowerState
  $results | Add-Member -MemberType NoteProperty -Name "VMGuestInfo" -Value $VM.Guest
  $results | Add-Member -MemberType NoteProperty -Name "CPUs" -Value $VM.NumCpu
  $results | Add-Member -MemberType NoteProperty -Name "Memory GB" -Value $VM.MemoryGB
  $results | Add-Member -MemberType NoteProperty -Name "VM Version" -Value $VM.Version
  $results | Add-Member -MemberType NoteProperty -Name "UsedSpaceGB" -Value $VM.UsedSpaceGB
  $results | Add-Member -MemberType NoteProperty -Name "ProvisionedSpaceGB" -Value $VM.ProvisionedSpaceGB
  $results | Add-Member -MemberType NoteProperty -Name "Datastores" -Value $VMDatastoresNames
  $results | Add-Member -MemberType NoteProperty -Name "DatastoreCluster" -Value $TempDataStoreCluster
  $AllVMInfo += $results
  $CompleteVMInfo += $results
  ForEach ($DataStore in $VMDatastores){
    $VMDisks = Get-HardDisk -VM $VM
    ForEach ($VMDisk in $VMDisks){
      $RegExFindBracket = "\[(.*?)\]"
      $TempDataStore = $VMDisk.Filename | Select-String -Pattern $RegExFindBracket | %{$_.Matches.Value}
      IF(!$TempDataStore) {$TempDataStore = "Error determining datastore"}
      $VMDiskInfoTemp = @()
      $VMDiskInfoTemp = New-Object psobject
      $VMDiskInfoTemp | Add-Member -MemberType NoteProperty -Name "VCenter" -Value $VCenterServer     
      $VMDiskInfoTemp | Add-Member -MemberType NoteProperty -Name "Host" -Value $VM.VMHost     
      $VMDiskInfoTemp | Add-Member -MemberType NoteProperty -Name "VM" -Value $VM.Name     
      $VMDiskInfoTemp | Add-Member -MemberType NoteProperty -Name "Name" -Value $VMDisk.Name
      $VMDiskInfoTemp | Add-Member -MemberType NoteProperty -Name "Description" -Value $VMDisk.Description
      $VMDiskInfoTemp | Add-Member -MemberType NoteProperty -Name "StorageFormat" -Value $VMDisk.StorageFormat
      $VMDiskInfoTemp | Add-Member -MemberType NoteProperty -Name "CapacityGB" -Value $VMDisk.CapacityGB
      $VMDiskInfoTemp | Add-Member -MemberType NoteProperty -Name "FileName" -Value $VMDisk.FileName
      $VMDiskInfoTemp | Add-Member -MemberType NoteProperty -Name "Parent" -Value $VMDisk.Parent
      $VMDiskInfoTemp | Add-Member -MemberType NoteProperty -Name "Datastore" -Value $TempDataStore
      $VMDiskInfo += $VMDiskInfoTemp
    }
  }
}
$AllVMInfo | Sort-Object VCenterServer, Host, Name | Export-Excel -Path $OutputFile -WorkSheetname "AllVMs" -TableName "ALlVMs" -TableStyle Medium4 -AutoSize
$VMDiskInfo | Sort-Object VCenter, VMHost, VM, Name | Export-Excel -Path $OutputFile -WorkSheetname "VM Disk Info" -TableName "VMDiskInfo" -TableStyle Medium4 -AutoSize
$VMDiskInfo | Sort-Object VCenter, Host, VM | Export-Excel -Path $OutputFile -WorkSheetname "VM Disk Info PivotTable" -TableName "PT_VMDiskInfo" -HideSheet "VM Disk Info PivotTable" -TableStyle Medium4 -AutoSize -IncludePivotTable -PivotRows VCenter, Host, VM  -PivotData @{CapacityGB='sum'} -IncludePivotChart -ChartType PieExploded3D
$VMDiskInfo | Sort-Object DataStore, VCenter, VM | Export-Excel -Path $OutputFile -WorkSheetname "DataStore Pivot Data" -TableName "PT_DataStoreInfo" -HideSheet "DataStore Pivot Data" -TableStyle Medium4 -AutoSize -IncludePivotTable -PivotRows DataStore, VCenter, Host  -PivotData @{CapacityGB='sum'} -IncludePivotChart -ChartType PieExploded3D
$CompleteVMInfo | Sort-Object DataStoreCluster, DataStores, VCenter | Export-Excel -Path $OutputFile -WorkSheetname "DataCluster Info" -TableName "PT_DataClusterInfo" -HideSheet "DataCluster Info" -TableStyle Medium4 -AutoSize -IncludePivotTable -PivotRows DataStoreCluster, Datastores, VCenterServer -PivotData @{UsedSpaceGB='sum'} -IncludePivotChart -ChartType PieExploded3D
Disconnect-VIServer -Confirm:$False