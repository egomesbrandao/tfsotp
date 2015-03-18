$pathToAss2 = "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0"
$pathToAss4 = "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v4.5"
Add-Type -Path "$pathToAss2\Microsoft.TeamFoundation.Client.dll"
Add-Type -Path "$pathToAss2\Microsoft.TeamFoundation.Common.dll"
#Add-Type -Path "$pathToAss2\Microsoft.TeamFoundation.dll"
Add-Type -Path "$pathToAss2\Microsoft.TeamFoundation.WorkItemTracking.Client.dll"
Add-Type -Path "$pathToAss2\Microsoft.TeamFoundation.VersionControl.Client.dll"
Add-Type -Path "$pathToAss4\Microsoft.TeamFoundation.ProjectManagement.dll"

function Get-TfsCollection {
 Param(
       [string] $CollectionUrl
       )
    if ($CollectionUrl -ne "")
    {
        #if collection is passed then use it and select all projects
        $tfs = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($CollectionUrl)
    }
    else
    {
        #if no collection specified, open project picker to select it via gui
        $picker = New-Object Microsoft.TeamFoundation.Client.TeamProjectPicker([Microsoft.TeamFoundation.Client.TeamProjectPickerMode]::NoProject, $false)
        $dialogResult = $picker.ShowDialog()
        if ($dialogResult -ne "OK")
        {
            #exit
        }
        $tfs = $picker.SelectedTeamProjectCollection
    }
    Return $tfs
}

function Get-TfsCommonStructureService {
 Param(
       [Microsoft.TeamFoundation.Client.TfsTeamProjectCollection] $TfsCollection
       )
    Return $TfsCollection.GetService("Microsoft.TeamFoundation.Server.ICommonStructureService3")
}

$global:TfsWorkItemStoreCache
function Get-TfsWorkItemStore {
 Param(
       [Microsoft.TeamFoundation.Client.TfsTeamProjectCollection] $TfsCollection,
       [switch] $refresh
       )
       If ($global:TfsWorkItemStoreCache -eq $null -or $refresh -eq $true)
       {
       $global:TfsWorkItemStoreCache= $TfsCollection.GetService([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])
       }
    Return $global:TfsWorkItemStoreCache
}

function Get-TfsVersionControlServer {
    Param(
        [Microsoft.TeamFoundation.Client.TfsTeamProjectCollection] $TfsCollection
        )
    Return $TfsCollection.GetService("Microsoft.TeamFoundation.VersionControl.Client.VersionControlServer")
}

function Get-TfsProjectProcessConfigurationService {
    Param(
        [Microsoft.TeamFoundation.Client.TfsTeamProjectCollection] $TfsCollection
        )
    return $TfsCollection.GetService([Microsoft.TeamFoundation.ProcessConfiguration.Client.ProjectProcessConfigurationService]);
}

function Get-TfsTeamSettingsConfigurationService {
    Param(
        [Microsoft.TeamFoundation.Client.TfsTeamProjectCollection] $TfsCollection
        )
    return $TfsCollection.GetService([ Microsoft.TeamFoundation.ProcessConfiguration.Client.TeamSettingsConfigurationService]);
}