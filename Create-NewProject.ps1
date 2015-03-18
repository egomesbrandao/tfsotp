Param(
    [parameter(Mandatory=$false)][string] $CollectionUrlParam = $(Read-Host -prompt "Collection URL"),
    [parameter(Mandatory=$false)][string] $TeamProjectParam = $(Read-Host -prompt "Team Project"),
	[parameter(Mandatory=$true)][String] $GlobalListName = $(Read-Host -prompt "Project List"),
    [parameter(Mandatory=$true)][String] $GlobalEntryValue = $(Read-Host -prompt "New project name")
    )

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

function Add-TfsGlobalListItem {
    Param(
        [parameter(Mandatory=$true)][Microsoft.TeamFoundation.Client.TfsTeamProjectCollection] $TfsCollection,
        [parameter(Mandatory=$true)][String] $GlobalListName,
        [parameter(Mandatory=$true)][String] $GlobalEntryValue
        )
    # Get Global List
    $store = Get-TfsWorkItemStore $TfsCollection
    [xml]$export = $store.ExportGlobalLists();

    $globalLists = $export.ChildNodes[0];
    $globalList = $globalLists.SelectSingleNode("//GLOBALLIST[@name='$GlobalListName']")

    # if no GL then add it
    If ($globalList -eq $null)
    {
        $globalList = $export.CreateElement("GLOBALLIST");
        $globalListNameAttribute = $export.CreateAttribute("name");
        $globalListNameAttribute.Value = $GlobalListName
        $globalList.Attributes.Append($globalListNameAttribute);
        $globalLists.AppendChild($globalList);
    }

    #Create a new node.
    $GlobalEntry = $export.CreateElement("LISTITEM");
    $GlobalEntryAttribute = $export.CreateAttribute("value");
    $GlobalEntryAttribute.Value = $GlobalEntryValue
    $GlobalEntry.Attributes.Append($GlobalEntryAttribute);

    #Add new entry to list
    $globalList.AppendChild($GlobalEntry)
    # Import list to server
    $store.ImportGlobalLists($globalLists)
}

$TfsCollection = Get-TfsCollection $CollectionUrlParam

$TfsCollection.EnsureAuthenticated()

$store = Get-TfsWorkItemStore $TfsCollection

#If the list does not exist creates or just add the name of the new project
[xml] $export = $store.ExportGlobalLists();

$globalLists = $export.ChildNodes[0];

$globalList = $globalLists.SelectSingleNode("//GLOBALLIST[@name='$GlobalListName']")

If ($globalList -eq $null)
{
    $globalList = $export.CreateElement("GLOBALLIST");
    $globalListNameAttribute = $export.CreateAttribute("name");
    $globalListNameAttribute.Value = $GlobalListName
    $globalList.Attributes.Append($globalListNameAttribute);
    $globalLists.AppendChild($globalList);
}

$GlobalEntry = $export.CreateElement("LISTITEM");
$GlobalEntryAttribute = $export.CreateAttribute("value");
$GlobalEntryAttribute.Value = $GlobalEntryValue
$GlobalEntry.Attributes.Append($GlobalEntryAttribute);

$globalList.AppendChild($GlobalEntry)

$store.ImportGlobalLists($globalLists)

#Creates Team with the new name of the Project
$cssService = $TfsCollection.GetService("Microsoft.TeamFoundation.Server.ICommonStructureService3")

$TeamProject += $cssService.GetProjectFromName($TeamProjectParam)

$teamService = $TfsCollection.GetService("Microsoft.TeamFoundation.Client.TfsTeamService")

$Team = $teamService.CreateTeam($teamProject.Uri, $GlobalEntryValue, "", $null)