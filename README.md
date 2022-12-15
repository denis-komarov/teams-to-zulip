# TeamsToZulip

**TeamsToZulip** is a [PowerShell](https://microsoft.com/powershell) [module](https://technet.microsoft.com/en-us/library/dd901839.aspx)
that that allows you to transfer messages from Microsoft Teams chats to the corresponding Zulip topics.

**TeamsToZulip** has the following main functions:

- Get-TeamsChat
- Get-TeamsChatMember
- ConvertFrom-TeamsChatToZulipTopic

## Installation

### Create the necessary directories, for example:
```
d:\t2z
d:\t2z\download
d:\t2z\module
d:\t2z\module\TeamsToZulip
d:\t2z\trid
```

### Install the required modules from PSGallery
```powershell
$PSModuleDir = 'd:\t2z\module'
$ModuleRepository = 'PSGallery'
Save-Module -Name 'Microsoft.Graph.Users' -Path $PSModuleDir -Repository $ModuleRepository -Force
Save-Module -Name 'Microsoft.Graph.Teams' -Path $PSModuleDir -Repository $ModuleRepository -Force
Save-Module -Name 'Microsoft.Graph.Files' -Path $PSModuleDir -Repository $ModuleRepository -Force
Save-Module -Name 'MarkdownPrince'        -Path $PSModuleDir -Repository $ModuleRepository -Force
```

### Copy the program files (trid.exe and triddefs.trd) [TrID](http://mark0.net/soft-trid-e.html) into the:
```
d:\t2z\trid
```

### Copy from this repository files (**TeamsToZulip.psd1** and **TeamsToZulip.psm1**) into the:
```
d:\t2z\module\TeamsToZulip
```

## Usage

### Ask your system administrator for your Microsoft Teams tenant ID. This is an identifier that looks something like this:
```
b33cbe9f-8ebe-4f2a-912b-7e2a427f477f
```

### Get a list of your chats from Microsoft Teams that you participated in by running a script like this:
```powershell
$TenantId = 'b33cbe9f-8ebe-4f2a-912b-7e2a427f477f'

$env:PSModulePath += ';' + 'd:\t2z\module'

Import-Module -Name 'd:\t2z\module\TeamsToZulip'

Get-TeamsChat -TenantId $TenantId -TeamsChatType 'meeting'  #  -TeamsChatTopic 'WS-' 
| ForEach-Object { 
  '-' * 100
  $_.Id
  $_.ChatType
  $_.Topic
  $_.LastUpdatedDateTime.DateTime
  if ( $_.ChatType -in ( 'group', 'oneOnOne' ) )  #  ( 'meeting', 'group', 'oneOnOne' )
  { 
    'Members:' 
    Get-TeamsChatMember -TeamsAzureADTenantId $TenantId -TeamsChatId $_.Id 
    | ForEach-Object { 
      $_.DisplayName 
    } 
  }
}

```
