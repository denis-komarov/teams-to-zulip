# TeamsToZulip

**TeamsToZulip** is a [PowerShell](https://microsoft.com/powershell) [module](https://technet.microsoft.com/en-us/library/dd901839.aspx)
that that allows you to transfer messages from Microsoft Teams chats to the corresponding Zulip topics.

**TeamsToZulip** has the following main functions:

- Get-TeamsChat
- Get-TeamsChatMember
- ConvertFrom-TeamsChatToZulipTopic

## Installation

### Install the latest stable version of [PowerShell](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell) (from version 7.3 and above)

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
$env:PSModulePath += ';' + 'd:\t2z\module'

Import-Module -Name 'd:\t2z\module\TeamsToZulip'

Get-TeamsChat -TenantId 'b33cbe9f-8ebe-4f2a-912b-7e2a427f477f' -TeamsChatType 'meeting'  #  'meeting', 'group', 'oneOnOne'
| ForEach-Object { 
  '-' * 100
  $_.Id
  $_.ChatType
  $_.Topic
  $_.LastUpdatedDateTime.DateTime
  if ( $_.ChatType -in ( 'group', 'oneOnOne' ) )  #  ( 'meeting', 'group', 'oneOnOne' )
  { 
    Get-TeamsChatMember -TenantId 'b33cbe9f-8ebe-4f2a-912b-7e2a427f477f' -TeamsChatId $_.Id 
    'Members:' 
    | ForEach-Object { 
      $_.DisplayName 
    } 
  }
}

```

### Select from the list the chat that you want to transfer to Zulip. You need its id, it looks like this:
```
19:b8577894a63548969c5c92bb9c80c5e1@thread.v2
```

### To access Zulip you need to create your [API key](https://zulip.com/api/api-keys#get-your-api-key), it looks like this:
```
gjA04ZYcqXKalvYMA8OeXSfzUOLrtbZv
```


### In Zulip, you need to create or define a target topic to which you are going to transfer messages from Teams

### Well, now you can start transferring messages from Teams to Zulip by running a script like this:
```powershell
$env:PSModulePath += ';' + 'd:\t2z\module'

Import-Module -Name 'd:\t2z\module\TeamsToZulip'

$ConvertArgs = @{
  TeamsTenantId    = 'b33cbe9f-8ebe-4f2a-912b-7e2a427f477f'
  TeamsChatId      = '19:b8577894a63548969c5c92bb9c80c5e1@thread.v2'
  ZulipSite        = 'https://zulip.domain.com'
  ZulipEmail       = 'fn.ln@domain.com'
  ZulipKey         = 'gjA04ZYcqXKalvYMA8OeXSfzUOLrtbZv'
  ZulipStreamName  = 'Denmark'
  ZulipTopicName   = 'Castle'
  DownloadDir      = "d:\t2z\download"
  TrIDPathDir      = "d:\t2z\trid"
  ShowProgress     = $true
}  

ConvertFrom-TeamsChatToZulipTopic @ConvertArgs
```
