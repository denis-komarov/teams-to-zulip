# TeamsToZulip

**TeamsToZulip** is a [PowerShell](https://learn.microsoft.com/powershell/) [module](https://learn.microsoft.com/en-us/previous-versions/dd901839(v=vs.85))
that that allows you to transfer messages from [Microsoft Teams](https://www.microsoft.com/en-us/microsoft-teams/group-chat-software) chats to the corresponding [Zulip](https://zulip.com) topics.

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

### Install the required modules from PSGallery by running a script like this in the PowerShell console:
```powershell
$PSModuleDir = 'd:\t2z\module'
$ModuleRepository = 'PSGallery'
Save-Module -Name 'Microsoft.Graph.Users' -Path $PSModuleDir -Repository $ModuleRepository -Force
Save-Module -Name 'Microsoft.Graph.Teams' -Path $PSModuleDir -Repository $ModuleRepository -Force
Save-Module -Name 'Microsoft.Graph.Files' -Path $PSModuleDir -Repository $ModuleRepository -Force
Save-Module -Name 'MarkdownPrince'        -Path $PSModuleDir -Repository $ModuleRepository -Force
```

### Copy the program files of [TrID](http://mark0.net/soft-trid-e.html) into the directory as follows:
Download the file [trid_w32.zip](https://mark0.net/download/trid_w32.zip) and unzip its contents into the directory so that the files appear in it:
```
d:\t2z\trid\trid.exe
d:\t2z\trid\readme.txt
```
Download the file [triddefs.zip](https://mark0.net/download/triddefs.zip) and unzip its contents into the directory so that the file appear in it:
```
d:\t2z\trid\triddefs.trd
```

### Copy the files from this repository into the directory, so that the files appear in it:
```
D:\t2z\module\TeamsToZulip\TeamsToZulip.psd1
D:\t2z\module\TeamsToZulip\TeamsToZulip.psm1
```

## Usage

### Ask your system administrator for your Microsoft Teams tenant ID. This is an identifier that looks something like this:
```
b33cbe9f-8ebe-4f2a-912b-7e2a427f477f
```

### Get a list of your chats from Microsoft Teams that you participated in by running a script like this in the PowerShell console:
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


### In Zulip, you need to create or define a target [stream and select topic](https://zulip.com/help/streams-and-topics) name to which you are going to transfer messages from Teams chat

### Well, now you can start transferring messages from Teams to Zulip by running a script like this in the PowerShell console:
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

## 3rd party references

This module uses but does not include various external libraries and programs. Their authors have done a fantastic job.
- [Powershell SDK for Microsoft Graph](https://github.com/microsoftgraph/msgraph-sdk-powershell) - Copyright (c) Microsoft Corporation. All Rights Reserved. Licensed under the MIT license.
- [MarkdownPrince](https://github.com/EvotecIT/MarkdownPrince) - Copyright (c) 2011 - 2021 Przemyslaw Klys @ Evotec. All rights reserved.
- [TrID](http://mark0.net/soft-trid-e.html) -  Copyright (c) 2003-16 By Marco Pontello
