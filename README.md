# TeamsToZulip

**TeamsToZulip** is a [PowerShell](https://microsoft.com/powershell) [module](https://technet.microsoft.com/en-us/library/dd901839.aspx)
that that allows you to transfer messages from Microsoft Teams chats to the corresponding Zulip topics.

**TeamsToZulip** has 2 main functions:

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
