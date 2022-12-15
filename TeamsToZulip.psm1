
$Script:ProgressPreference = 'SilentlyContinue'

$Script:TeamsEnvironment = 'Global'

$Script:TeamsApiVersion  = 'Beta'

$Script:ZulipApiVersion  = 'v1'


$Script:TeamsDefaultDelegatedPermissions = @(
'User.Read'
'User.ReadBasic.All'
'Team.ReadBasic.All'
'Channel.ReadBasic.All'
'Chat.Read'
'ChatMessage.Read'
'Files.Read.All'
'Sites.Read.All'
)

if ( -not ( Get-Module 'MarkdownPrince' ) )
{
  Import-Module 'MarkdownPrince' -Force -ErrorAction 'Stop'
}

if ( -not ( Get-Module 'Microsoft.Graph.Authentication' ) )
{
  Import-Module 'Microsoft.Graph.Authentication' -Force -ErrorAction 'Stop'
}

#  change target API version
Select-MgProfile -Name $Script:TeamsApiVersion

if ( -not ( Get-Module 'Microsoft.Graph.Users' ) )
{
  Import-Module 'Microsoft.Graph.Users' -Force -ErrorAction 'Stop'
}
  
if ( -not ( Get-Module 'Microsoft.Graph.Teams' ) )
{
  Import-Module 'Microsoft.Graph.Teams' -Force -ErrorAction 'Stop'
}

$Script:MarkdownArgs = @{
  Content                  = ''
  UnknownTags              = 'Bypass'
  GithubFlavored           = $true
  RemoveComments           = $true
  SmartHrefHandling        = $true
  DefaultCodeBlockLanguage = ''
  Format                   = $false
}


Function Connect-Teams
{
  [ CmdletBinding( PositionalBinding = $false ) ]
  Param
  (
    [ Parameter(Mandatory) ] [ string ] $TenantId,
                             [ string ] $TeamsEnvironment = $Script:TeamsEnvironment,
                           [ string[] ] $TeamsDelegatedPermissions = $Script:TeamsDefaultDelegatedPermissions
  )

  if ( 
        ( -not $Script:TeamsCurrentSession ) -or 
        (      $Script:TeamsCurrentSession -and 
               ( 
                 ( $Script:TeamsCurrentSession.TenantId -ne $TenantId ) -or 
                 ( $Script:TeamsCurrentEnvironment      -ne $TeamsEnvironment     ) -or
                 ( Compare-Object -ReferenceObject $Script:TeamsCurrentSession.Scopes -DifferenceObject $TeamsDelegatedPermissions )
               ) 
        )
  )
  {
    
    #  connect using interactive authentication
    Connect-MgGraph -Environment $TeamsEnvironment -TenantId $TenantId -Scopes $TeamsDelegatedPermissions -ErrorAction 'Stop'
  
    $Script:TeamsCurrentSession     = Get-MgContext -ErrorAction 'Stop'
    
    $Script:TeamsCurrentEnvironment = $TeamsEnvironment
    
    #  get ms graph endpoint for environment
    $Script:TeamsEndPoint           = ( Get-MgEnvironment | Where-Object { $_.Name -eq $TeamsEnvironment } ).GraphEndpoint + '/' + $Script:TeamsApiVersion
    
  }  
  
}


Function Get-TeamsChat
{
  [ CmdletBinding( PositionalBinding = $false ) ]
  Param
  (
    [ Parameter(Mandatory) ] [ string ] $TenantId,
                             [ string ] $TeamsEnvironment = $Script:TeamsEnvironment,
                           [ string[] ] $TeamsDelegatedPermissions = $Script:TeamsDefaultDelegatedPermissions,
                           [ string[] ] $TeamsChatType,  #  ( 'meeting', 'group', 'oneOnOne' )
                             [ string ] $TeamsChatTopic
  )
  
  Connect-Teams -TeamsEnvironment $TeamsEnvironment -TenantId $TenantId -TeamsDelegatedPermissions $TeamsDelegatedPermissions
  
  #  get chat list ordered by last activity time desc
  Get-MgChat -All -Sort 'lastMessagePreview/createdDateTime desc'
  | Where-Object { 
    ( ( -not ( $TeamsChatType  ) ) -or ( $TeamsChatType  -and ( $_.ChatType -in    $TeamsChatType  ) ) ) -and
    ( ( -not ( $TeamsChatTopic ) ) -or ( $TeamsChatTopic -and ( $_.Topic    -match $TeamsChatTopic ) ) )
  }

}


Function Get-TeamsChatMember
{
  [ CmdletBinding( PositionalBinding = $false ) ]
  Param
  (
    [ Parameter(Mandatory) ] [ string ] $TenantId,
                             [ string ] $TeamsEnvironment = $Script:TeamsEnvironment,    
                           [ string[] ] $TeamsDelegatedPermissions = $Script:TeamsDefaultDelegatedPermissions,
    [ Parameter(Mandatory) ] [ string ] $TeamsChatId
  )
  
  Connect-Teams -TeamsEnvironment $TeamsEnvironment -TenantId $TenantId -TeamsDelegatedPermissions $TeamsDelegatedPermissions
  
  #  get chat member list
  Get-MgChatMember -ChatId $TeamsChatId

}


Function Connect-Zulip
{
  [ CmdletBinding( PositionalBinding = $false ) ]
  Param
  (
    [ Parameter(Mandatory) ] [ string ] $ZulipSite,
    [ Parameter(Mandatory) ] [ string ] $ZulipEmail,
    [ Parameter(Mandatory) ] [ string ] $ZulipKey
  )

  $ZulipEndPoint = "$($ZulipSite)/api/$($Script:ZulipApiVersion)"
 
  $ZulipCredential = New-Object 'System.Management.Automation.PSCredential' -ArgumentList $ZulipEmail, ( $ZulipKey | ConvertTo-SecureString -AsPlainText -Force )

  $Script:ZulipGetStream = 
  @{
    Authentication = 'Basic'
    Credential     = $ZulipCredential
    ContentType    = 'application/x-www-form-urlencoded; charset=utf-8'
    Method         = 'Get'
    Uri            = ''
  }
  
  $Script:ZulipGetStreamUri = "$($ZulipEndPoint)/get_stream_id?stream="
  
  #  get all users
  $Script:ZulipGetAllUsers = 
  @{
    Authentication = 'Basic'
    Credential     = $ZulipCredential
    ContentType    = 'application/x-www-form-urlencoded; charset=utf-8'
    Method         = 'Get'
    Uri            = "$($ZulipEndPoint)/users"
  }
  
  #  file upload
  $Script:ZulipUploadFile = 
  @{
    Authentication     = 'Basic'
    Credential         = $ZulipCredential
    ContentType        = 'application/x-www-form-urlencoded; charset=utf-8'
    Method             = 'Post'
    Uri                = "$($ZulipEndPoint)/user_uploads"
    StatusCodeVariable = 'HttpStatus'
  }

  #  send message
  $Script:ZulipSendMessage = 
  @{
    Authentication = 'Basic'
    Credential     = $ZulipCredential
    ContentType    = 'application/x-www-form-urlencoded; charset=utf-8'
    Method         = 'Post'
    Uri            = "$($ZulipEndPoint)/messages"
  }
    
}

<#
 .Synopsis
  Migration of messages from Microsoft Teams to Zulip.

 .Description
  Migration of messages with all their contents from source Microsoft Teams chat to target Zulip stream/topic.

 .Parameter TeamsTenantId
  Microsoft Teams Tenant ID (for example : "b33cbe9f-8ebe-4f2a-912b-7e2a427f477f").

 .Parameter TeamsEnvironment
  Target Microsoft cloud name (for example : "Global").

 .Parameter TeamsDelegatedPermissions
  Delegated permissions for call Microsoft Graph REST API as a signed in user (for example : @('Team.ReadBasic.All','ChatMessage.Read') ).
  
 .Parameter TeamsChatId
  Message source Microsoft Teams internal chat id (for example : "19:b8577894a63548969c5c92bb9c80c5e1@thread.v2").
  
 .Parameter ZulipSite
  Target Zulip site (for example : "https://zulip.domain.com").

 .Parameter ZulipEmail
  The user email on whose behalf the migration is performed (for example : "fn.ln@domain.com").

 .Parameter ZulipKey
  The user API key on whose behalf the migration is performed (for example : "gjA04ZYcqXKalvYMA8OeXSfzUOLrtbZv").
  
 .Parameter ZulipStreamName
  Target Zulip stream name (for example : "Denmark").
  
 .Parameter ZulipTopicName
  Target Zulip topic name (for example : "Castle").
  
 .Parameter ZulipAccountList
  Zulip account list on whose behalf messages are created (for example : @{ 'f1.n1@domain.com' = 'gjA04ZYcqXKalvYMA8OeXSfzUOLrtbZx'; 'f2.n2@domain.com' = 'gjA04ZYcqXKalvYMA8OeXSfzUOLrtbZy' } ).
  
 .Parameter DownloadDir
  Directory path for storing files downloaded from Teams (for example : "d:\teams_to_zulip\download").
  
 .Parameter TrIDPathDir
  The catalog in which the program TrID (C) 2003-16 By Marco Pontello (https://mark0.net/soft-trid-e.html) is located (for example : "d:\teams_to_zulip\trid").

 .Parameter ShowProgress
  Displays a progress bar in a PowerShell command window.

 .Example
  $ConvertArgs = @{
    #  Teams connect
    TeamsTenantId        = 'b33cbe9f-8ebe-4f2a-912b-7e2a427f477f'
    #  Teams source chat
    TeamsChatId          = '19:b8577894a63548969c5c92bb9c80c5e1@thread.v2'
    #  Zulip connect
    ZulipSite            = 'https://zulip.domain.com'
    ZulipEmail           = 'fn.sn@domain.com'
    ZulipKey             = 'gjA04ZYcqXKalvYMA8OeXSfzUOLrtbZv'
    #  Zulip target steeam/topic
    ZulipStreamName      = 'Denmark'
    ZulipTopicName       = 'Castle'
    #  Paths
    DownloadDir          = 'd:\teams_to_zulip\download'
    TrIDPathDir          = 'd:\teams_to_zulip\trid'
    #  Progress
    ShowProgress         = $true
  }  
  ConvertFrom-TeamsChatToZulipTopic @ConvertArgs

#>
function ConvertFrom-TeamsChatToZulipTopic
{
  [ CmdletBinding( PositionalBinding = $false ) ]
  Param
  (
    [ Parameter(Mandatory) ] [ string ] $TeamsTenantId,
                             [ string ] $TeamsEnvironment = $Script:TeamsEnvironment,    
                           [ string[] ] $TeamsDelegatedPermissions = $Script:TeamsDefaultDelegatedPermissions,
    [ Parameter(Mandatory) ] [ string ] $TeamsChatId,
    [ Parameter(Mandatory) ] [ string ] $ZulipSite,
    [ Parameter(Mandatory) ] [ string ] $ZulipEmail,
    [ Parameter(Mandatory) ] [ string ] $ZulipKey,
    [ Parameter(Mandatory) ] [ string ] $ZulipStreamName,
    [ Parameter(Mandatory) ] [ string ] $ZulipTopicName,
                          [ hashtable ] $ZulipAccountList,
    [ Parameter(Mandatory) ] [ string ] $DownloadDir,
    [ Parameter(Mandatory) ] [ string ] $TrIDPathDir,
                             [ switch ] $ShowProgress
  )
  
  Write-Verbose -Message "Starting: `n$($MyInvocation.MyCommand)"
  
  Connect-Teams -TeamsEnvironment $TeamsEnvironment -TenantId $TeamsTenantId -TeamsDelegatedPermissions $TeamsDelegatedPermissions
  
  Connect-Zulip -ZulipSite $ZulipSite -ZulipEmail $ZulipEmail -ZulipKey $ZulipKey
  
  try
  {
    
    Write-Verbose -Message "Get Zulip stream id for : $($ZulipStreamName)"
    
    #  get stream id
    $Script:ZulipGetStream.Uri = $Script:ZulipGetStreamUri + $ZulipStreamName
    
    $ZulipStreamId = ( Invoke-RestMethod @Script:ZulipGetStream ).stream_id
    
    Write-Verbose -Message "Zulip stream id : $($ZulipStreamId)"
    
  }
  catch
  {
    if ( $_.ErrorDetails.Message ) { $ErrorMessage = $_.ErrorDetails.Message } else { $ErrorMessage = $_ }
    Write-Error "Zulip stream $($ZulipStreamName) not found! `n$($ErrorMessage)"
    return
  }    
    
  
  #  html document object
  $HtmlDocument = New-Object -Com 'HTMLFile'
  
  try
  {
    
    Write-Verbose -Message "Get Teams user list"
    
    #  get teams user list
    $TeamsUserList = @{}
    Get-MgUser -All 
    | ForEach-Object { 
      $TeamsUserList[ $_.Id ] = @{ 
        DisplayName = $_.DisplayName 
        Mail        = $_.Mail
      } 
    }
  }
  catch
  {
    throw "Teams user list not found!"
  }  
  
  try
  {
    
    Write-Verbose -Message "Get Zulip user list"
    
    #  get zulip user list
    $ResponseZulipGetAllUsers = Invoke-RestMethod @Script:ZulipGetAllUsers

    $ZulipUserListByEmail = @{}
    $ResponseZulipGetAllUsers.members 
    | ForEach-Object { 
      $ZulipUserListByEmail[ $_.email ] = @{ 
        email     = $_.email 
        user_id   = $_.user_id
        full_name = $_.full_name
      } 
    }
  }
  catch
  {
    throw "Zulip user list not found!"
  }  
  
  try
  {
    
    Write-Verbose -Message "Get Teams chat message list"
    
    $ExportMessageList = @{}

    #  get chat message list
    $ChatMessageList = Get-MgChatMessage -ChatId $TeamsChatId -All  
    | Where-Object { ( $_.MessageType -in ( 'message' ) ) -and ( -not ( $_.DeletedDateTime ) ) }
    | Sort-Object -Property Id
  }
  catch
  {
    throw "Teams chat message list not found!"
  }  
  
  #  process message list  
  foreach ( $ChatMessage in $ChatMessageList )
  {
    
    if ( $ZulipAccountList )
    {
      
      $ChatMessageFromUserMail = ( $TeamsUserList[ $ChatMessage.From.User.Id ] ).Mail
      
      if ( $ZulipAccountList.Contains( $ChatMessageFromUserMail ) )
      {
        #  connection on behalf of user from account list
        Connect-Zulip -ZulipSite $ZulipSite -ZulipEmail $ChatMessageFromUserMail -ZulipKey $ZulipAccountList[ $ChatMessageFromUserMail ]
        
        Write-Verbose -Message "Connection on behalf of user account : $($ChatMessageFromUserMail)"
        
      }
      else
      {
        #  connection on behalf of main user
        Connect-Zulip -ZulipSite $ZulipSite -ZulipEmail $ZulipEmail -ZulipKey $ZulipKey
        
        Write-Verbose -Message "Connection on behalf of main account : $($ZulipEmail)"
        
      }  
      
    }  
    
    $ExportMessageList[ $ChatMessage.Id ] = @{ 
      Id                  = $ChatMessage.Id
      CreatedDateTime     = $ChatMessage.CreatedDateTime
      FromUserId          = $ChatMessage.From.User.Id
      QuoteMark           = ''
    }
    
    Write-Verbose -Message "Teams message id: $($ChatMessage.Id)"
    
    Write-Verbose -Message "Teams HTML initial message body: `n`n$($ChatMessage.Body.Content)`n"
    
    Write-Verbose -Message "Parse Teams HTML message body"
    
    #  parse message body html 
    $HtmlDocument.write( [System.Text.Encoding]::Unicode.GetBytes( $ChatMessage.Body.Content ) )
    $HtmlDocument.Close()
    
    #  get image links
    foreach ( $Image in $HtmlDocument.Images )
    { 
    
      $ImageId  = $Image.GetAttribute( 'itemid' )
      
      $ImageUrl = $Image.GetAttribute( 'src' )
      
      $ImageItemType  = $Image.GetAttribute( 'itemtype' )
      
      if ( $ImageItemType -eq 'http://schema.skype.com/Emoji' )
      {
        #  emoji, so no need to download the image file
        $Image.SetAttribute( 'src', '' )
      }  
      else
      {
        #  download image file
        try
        {
          
          Write-Verbose -Message "Download image file from Teams : $($DownloadDir)\$($ImageId)"
          
          #  download image to file
          Invoke-MgGraphRequest -Uri $ImageUrl -Method 'Get' -OutputFilePath "$($DownloadDir)\$($ImageId)" -ErrorAction 'SilentlyContinue'
          
          if ( Test-Path "$($DownloadDir)\$($ImageId)" -PathType 'Leaf' )
          {
          
            #  set real file extension for image file
            & "$($TrIDPathDir)\trid.exe" @( "$($DownloadDir)\$($ImageId)", '-ae' ) > $null
            
            #  get image file full name
            $ImageFilePath = ( Get-ChildItem -Path "$($DownloadDir)\$($ImageId).*" ).FullName
            
            do
            {
              
              try
              {
                
                Write-Verbose -Message "Upload image file to Zulip : $($ImageFilePath)"
                
                #  upload file to zulip
                $Script:ZulipUploadFile.Form = @{ f = Get-Item -Path $ImageFilePath }
                $ResponseZulipUploadFile = Invoke-RestMethod @Script:ZulipUploadFile
                
                #  change image url to zulip url
                $Image.SetAttribute( 'src', "$($ZulipSite)$($ResponseZulipUploadFile.uri)" )
              
              }
              catch
              {
                if ( $_.ErrorDetails.Message ) { $ErrorMessage = $_.ErrorDetails.Message | ConvertFrom-Json } else { $ErrorMessage = $_ }
                
                #  detect api throttling
                if ( $ErrorMessage.code -eq 'RATE_LIMIT_HIT' )
                {
                  Write-Verbose -Message "Zulip API usage exceeded rate limit. Wait for $($ErrorMessage.'retry-after') sec."
                  
                  Start-Sleep -Milliseconds ( $ErrorMessage.'retry-after' * 1000 )
                } 
                else
                {
                  Write-Verbose -Message "The file is not uploaded to Zulip : $($ImageFilePath) `n$($ErrorMessage)"
                  
                  $Image.SetAttribute( 'src', '' )
                  
                  break
                } 
                
              }  
              
            }
            until ( $ResponseZulipUploadFile.result -eq 'success' )
          
          }
          else
          {
            $Image.SetAttribute( 'src', '' )
          }  
          
        }
        catch
        {
          if ( $_.ErrorDetails.Message ) { $ErrorMessage = $_.ErrorDetails.Message } else { $ErrorMessage = $_ }
          
          Write-Verbose -Message "The file is not downloaded from Teams : $($ImageUrl) `n$($ErrorMessage)"
          
          $Image.SetAttribute( 'src', '' )
        }   

      }        
      
    }

    
    #  get final html text
    $HtmlText = $HtmlDocument.GetElementsByTagName( 'HTML' ) | Join-String -Property { $_.OuterHTML }
    
    Write-Verbose -Message "Teams HTML final message body: `n`n$($HtmlText)`n"
    
    Write-Verbose -Message "Convert Teams HTML message body to Markdown"
    
    #  convert html to markdown
    $MarkdownArgs.Content = $HtmlText
    $MarkdownText = ConvertFrom-HTMLToMarkdown @MarkdownArgs

    #  add author and date to message body
    $MarkdownText = 
    "@_**{0}|{1}** {2}`n{3}" -f (
      ( $ZulipUserListByEmail[ ( $TeamsUserList[ $ChatMessage.From.User.Id ] ).Mail ] ).full_name,
      ( $ZulipUserListByEmail[ ( $TeamsUserList[ $ChatMessage.From.User.Id ] ).Mail ] ).user_id,
      $ChatMessage.CreatedDateTime,
      $MarkdownText
    )
    
    
    #  get attchments
    foreach ( $Attachment in $ChatMessage.Attachments )
    {
      switch ( $Attachment.ContentType )
      {
        'reference'
        {
          
          $Base64Value = [System.Convert]::ToBase64String( [Text.Encoding]::UTF8.GetBytes( $Attachment.ContentUrl ), [Base64FormattingOptions]::None )

          $EncodedUrl = 'u!' + $Base64Value.TrimEnd( '=' ).Replace( '/', '_' ).Replace( '+', '-' )    
          
          try
          {
            
            #  get drive item
            $DriveItem = Invoke-MgGraphRequest -Uri "$($Script:TeamsEndPoint)/shares/$($EncodedUrl)/driveItem" -Method 'Get' 
            
            try
            {
              
              Write-Verbose -Message "Download attachment file from Teams : $($DownloadDir)\$($DriveItem.name)"
              
              #  download drive item to file
              Invoke-WebRequest -Uri $DriveItem.'@microsoft.graph.downloadUrl' -OutFile "$($DownloadDir)\$($DriveItem.name)"
              
              do
              {
                
                try
                {
                
                  Write-Verbose -Message "Upload attachment file to Zulip : $($DownloadDir)\$($DriveItem.name)"
                  
                  #  upload file to zulip
                  $Script:ZulipUploadFile.Form = @{ f = Get-Item -Path "$($DownloadDir)\$($DriveItem.name)" }
                  $ResponseZulipUploadFile = Invoke-RestMethod @Script:ZulipUploadFile
                  
                  #  add link to zulip url
                  $MarkdownText = "{0}`n[{1}]({2})" -f $MarkdownText, $Attachment.Name, "$($ZulipSite)$($ResponseZulipUploadFile.uri)"
                  
                }
                catch
                {
                  if ( $_.ErrorDetails.Message ) { $ErrorMessage = $_.ErrorDetails.Message | ConvertFrom-Json } else { $ErrorMessage = $_ }
                  
                  #  detect api throttling
                  if ( $ErrorMessage.code -eq 'RATE_LIMIT_HIT' )
                  {
                    Write-Verbose -Message "Zulip API usage exceeded rate limit. Wait for $($ErrorMessage.'retry-after') sec."
                    
                    Start-Sleep -Milliseconds ( $ErrorMessage.'retry-after' * 1000 )
                  } 
                  else
                  {
                    Write-Verbose -Message "The file is not uploaded to Zulip : $($Attachment.Name) `n$($ErrorMessage)"
                    
                    $MarkdownText = "{0}`n[{1}]({2})" -f $MarkdownText, $Attachment.Name, ''
                    
                    break
                  } 
                  
                }  
                
              }
              until ( $ResponseZulipUploadFile.result -eq 'success' )
              
            }
            catch
            {
              if ( $_.ErrorDetails.Message ) { $ErrorMessage = $_.ErrorDetails.Message } else { $ErrorMessage = $_ }
              Write-Verbose -Message "The file is not downloaded from Teams : $($Attachment.Name) `n$($ErrorMessage)"
              
              $MarkdownText = "{0}`n[{1}]({2})" -f $MarkdownText, $Attachment.Name, ''
            }
          
          }
          catch
          {
            if ( $_.ErrorDetails.Message ) { $ErrorMessage = $_.ErrorDetails.Message } else { $ErrorMessage = $_ }
            Write-Verbose -Message "The drive item is not obtained from Teams : $($Attachment.ContentUrl) `n$($ErrorMessage)"
            
            $MarkdownText = "{0}`n[{1}]({2})" -f $MarkdownText, $Attachment.Name, ''
          }  
          
        }
        'messageReference'
        {
          
          $ReferenceMessageId = ( ConvertFrom-Json -InputObject $Attachment.Content ).messageId
          
          $QuoteMark = ( $ExportMessageList[ $ReferenceMessageId ] ).QuoteMark
          
          if ( $QuoteMark ) { $QuoteMark += '`' }  else { $QuoteMark = '```' } 

          ( $ExportMessageList[ $ChatMessage.Id ] ).QuoteMark = $QuoteMark
          
          $ZulipUser = $ZulipUserListByEmail[ ( $TeamsUserList[ ( $ExportMessageList[ $ReferenceMessageId ] ).FromUserId ] ).Mail ]
          
          #  add quote
          $MarkdownText = 
@'
@_**{0}|{1}** [said]($($ZulipSite)/#narrow/stream/{2}-{3}/topic/{4}/near/{5}):
{6}quote
{7}
{8}
{9}
'@        -f (
                $ZulipUser.full_name,
                $ZulipUser.user_id,
                $ZulipStream.Id,
                $ZulipStreamName,
                $ZulipTopicName,
                ( $ExportMessageList[ $ReferenceMessageId ] ).ZulipMessageId, 
                $QuoteMark,
                ( $ExportMessageList[ $ReferenceMessageId ] ).MarkdownText,
                $QuoteMark,
                $MarkdownText
             )
          
        }        
      }
      
    }
    
    #  adapt markdown to zulip markdown implementation
    $MarkdownText = $MarkdownText.Replace( '![', '[' )
    $MarkdownText = $MarkdownText.Replace( '\_', '_' )
    
    Write-Verbose -Message "Post message to Zulip"
    
    #  post message
    $Script:ZulipSendMessage.Body = "type=stream&to=[$($ZulipStreamId)]&topic=$($ZulipTopicName)&content=$($MarkdownText)"
    
    do 
    {    
    
      try
      {
        
        $ResponseZulipSendMessage = Invoke-RestMethod @Script:ZulipSendMessage
        
        Write-Verbose -Message "Message id : $($ResponseZulipSendMessage.id)"
        
        ( $ExportMessageList[ $ChatMessage.Id ] ).Add( 'ZulipMessageId', $ResponseZulipSendMessage.id )
        ( $ExportMessageList[ $ChatMessage.Id ] ).Add( 'MarkdownText'  , $MarkdownText )
        
        Write-Verbose -Message "Zulip Markdown message body: `n`n$($MarkdownText)`n"
      
      }
      catch
      {
        
        if ( $_.ErrorDetails.Message ) { $ErrorMessage = $_.ErrorDetails.Message | ConvertFrom-Json } else { $ErrorMessage = $_ }
        
        #  detect api throttling
        if ( $ErrorMessage.code -eq 'RATE_LIMIT_HIT' )
        {
          Write-Verbose -Message "Zulip API usage exceeded rate limit. Wait for $($ErrorMessage.'retry-after') sec."
          
          Start-Sleep -Milliseconds ( $ErrorMessage.'retry-after' * 1000 )
        } 
        else
        {
          throw "Message could not be sent to Zulip `n$($ErrorMessage)"
        } 
        
      }  
    } 
    until ( $ResponseZulipSendMessage.id )
    
    #  display a progress bar
    if ( $ShowProgress )
    {
      $PercentComplete = [math]::Round( ( $ExportMessageList.Count / $ChatMessageList.Count ) * 100 )
      
      $ProgressPreferenceCurrent = $ProgressPreference
      
      $ProgressPreference = 'Continue'
      
      Write-Progress -Activity "Migration in Progress" -Status "$($PercentComplete)% Complete:" -PercentComplete $PercentComplete
      
      $ProgressPreference = $ProgressPreferenceCurrent
    }  
    
  }
  
}  


Export-ModuleMember -Function Connect-Teams

Export-ModuleMember -Function Get-TeamsChat

Export-ModuleMember -Function Get-TeamsChatMember

Export-ModuleMember -Function Connect-Zulip

Export-ModuleMember -Function ConvertFrom-TeamsChatToZulipTopic