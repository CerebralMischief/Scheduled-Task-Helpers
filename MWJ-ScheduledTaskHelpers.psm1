function New-TaskFolder 
{
  <#
    .SYNOPSIS

    Creates a user task folder.

    Author: Matthew Johnson (@mwjcomputing)
    License: GPLv3
    Required Dependencies: None
    Optional Dependencies: None

    NOTE: This script is a modified version of what was posted by Ed Wilson (@scriptingguys) at http://blogs.technet.com/b/heyscriptingguy/archive/2015/01/15/use-powershell-to-create-scheduled-tasks-folders.aspx.

    .DESCRIPTION

    New-TaskFolder creates a user created task folder that is specified by name.

    .PARAMETER Name

    Specifies the name of the folder to remove.

    .EXAMPLE

    New-TaskFolder -Name 'PowerShell Tasks'

    .EXAMPLE

    'PowerShell Tasks' | New-TaskFolder

    .LINK

    http://www.mwjcomputing.com

    .LINK

    https://github.com/mwjcomputing

    .LINK
    
    Remove-TaskFolder
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,
               HelpMessage = 'Enter name of folder to create.',
               ValueFromPipeline = $true)]
    [String] $Name
  )
  begin{
    Write-Verbose -Message '----- Starting function New-TaskFolder -----'

    Write-Verbose -Message '[x] Creating Schedule.Service COM Object'
    $ScheduleObject = New-Object -ComObject 'schedule.service'

    Write-Verbose -Message '[x] Connecting to the Schedule.Service COM Object'
    $ScheduleObject.Connect()

    Write-Verbose -Message '[x] Setting Root folder to Task Root'
    $RootFolder = $ScheduleObject.GetFolder('\')
  }
  process{
    Write-Verbose -Message "[x] Creating folder: $Name"
    $RootFolder.CreateFolder($Name)    
  }
  end{
    Write-Verbose -Message '[-] Removing variable RootFolder'
    Remove-Variable -Name 'RootFolder'

    Write-Verbose -Message '[-] Removing variable ScheduleObject'
    Remove-Variable -Name 'ScheduleObject'
        
    Write-Verbose -Message '----- Ending function New-TaskFolder -----'
  }
}

function Remove-TaskFolder
{
  <#
    .SYNOPSIS

    Removes a user created task folder.

    Author: Matthew Johnson (@mwjcomputing)
    License: GPLv3
    Required Dependencies: None
    Optional Dependencies: None

    NOTE: This script is a modified version of what was posted by Ed Wilson (@scriptingguys) at http://blogs.technet.com/b/heyscriptingguy/archive/2015/01/15/use-powershell-to-create-scheduled-tasks-folders.aspx.

    .DESCRIPTION

    Remove-TaskFolder removes a user created task folder that is specified by name.

    .PARAMETER Name

    Specifies the name of the folder to remove.

    .INPUTS

    System.String. 

    .OUTPUTS

    None

    .EXAMPLE

    Remove-TaskFolder -Name 'Server Tasks'

    .EXAMPLE

    'Server Tasks', 'PowerShell Tasks' | Remove-TaskFolder

    .LINK

    http://www.mwjcomputing.com

    .LINK

    https://github.com/mwjcomputing

    .LINK

    New-TaskFolder
  #>

  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,
               HelpMessage = 'Enter name of folder to remove.',
               ValueFromPipeline = $true)]
    [String] $Name
  )

  begin{
    Write-Verbose -Message '----- Starting function Remove-TaskFolder -----'

    Write-Verbose -Message '[x] Creating Schedule.Service COM Object'
    $ScheduleObject = New-Object -ComObject 'schedule.service'

    Write-Verbose -Message '[x] Connecting to the Schedule.Service COM Object'
    $ScheduleObject.Connect()

    Write-Verbose -Message '[x] Setting Root folder to Task Root'
    $RootFolder = $ScheduleObject.GetFolder('\')
  }

  process {
    Write-Verbose -Message "[-] Removing folder $Name"
    $RootFolder.DeleteFolder($Name,$null)
  }

  end {
    Write-Verbose -Message '[-] Removing variable RootFolder'
    Remove-Variable -Name 'RootFolder'

    Write-Verbose -Message '[-] Removing variable ScheduleObject'
    Remove-Variable -Name 'ScheduleObject'

    Write-Verbose -Message '----- Ending function Remove-TaskFolder -----'
  }
}

New-Alias -Name 'NTF' -Value 'New-TaskFolder'
New-Alias -Name 'RTF' -Value 'Remove-TaskFolder'
