Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                                      = New-Object system.Windows.Forms.Form
$Form.ClientSize                           = '600,350'
$Form.text                                 = "AAD Pending Computers Fix"
$Form.TopMost                              = $true
#----------------------
#
$Status                                    = New-Object 'system.Windows.Forms.Label'
$Status.text                               = "Waiting for computer name"
$Status.AutoSize                           = $true
$Status.location                           = New-Object System.Drawing.Point(40,50)
$Status.font                               = 'Microsoft Sans Serif,10'
#
$ComputerName                              = New-Object 'system.Windows.Forms.Label'
$ComputerName.text                         = "Computer Name"
$ComputerName.AutoSize                     = $true
$ComputerName.location                     = New-Object System.Drawing.Point(20,100)
$ComputerName.font                         = 'Microsoft Sans Serif,10'
#
$ComputerNameTextBox                       = New-Object 'system.Windows.Forms.TextBox'
$ComputerNameTextBox                       = New-Object system.Windows.Forms.TextBox
$ComputerNameTextBox.multiline             = $false
$ComputerNameTextBox.Size                  = New-Object System.Drawing.Point(200,350)
$ComputerNameTextBox.location              = New-Object System.Drawing.Point(130,100)
$ComputerNameTextBox.Font                  = 'Microsoft Sans Serif,10'
$ComputerNameTextBox.Enabled               = $true
#
$ADSyncServer                              = New-Object 'system.Windows.Forms.Label'
$ADSyncServer.text                         = "AD Sync Server"
$ADSyncServer.AutoSize                     = $true
$ADSyncServer.location                     = New-Object System.Drawing.Point(20,150)
$ADSyncServer.font                         = 'Microsoft Sans Serif,10'
#
$ADSyncServerTextBox                       = New-Object 'system.Windows.Forms.TextBox'
$ADSyncServerTextBox                       = New-Object system.Windows.Forms.TextBox
$ADSyncServerTextBox.multiline             = $false
$ADSyncServerTextBox.Size                  = New-Object System.Drawing.Point(200,350)
$ADSyncServerTextBox.location              = New-Object System.Drawing.Point(130,150)
$ADSyncServerTextBox.Font                  = 'Microsoft Sans Serif,10'
$ADSyncServerTextBox.Enabled               = $true
#
$Timer                                     = New-Object 'system.Windows.Forms.Label'
$Timer.AutoSize                            = $true
$Timer.location                            = New-Object System.Drawing.Point(20,200)
$Timer.font                                = 'Microsoft Sans Serif,10'
#----------
$Apply                                     = New-Object system.Windows.Forms.Button
$Apply.text                                = "Fix"
$Apply.width                               = 99
$Apply.height                              = 30
$Apply.location                            = New-Object System.Drawing.Point(70,280)
$apply.Add_Click({check})
#----------
$Cancel                                   = New-Object system.Windows.Forms.Button
$Cancel.text                              = "Close"
$Cancel.width                             = 98
$Cancel.height                            = 30
$Cancel.location                          = New-Object System.Drawing.Point(200,280)
$Cancel.Add_Click({$Form.Close()})

$Form.Controls.AddRange(@($Status,$ComputerNameTextBox, $ComputerName,$ADSyncServer,$ADSyncServerTextBox, $Timer, $apply, $cancel))

function rejoincheck ()
{
   sleep 5
   $CersChecker = Invoke-Command -ComputerName $ComputerNameTextBox.text -ScriptBlock {Get-ChildItem Cert:\LocalMachine\My\ | Select-Object -ExpandProperty Issuer } |  Select-String -Pattern "MS-Organization" | Measure-Object | Select-Object -ExpandProperty Count
   If ($CersChecker -eq 2)
   {
        $Status.text              = "All has been set"
        $Status.foreColor         = "green"
        $timer.Visible            = $false
        $Cancel.Visible           = $true
    }
    else
    {
        $count = $count +1
        rejoincheck
        if ($count -gt 5)
        {
            $Status.text              = "There is an issue with "+$ComputerNameTextBox.text+". The certicates could not be installed."
            $Status.foreColor         = "Red"
            $Cancel.Visible           = $true
        }
    }
    return
}


function register ()

{
    $Status.text              = "Removing computer "+$ComputerNameTextBox.text+" from AAD"
    $Status.foreColor         = "DarkOrange"
    Invoke-Command -ComputerName $ComputerNameTextBox.text -ScriptBlock { dsregcmd /leave}
    $i = 60
    do {
        $Timer.text = "Seconds remaining : $($i)"
        Sleep 1
        $i--
    } while ($i -gt 0)

    Invoke-Command -ComputerName wmradsync01 -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta}
    $Status.text              = "Waiting for sync to be completed"
    $Status.foreColor         = "Purple"
    $i = 60
    do {
        $Timer.text = "Seconds remaining : $($i)"
        Sleep 1
        $i--
    } while ($i -gt 0)
    $Status.text              = "Please wait, re-joining computer to AAD"
    Invoke-Command -ComputerName $ComputerNameTextBox.text -ScriptBlock { schtasks.exe  /run /tn "\Microsoft\Windows\Workplace Join\Automatic-Device-Join"}
    rejoincheck
}

function check ()
{

    $Cancel.Visible           = $false
    $Status.text              = "Trying to communicate with "+$ComputerNameTextBox.text+" please wait"
    $Status.foreColor         = "orange"
    if (Test-Connection $ComputerNameTextBox.text -Quiet)
    {
        register
    }
else

{
    $Status.text              = $ComputerNameTextBox.text+" is not avialable. Is it connected to your LAN?"
    $Status.foreColor         = "Red"
    $Cancel.Visible           = $true
}
}

[void] $Form.ShowDialog()
