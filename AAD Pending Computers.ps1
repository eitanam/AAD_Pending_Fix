Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                                      = New-Object system.Windows.Forms.Form
$Form.ClientSize                           = '500,300'
$Form.text                                 = "Tnuva AAD pending computers Fix"
$Form.TopMost                              = $true
#----------------------
#
$Status                                    = New-Object 'system.Windows.Forms.Label'
$Status.text                               = "Waiting for computer name"
$Status.Size                               = New-Object System.Drawing.Point(400,50)
$Status.location                           = New-Object System.Drawing.Point(60,50)
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
$Timer                                     = New-Object 'system.Windows.Forms.Label'
$Timer.AutoSize                            = $true
$Timer.location                            = New-Object System.Drawing.Point(20,150)
$Timer.font                                = 'Microsoft Sans Serif,10'
#----------
$Apply                                     = New-Object system.Windows.Forms.Button
$Apply.text                                = "Fix"
$Apply.width                               = 99
$Apply.height                              = 30
$Apply.location                            = New-Object System.Drawing.Point(30,220)
$apply.Add_Click({check})
#----------
$Cancel                                   = New-Object system.Windows.Forms.Button
$Cancel.text                              = "Close"
$Cancel.width                             = 98
$Cancel.height                            = 30
$Cancel.location                          = New-Object System.Drawing.Point(160,220)
$Cancel.Add_Click({$Form.Close()})

$Form.Controls.AddRange(@($Status,$ComputerNameTextBox, $ComputerName,$Timer, $apply, $cancel))


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
        $apply.Visible            = $false
   }
    else
    {
        $count = $count +1
        $Status.text              = "Trying "+$count+" of max 5 times"
        if ($count -gt 5)
        {
            $Status.text              = "There is an issue with "+$ComputerNameTextBox.text+", please re-run the script"
            $Status.foreColor         = "Red"
            $Cancel.Visible           = $true
        }
        sync
    }
 }


function sync()
{
    
   Invoke-Command -ComputerName wmradsync01 -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta}
    $Status.text              = "Waiting for sync to be completed"
    $Status.foreColor         = "Purple"
    $i = 60
    do {
        $Timer.text = "Seconds remaining : $($i)"
        $Form.Refresh()
        Sleep 1
        $i--
    } while ($i -gt 0)
    $Status.text              = "Please wait, re-joining computer to AAD"
    $Timer.Text               = " "
    Invoke-Command -ComputerName $ComputerNameTextBox.text -ScriptBlock { schtasks.exe  /run /tn "\Microsoft\Windows\Workplace Join\Automatic-Device-Join"}
    rejoincheck 

}

function register ()

{
    $Status.text              = "Removing computer "+$ComputerNameTextBox.text+" from AAD"
    $Status.foreColor         = "DarkOrange"
    Invoke-Command -ComputerName $ComputerNameTextBox.text -ScriptBlock { dsregcmd /leave}
    $i = 60
    do {
        $Timer.text = "Seconds remaining : $($i)"
        $Form.Refresh()
        Sleep 1
        $i--
    }while ($i -gt 0)
    sync
}

function check ()
{

    $Cancel.Visible           = $false
    $apply.Visible            = $false
    $Status.text              = "Trying to communicate with "+$ComputerNameTextBox.text+" please wait"
    $Status.foreColor         = "orange"
    if (Test-Connection $ComputerNameTextBox.text -Quiet)
    {
        register
    }
else

{
    $Status.text              = $ComputerNameTextBox.text+" is not avialable"
    $Status.foreColor         = "Red"
    $Cancel.Visible           = $true
}
}

[void] $Form.ShowDialog()
