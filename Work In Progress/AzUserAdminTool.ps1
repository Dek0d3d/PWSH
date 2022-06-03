####################################################################
# This Script Renders a GUI which will assist in running commands. #
####################################################################

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$gui = New-Object System.Windows.Forms.Form
$flowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$flowPanel.AutoSize = $true
$gui.Text = 'Azure AD User Admin Tool'
$gui.Width = 600
$gui.Height = 400

$mgInstalled = if(!(Get-InstalledModule | Where-Object Name -like "*Microsoft.Graph*")){$false}

$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Microsoft Graph Powershell SDK installed:" + $mgInstalled
$Label.AutoSize = $true
$flowPanel.Controls.Add($Label)

$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Size = New-Object System.Drawing.Size(120,23)
$connectButton.Text = "Connect"

if(!$mgInstalled){
    $installButton = New-Object System.Windows.Forms.Button
    $installButton.Size = New-Object System.Drawing.Size(120,23)
    $installButton.Text = "Connect"
    $flowPanel.Controls.Add($installButton)
}
else {
    
    $flowPanel.Controls.Add($connectButton)
}

$gui.Controls.Add($flowPanel)
$gui.ShowDialog()

$installButton.Add_Click(
    {
        Install-Module Microsoft.Graph -Scope CurrentUser
        $flowPanel.Controls.Remove($installButton)
        $flowPanel.Controls.Add($connectButton)
    }
)

$connectButton.Add_Click(
    {
        Import-Module Microsoft.Graph
        Connect-MgGraph
    }
)

function Show-ADUsersOnForm{
    param (
            [Parameter (Mandatory = $true)] [System.Windows.Forms.Form] $FormName
    )
    
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "AD users"
    $Label.Location  = New-Object System.Drawing.Point(0,10)
    $Label.AutoSize = $true
    $FormName.Controls.Add($Label)

    $ComboBox = New-Object System.Windows.Forms.ComboBox
    $ComboBox.Width = 300
    $Users = get-aduser -filter * -Properties SamAccountName
    Foreach ($User in $Users){
        $ComboBox.Items.Add($User.SamAccountName);
    }
    $ComboBox.Location  = New-Object System.Drawing.Point(60,10)
    $FormName.Controls.Add($ComboBox)
}