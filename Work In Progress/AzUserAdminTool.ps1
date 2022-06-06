####################################################################
# This Script Renders a GUI which will assist in running commands. #
####################################################################
$Script:tenant = ""
$Script:connected = $false

Function CheckModule{
    if(!(Get-InstalledModule | Where-Object Name -like "*Microsoft.Graph*")){
        return $false
    } else {
        return $true
    }
}

Function RenderGUI {
    $gui.Close()
    $gui.Dispose()
    GenerateGUI
}

Function InstallModule {
    Install-Module Microsoft.Graph -Scope CurrentUser
    RenderGUI
}

Function ConnectModule {
    Import-Module Microsoft.Graph
    Connect-MgGraph
    GetConnection
    RenderGUI
}

Function GetConnection {
    $Script:tenant = Get-MgOrganization | Select-Object -Property *
    $Script:connected = $true
}

Function DisplayTenantInfo {
    param([Parameter()][System.Windows.Forms.GroupBox]$gpBox)

    $TenantLabel = New-Object System.Windows.Forms.Label
    $TenantLabel.Text = "$(($Script:tenant).DisplayName) : $(($Script:tenant).Id)"
    $TenantLabel.Width = 275
    $TenantLabel.Height = 25
    
    $UserLabel = New-Object System.Windows.Forms.Label
    $UserLabel.Text = "$((Get-MgContext).Account)"
    $UserLabel.Width = 275
    $UserLabel.Height = 25
    
    $gpBox.Controls.Add($TenantLabel)
    $gpBox.Controls.Add($UserLabel)
}

Function DisplayConnection {
    param([Parameter()][System.Windows.Forms.FlowLayoutPanel]$flowPanel)

    $modLabel = New-Object System.Windows.Forms.Label
    $modLabel.Text = "Microsoft Graph Powershell SDK installed:" + $(CheckModule)
    $modlabel.Width = 300
    $modLabel.Height = 25

    if(CheckModule){
        $connectButton = New-Object System.Windows.Forms.Button
        $connectButton.Size = New-Object System.Drawing.Size(120,23)
        $connectButton.Text = "Connect"
        if($Script:tenant){
            $connectButton.Enabled = $false
        } else {
            $connectButton.Add_Click({ConnectModule})
        }
        $modButton = $connectButton
    } else {
        $installButton = New-Object System.Windows.Forms.Button
        $installButton.Size = New-Object System.Drawing.Size(120,23)
        $installButton.Text = "Install"
        $installButton.Add_Click({InstallModule})
        $modButton = $installButton
    }

    $gpBox = New-Object System.Windows.Forms.GroupBox
    $gpBox.Controls.Add($modLabel)
    $gpBox.Controls.Add($modButton)

    if($Script:connected){
        DisplayTenantInfo $gpBox
    }

}

Function GenerateGUI {
    
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $gui = New-Object System.Windows.Forms.Form
    $gui.Text = 'Azure AD User Admin Tool'
    $gui.Width = 600
    $gui.Height = 400

    $flowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $flowPanel.Width = 590
    $flowPanel.Height = 390

    DisplayConnection $flowPanel

    $gui.Controls.Add($flowPanel)
    $gui.ShowDialog()

}

try {
    GenerateGUI
} finally {
    Write-Host Good Job
}
