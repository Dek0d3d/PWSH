####################################################################
# This Script Renders a GUI which will assist in running commands. #
####################################################################

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
}

Function GenerateGUI {
    
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $script:gui = New-Object System.Windows.Forms.Form
    $gui.Text = 'Azure AD User Admin Tool'
    $gui.Width = 600
    $gui.Height = 400

    $flowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $flowPanel.AutoSize = $true
    
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Microsoft Graph Powershell SDK installed:" + $mgInstalled
    $label.Width = 400
    $Label.AutoSize = $true
    $flowPanel.Controls.Add($Label)

    $connectButton = New-Object System.Windows.Forms.Button
    $connectButton.Size = New-Object System.Drawing.Size(120,23)
    $connectButton.Text = "Connect"
    $connectButton.Add_Click({ConnectModule})

    $installButton = New-Object System.Windows.Forms.Button
    $installButton.Size = New-Object System.Drawing.Size(120,23)
    $installButton.Text = "Install"
    $installButton.Add_Click({InstallModule})

    $showButton = if(CheckModule){
        $connectButton
    }
    else {
        $installButton
    }

    $flowPanel.Controls.Add($Label)
    $flowPanel.Controls.Add($showButton)
    $gui.Controls.Add($flowPanel)
    $gui.ShowDialog()

}

GenerateGUI
