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

Function InitModPanel {
    $gui.Close()
    $gui.Dispose()
    GenerateGUI
}

Function InstallModule {
    Install-Module Microsoft.Graph -Scope CurrentUser
    DisplayConnection $guiPanel
}

Function ConnectModule {
    Import-Module Microsoft.Graph
    Connect-MgGraph
    GetConnection
    DisplayConnection $guiPanel
}

Function UpdateModButton {
    param([Parameter()][System.Windows.Forms.Button]$modButton)

    if(CheckModule){
        if($Script:tenant){
            $modButton.Text = "Connected"
            $modButton.Enabled = $false
        } else {
            $modButton.Text = "Connect"
            $modButton.Add_Click({ConnectModule})
        }
    } else {
        $modButton.Text = "Install"
        $modButton.Add_Click({InstallModule})
    }
    
}

Function GetConnection {
    $Script:tenant = Get-MgOrganization | Select-Object -Property *
    $Script:connected = $true
}

Function DisplayTenantInfo {
    param([Parameter()][System.Windows.Forms.TableLayoutPanel]$thisPanel)

    $TenantLabel = New-Object System.Windows.Forms.Label
    $TenantLabel.AutoSize = $true
    $TenantLabel.TextAlign = "MiddleLeft"
    $TenantLabel.Text = "$(($Script:tenant).DisplayName) : $(($Script:tenant).Id)"
    $TenantLabel.Anchor = 'Bottom,Left'
    
    $UserLabel = New-Object System.Windows.Forms.Label
    $UserLabel.TextAlign = "MiddleRight"
    $UserLabel.AutoSize = $true
    $UserLabel.Text = "$((Get-MgContext).Account)"
    $UserLabel.Anchor = 'Bottom,Right'
    
    $thisPanel.Controls.Add($TenantLabel)
    $thisPanel.Controls.Add($UserLabel)
}
Function GetAzAdUserData {

    $UsersDataHash = @{}

    $Users = Get-MgUser | Select-Object -Property *
    ForEach ($User in $Users) { 
        $UserData = ConvertUserToHash $User
        $UPN = $UserData.UserPrincipalName
        $UsersDataHash.Add($UPN,$UserData)
    }

    return $UsersDataHash

}

Function ConvertUserToHash {
    param([Parameter()][PSCustomObject]$User)

    $UserHash = @{}
    $User | Get-Member | Where-Object MemberType -ne "Method" | ForEach-Object{
        $PropName = $_.Name
        $PropDef = $_.Definition
        $PropDef = $($PropDef.Split(' '))[1]
        $PropDef = $($PropDef.Split('='))[1]
        if(($PropDef -notlike "Microsoft.Graph*") -and ($PropDef -notlike "System.*")){
            $UserHash.Add($PropName,$PropDef)
        }
    }

    return $UserHash
}

Function ValidProps {

    $NewPropList = @{}
    $Props = GetAzAdUserData
    $Props = $($Props.GetEnumerator() | Select-Object -First 1 | Select-Object -ExpandProperty Value)
    $PropNames = $Props.GetEnumerator() | Select-Object Name
    $PropList = $(Get-Command New-MgUser).ParameterSets[0] | 
        Select-Object -ExpandProperty parameters | 
        Where-Object ParameterType -notlike "Microsoft.Graph*"
    ForEach($Prop in $PropNames){
        if($PropList.Name -contains $Prop.Name){
            $Name = $Prop.Name 
            $Type = $PropList[$PropList.Name.IndexOf($Prop.Name)].ParameterType
            $NewPropList += @{$Name = $Type}
        }
    }

    return $NewPropList

}

Function SetupDataPanel {
    param([Parameter()][System.Windows.Forms.TableLayoutPanel]$thisPanel)

    $UPNlbl = New-Object System.Windows.Forms.Label
    $UPNlbl.Text = "UserPrincipalName:"
    $UPNlbl.AutoSize = $true
    $UPNlbl.Dock = 'Fill'
    $thisPanel.Controls.Add($UPNlbl)

    $UPNtxtBox = New-Object System.Windows.Forms.TextBox
    $UPNtxtBox.Name = "UPNText"
    $UPNtxtBox.AutoSize = $true
    $UPNtxtBox.Dock = 'Fill'
    $thisPanel.Controls.Add($UPNtxtBox)

    ForEach($Prop in $(ValidProps).GetEnumerator()){
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "$($Prop.Name):"
        $lbl.AutoSize = $true
        $lbl.Dock = 'Fill'
        $thisPanel.Controls.Add($lbl)

        $txtbx = New-Object System.Windows.Forms.TextBox
        $txtbx.Name = "$($Prop.Name)txt"
        $txtbx.Text = ""
        $txtbx.Enabled = $false
        $txtbx.AutoSize = $true
        $txtbx.Dock = 'Fill'
        $thisPanel.Controls.Add($txtbx)
    }

}

Function PopulateDataPanel {
    param([Parameter()][System.Windows.Forms.TableLayoutPanel]$thisPanel)

    $Users = $(GetAzAdUserData).GetEnumerator()
    ForEach($User in $Users){
        
    }

}

Function GenerateDataPanel {
    param([Parameter()][System.Windows.Forms.TableLayoutPanel]$guiPanel)

    $ModeGroupBox = New-Object System.Windows.Forms.GroupBox
    $ModeGroupBox.Text = "Mode of Operation"
    $ModeGroupBox.AutoSize
    $ModeGroupBox.AutoSizeMode = 'GrowAndShrink'
    $ModeGroupBox.MaximumSize = '4000,50'
    $ModeGroupBox.Dock = 'Fill'
    $ModeGroupBox.Padding = '5,5,5,5'
    $guiPanel.Controls.Add($ModeGroupBox)

    $GetModeRB = New-Object System.Windows.Forms.RadioButton
    $GetModeRB.Text = "Display/Modify"
    $GetModeRB.AutoSize = $true
    $GetModeRB.Location = '100,20'

    $CreateModeRB = New-Object System.Windows.Forms.RadioButton
    $CreateModeRB.Text = "Create"
    $CreateModeRB.AutoSize = $true
    $CreateModeRB.Location = '250,20'
    $ModeGroupBox.Controls.AddRange(@($GetModeRB,$CreateModeRB))

    $DataPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $DataPanel.RowCount = 0
    $DataPanel.ColumnCount = 2
    $DataPanel.AutoSize = $true
    $DataPanel.AutoSizeMode = 'GrowAndShrink'
    $DataPanel.Padding = '5,5,5,5'
    $DataPanel.BackColor = "White"
    $DataPanel.BorderStyle = "Fixed3D"
    $DataPanel.Anchor = 'Left,Right,Bottom'
    
    SetupDataPanel $DataPanel
    $guiPanel.Controls.Add($DataPanel)

}

Function DisplayConnection {
    param([Parameter()][System.Windows.Forms.TableLayoutPanel]$guiPanel)

    if ($guiPanel.Controls.ContainsKey("ConnectPanel")){
        $guiPanel.Controls.RemoveByKey("ConnectPanel")
    }

    $connectPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $connectPanel.Name = "ConnectPanel"
    $connectPanel.RowCount = 2
    $connectPanel.ColumnCount = 2
    $connectPanel.AutoSize = $true
    $connectPanel.AutoSizeMode = 'GrowAndShrink'
    $connectPanel.Padding = '5,5,5,5'
    $connectPanel.BackColor = "White"
    $connectPanel.BorderStyle = "Fixed3D"
    $connectPanel.Anchor = 'Top,Left,Right,Bottom'
    $guiPanel.Controls.Add($connectPanel)

    $modLabel = New-Object System.Windows.Forms.Label
    $modLabel.TextAlign = "MiddleLeft"
    $modLabel.Text = "Microsoft Graph Powershell SDK installed:" + $(CheckModule)
    $modLabel.AutoSize = $true
    $modlabel.Anchor = 'Top,Left'

    $modButton = New-Object System.Windows.Forms.Button
    UpdateModButton $modButton
    $modButton.AutoSize = $true
    $modButton.AutoSizeMode = 'GrowAndShrink'
    $modButton.Anchor = 'Top,Right'
    
    $connectPanel.Controls.Add($modLabel)
    $connectPanel.Controls.Add($modButton)

    if($Script:connected){
        DisplayTenantInfo $connectPanel
        GenerateDataPanel $guiPanel
    } 

}

Function GenerateGUI {
    
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $gui = New-Object System.Windows.Forms.Form
    $gui.Text = 'Azure AD User Admin Tool'
    $gui.Size = '600,400'

    $guiPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $guiPanel.RowCount = 3
    $guiPanel.ColumnCount = 1
    $guiPanel.AutoSize = $true
    $guiPanel.AutoSizeMode = 'GrowAndShrink'
    $guiPanel.Padding = '5,5,5,5'
    $guiPanel.Dock = 'Fill'
    $guiPanel.AutoScroll = $true
    $guiPanel.
    $gui.Controls.Add($guiPanel)

    DisplayConnection $guiPanel

    $gui.ShowDialog()

}

try {
    GenerateGUI
} finally {
    
}
