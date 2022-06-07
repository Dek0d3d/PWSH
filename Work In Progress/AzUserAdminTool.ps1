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
    DisplayConnection $gui
}

Function ConnectModule {
    Import-Module Microsoft.Graph
    Connect-MgGraph
    GetConnection
    DisplayConnection $gui
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
    param([Parameter()][System.Windows.Forms.FlowLayoutPanel]$thisPanel)

    $TenantLabel = New-Object System.Windows.Forms.Label
    $TenantLabel.TextAlign = "MiddleLeft"
    $TenantLabel.Text = "$(($Script:tenant).DisplayName) : $(($Script:tenant).Id)"
    $TenantLabel.Width = 280
    $TenantLabel.Height = 25
    
    $UserLabel = New-Object System.Windows.Forms.Label
    $UserLabel.TextAlign = "MiddleRight"
    $UserLabel.Text = "$((Get-MgContext).Account)"
    $UserLabel.Width = 280
    $UserLabel.Height = 25
    
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
    $UPNlbl.AutoSize
    $UPNlbl.TextAlign = 'MiddleLeft'
    $thisPanel.Controls.Add($UPNlnl)

    $UPNtxtBox = New-Object System.Windows.Forms.TextBox
    $UPNtxtBox.AutoSize
    $UPNtxtBox.TextAlign = 'MiddleLeft'
    $UPNtxtBox.Name = "UPNText"
    $thisPanel.Controls.Add($UPNtxtBox)

    ForEach($Prop in $(ValidProps).GetEnumerator()){
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "$($Prop.Name):"
        $lbl.AutoSize
        $lbl.TextAlign = 'MiddleLeft'
        $thisPanel.Controls.Add($lbl)

        $txtbx = New-Object System.Windows.Forms.TextBox
        $txtbx.Name = "$($Prop.Name)txt"
        $txtbx.AutoSize
        $txtbx.TextAlign = 'MiddleLeft'
        $thisPanel.Controls.Add($lbl)
    }

}

Function PopulateDataPanel {
    param([Parameter()][System.Windows.Forms.TableLayoutPanel]$thisPanel)

    $Users = $(GetAzAdUserData).GetEnumerator()
    ForEach($User in $Users){
        
    }

}

Function GenerateDataPanel {
    param([Parameter()][System.Windows.Forms.Form]$gui)

    $DataPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $DataPanel.RowCount = 0
    $DataPanel.ColumnCount = 2
    $DataPanel.Size = New-Object System.Drawing.Size(575,250)
    $DataPanel.Location = New-Object System.Drawing.Point(5,110)
    $DataPanel.AutoSize
    $DataPanel.BackColor = "White"
    $DataPanel.BorderStyle = "Fixed3D"

    $GetModeRB = New-Object System.Windows.Forms.RadioButton
    $GetModeRB.Width = 200
    $GetModeRB.Location = New-Object System.Drawing.Point(100,15)
    $GetModeRB.Text = "Display/Modify"

    $CreateModeRB = New-Object System.Windows.Forms.RadioButton
    $CreateModeRB.Width = 200
    $CreateModeRB.Location = New-Object System.Drawing.Point(330,15)
    $CreateModeRB.Text = "Create"

    $ModeGroupBox = New-Object System.Windows.Forms.GroupBox
    $ModeGroupBox.Text = "Mode of Operation"
    $ModeGroupBox.Size = New-Object System.Drawing.Size(575,45)
    $ModeGroupBox.Location = New-Object System.Drawing.Point(5,60)
    $ModeGroupBox.Controls.AddRange(@($GetModeRB,$CreateModeRB))

    $gui.Controls.Add($ModeGroupBox)

    SetupDataPanel $DataPanel
    $gui.Controls.Add($DataPanel)

}

Function DisplayConnection {
    param([Parameter()][System.Windows.Forms.Form]$gui)

    if ($gui.Controls.ContainsKey("ConnectPanel")){
        $gui.Controls.RemoveByKey("ConnectPanel")
    }

    $modLabel = New-Object System.Windows.Forms.Label
    $modLabel.TextAlign = "MiddleLeft"
    $modLabel.Text = "Microsoft Graph Powershell SDK installed:" + $(CheckModule)
    $modlabel.Width = 440
    $modLabel.Height = 25

    $modButton = New-Object System.Windows.Forms.Button
    $modButton.Size = New-Object System.Drawing.Size(120,20)
    UpdateModButton $modButton

    $connectPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $connectPanel.Name = "ConnectPanel"
    $connectPanel.AutoSize = $true
    $connectPanel.MaximumSize = $gui.Size
    $connectPanel.Width = 575
    $connectPanel.Height = 50
    $connectPanel.Location = New-Object System.Drawing.Point(5,5)
    $connectPanel.BackColor = "White"
    $connectPanel.BorderStyle = "Fixed3D"
    $connectPanel.Controls.Add($modLabel)
    $connectPanel.Controls.Add($modButton)

    if($Script:connected){
        DisplayTenantInfo $connectPanel
        $gui.Controls.Add($connectPanel)
        GenerateDataPanel $gui
    } else {
        $gui.Controls.Add($connectPanel)
    }

}

Function GenerateGUI {
    
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $gui = New-Object System.Windows.Forms.Form
    $gui.Text = 'Azure AD User Admin Tool'
    $gui.Width = 600
    $gui.Height = 400
    $gui.AutoSize

    DisplayConnection $gui

    $gui.ShowDialog()

}

try {
    GenerateGUI
} finally {
    
}
