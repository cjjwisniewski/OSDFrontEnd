<#--------------------------
Name: OSDFrontEnd.ps1
Author: Cameron Wisniewski
Date: 6/20/19
Version: 5.0
Comment: Basic frontend UI for OSD through SCCM
Link: N/A
Notes: See README.txt for further documentation. 
Recent revisions:
Features to add: Redo height definitions to scale to screen resolution ($WindowHeight) and utilize scrollbars
                 Retheme to look mildly better (aero maybe?)
--------------------------#>

#Load in Windows forms components
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#Declare some vars
$SiteServer = "<SITESERVER>"
$SiteCode = "<SITECODE>"
$SoftwareBuildGroupName = "<Name of group in TS containing software builds>" #As defined in task sequence
$AdditionalSoftwareGroupName = "<Name of group in TS containing additional software>" #As defined in task sequence
$IPAddressPrefix = "<First octet of IP address of active network adapter>" #Used to determine the network adapter connected to the corporate network
$CompanyLogo = "$PSScriptRoot\Files\logo_header.png" #Path to company logo

$WindowHeight = ((Get-WMIObject -Class Win32_DisplayConfiguration).PelsHeight - 100) #https://docs.microsoft.com/en-us/previous-versions/aa394137(v%3Dvs.85)
$WindowWidth = ((Get-WMIObject -Class Win32_DisplayConfiguration).PelsWidth - 100) 
$OSDTSVarName = $null
$TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
[xml]$TaskSequenceXML = $TSEnv.Value("_SMSTSTaskSequence")

#Load in SoftwareBuildList from _SMSTSTaskSequence
$SoftwareBuildList = (($TaskSequenceXML | Select-XML -XPath "//group" | Select-Object -ExpandProperty "Node" | Where-Object -Property "Name" -Like "$SoftwareBuildGroupName").SubTaskSequence | Select-Object -Property "Name").Name

#Load in TimeZoneList.txt and MasterTimeZoneList.csv
$TimeZoneCSV = Import-CSV -Path "$PSScriptRoot\Files\MasterTimeZoneList.csv"
$TimeZoneList = Get-Content -Path "$PSScriptRoot\Files\TimeZoneList.txt"

#Load in AdditionalSoftwareList from _SMSTSTaskSequence
$AdditionalSoftwareList = (($TaskSequenceXML | Select-XML -XPath "//group" | Select-Object -ExpandProperty "Node" | Where-Object -Property "Name" -Like "$AdditionalSoftwareGroupName").Step | Select-Object -Property "Name").Name

#InfoBox Text Stuff
$SerialNumber = (Get-WMIObject -Class Win32_BIOS).SerialNumber
$ComputerName = "CCIC"+$SerialNumber
$Manufacturer = (Get-WMIObject -Class Win32_ComputerSystem).Manufacturer
$Model = (Get-WMIObject -Class Win32_ComputerSystem).Model
$Version = (Get-WMIObject -Class Win32_ComputerSystemProduct).Version
$ActiveAdapter = (Get-WMIObject -Class Win32_NetworkAdapterConfiguration -Property * | Where-Object -Property "IPAddress" -like $IPAddressPrefix)
$IPAddress = $ActiveAdapter.IPAddress[0]
$MACAddress = $ActiveAdapter.MACAddress
$UUID = (Get-WMIObject -Class Win32_ComputerSystemProduct -Property *).UUID

$InfoBoxText = "Computer Name: $ComputerName `nManufacturer: $Manufacturer`nModel: $Model`nVersion: $Version`nIP Address: $IPAddress`nMAC Address: $MACAddress`nUUID: $UUID"

#Confirm Button functionality, writes TS variables based on selections within this UI
$ConfirmButton_Click = {
    #Read software build selection
    $SoftwareBuildBox.Controls | ForEach-Object {
        if($_.Checked) {
            $OSDTSVarName = $SoftwareBuildList[$SoftwareBuildBox.Controls.IndexOf($_)]
        }
    }
    
    #Read time zone selection and map to OSDTimeZone value (see link). 
    #https://docs.microsoft.com/en-us/previous-versions/orphan-topics/ws.10/cc772783(v=ws.10)
    $TimeZoneValue = ($TimeZoneCSV | Where-Object Description -like $TimeZoneDropDown.SelectedItem)."Time Zone"

    #Read additional software selections
    $AdditionalSoftwareSelectionDisplay = "`nAdditional Software: `n"
    $AdditionalSoftwareBox.Controls | ForEach-Object {
        if($_.Checked) {
            $AdditionalSoftwareSelection += $AdditionalSoftwareList[$AdditionalSoftwareBox.Controls.IndexOf($_)]
            $AdditionalSoftwareSelectionDisplay += $AdditionalSoftwareList[$AdditionalSoftwareBox.Controls.IndexOf($_)]+"`n"
        }
    }

    #Clean up old PC name text
    $OSDOldPCName = $PCMOldPCNameInput.Text.ToUpper()
    if($OSDOldPCName) {
        $OSDOldPCNameDisplay = "`nOld PC Name: "
        $OSDOldPCNameDisplay += $OSDOldPCName
    }


    #Confirm settings, done this way to remove Additional Software if there's none selected.
    $Confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want image this computer with the following settings?`n`nSoftware Build: $OSDTSVarName`nTime Zone: "+$TimeZoneDropDown.SelectedItem+$AdditionalSoftwareSelectionDisplay+$OSDOldPCNameDisplay, "Confirm Image Settings", "OkCancel", "Exclamation")

    if($Confirm -eq "Ok"){
        #Write OSDComputerName
        $TSEnv.Value("OSDComputerName") = $ComputerName

        #Write OSDSoftwareBuild
        $TSEnv.Value("OSDSoftwareBuild") = $OSDTSVarName

        #Write OSDTimeZone
        $TSEnv.Value("OSDTimeZone") = $TimeZoneValue

        #Write OSDAdditionalSoftware
        $TSEnv.Value("OSDAdditionalSoftware") = $AdditionalSoftwareSelection

        #Write OSDOldPCName
        if($OSDOldPCName) {
            $TSEnv.Value("OSDOldPCName") = $OSDOldPCName
        }

        $MainWindow.Close()
    }
}

#Credentials Button
$CredentialsButton_Click = {
    #Get credentials
    $Credential = Get-Credential -Message "Please enter your credentials prefixed by the domain."

    #Try to load in SCCM info       
    $script:CMDeviceObject = Get-WMIObject -ComputerName "$SiteServer" -NameSpace "root\sms\site_$SiteCode" -Query "SELECT * FROM sms_r_system WHERE smbiosguid = '$UUID'" -Credential $Credential
    
    #Enable buttons if it actually retrieved an object
    if(($error[0] -like "*Access is denied.*") -or ($Credential -eq $null)) {
        $CMBox.Text = "The specified credentials are invalid or may not have the necessary permissions."
    } else {
        $CMBox.Text = "The provided credentials for "+$Credential.UserName+" are valid."
        $SCCMReportButton.Enabled = $true
    }
}

#SCCM Report Button
$SCCMReportButton_Click = {
    if($CMDeviceObject) {
        $CMBox.Text = "Generating report. This may take some time."

        $CMName = $CMDeviceObject.Name
        $CMSystemOUName = ($CMDeviceObject.SystemOUName -split '[\r\n]')[($CMDeviceObject.SystemOUName -split '[\r\n]').Length-1]
        $CMLastLoggedOnUserName = $CMDeviceObject.LastLogonUserName
        $CMLastLogonTimeStamp = $CMDeviceObject.LastLogonTimestamp.Substring(4,2)+"/"+$CMDeviceObject.LastLogonTimestamp.Substring(6,2)+"/"+$CMDeviceObject.LastLogonTimestamp.Substring(0,4)
        $CMBoxText = "Last known CM record for this machine:`n`nLast Computer Name: $CMName `nLast User: $CMLastLoggedOnUserName `nLast Logon Time: $CMLastLogonTimeStamp `nLast System OU: $CMSystemOUName `n`nLast known collection memberships: `n"
        
        $Collections = Get-WMIObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Class SMS_Collection
        $Collections | ForEach-Object {
            if(Get-WMIObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Class SMS_CM_RES_COLL_$($_.CollectionID) -Filter "Name='$($ComputerName)'") {
                $CMBoxText += "`n"+$_.Name
                $CMBox.Text = $CMBoxText
                }
            }

        #Enable SCCM delete button
        $SCCMDeleteButton.Enabled = $true
    } else {
        $CMBoxText = "No record exists for this device."
    }
    $CMBox.Text = $CMBoxText
}

#SCCM Delete button
#Check creds against something???
$SCCMDeleteButton_Click = {
    $ConfirmDelete = [System.Windows.Forms.MessageBox]::Show("Are you sure you want delete this device from the SCCM database? This will remove it from all collections. You can not undo this action.", "Confirm Record Deletion", "OkCancel", "Exclamation")
    if($ConfirmDelete -eq "Ok") {
        $CMDeviceObject.PSBase.Delete()
    }
}

#Container for SoftwareBuildList radio buttons
$SoftwareBuildBox = New-Object System.Windows.Forms.GroupBox
$SoftwareBuildBox.Location = New-Object System.Drawing.Size(0,5)
$SoftwareBuildBox.Text = "Software Builds"
$SoftwareBuildBox.Size = New-Object System.Drawing.Size(($WindowWidth/2-5),(25*$SoftwareBuildList.Length+5))

#Dynamically draw Software Build List based on the content of SoftwareBuildList
$SoftwareBuildList | ForEach-Object {
    $SoftwareBuildListIndex = $SoftwareBuildList.IndexOf($_)
    $VariableName = "SoftwareBuild"+$SoftwareBuildListIndex
    New-Variable -Name "$VariableName" -Force
    $VariableName = New-Object System.Windows.Forms.RadioButton
    $VariableName.Location = New-Object System.Drawing.Size(5,(20+23*$SoftwareBuildListIndex))
    $VariableName.Size = New-Object System.Drawing.Size(300,19)
    $VariableName.Checked = $false
    $VariableName.Text = $SoftwareBuildList[$SoftwareBuildListIndex]
    $SoftwareBuildBox.Controls.Add($VariableName)
}

#Container for additional software selections
$AdditionalSoftwareBox = New-Object System.Windows.Forms.GroupBox
$AdditionalSoftwareBox.Text = "Additional Software"
$AdditionalSoftwareBox.Size = New-Object System.Drawing.Size(($SoftwareBuildBox.Size.Width-10),($SoftwareBuildBox.Size.Height))

#Dynamically draw Additional Software List based on the content of AdditionalSoftwareList.txt
$AdditionalSoftwareList | ForEach-Object {
    $AdditionalSoftwareListIndex = $AdditionalSoftwareList.IndexOf($_)
    $VariableName = "AdditionalSoftware"+$AdditionalSoftwareListIndex
    New-Variable -Name "$VariableName" -Force
    $VariableName = New-Object System.Windows.Forms.CheckBox
    $VariableName.Location = New-Object System.Drawing.Size(5,(20+23*$AdditionalSoftwareListIndex))
    $VariableName.Size = New-Object System.Drawing.Size(300,19)
    $VariableName.Checked = $false
    $VariableName.Text = $AdditionalSoftwareList[$AdditionalSoftwareListIndex]
    $AdditionalSoftwareBox.Controls.Add($VariableName)
}

#Declare UI components
#Base container window
$MainWindow = New-Object System.Windows.Forms.Form
$MainWindow.Text = "Crown Castle - Select Software Build"
$MainWindow.Size = New-Object System.Drawing.Size($WindowWidth,0)
$MainWindow.StartPosition = "CenterScreen"
$MainWindow.FormBorderStyle = "None"
$MainWindow.Font = New-Object System.Drawing.Font("Consolas",8,[System.Drawing.FontStyle]::Regular)
$MainWindow.MaximizeBox = $false

#Logo
$CCILogo = [System.Drawing.Image]::FromFile($CompanyLogo)
$CCILogo_Embed = New-Object Windows.Forms.PictureBox
$CCILogo_Embed.Width = $CCILogo.Size.Width
$CCILogo_Embed.Height = $CCILogo.Size.Height
$CCILogo_Embed.Size = New-Object System.Drawing.Size(240,64)
$CCILogo_Embed.Location = New-Object Drawing.Point((($MainWindow.Size.Width/2)-($CCILogo.Size.Width/2)),5)
$CCILogo_Embed.Image = $CCILogo
$MainWindow.Controls.Add($CCILogo_Embed)

#Instructions text
$Instructions = New-Object System.Windows.Forms.Label
$Instructions.Location = New-Object System.Drawing.Size(5,($CCILogo_Embed.Size.Height+10))
$Instructions.Size = New-Object System.Drawing.Size(($MainWindow.Size.Width-20),17)
$Instructions.TextAlign = "TopCenter"
$Instructions.AutoSize = $false
$Instructions.Text = "Please select the appropriate software build and time zone for this PC."
$MainWindow.Controls.Add($Instructions)

#TabControl
$TabControl = New-Object System.Windows.Forms.TabControl
$TabControl.DataBindings.DefaultDataSourceUpdateMode = 0
$TabControl.Location = New-Object System.Drawing.Size(5,($Instructions.Location.Y+$Instructions.Size.Height))
$MainWindow.Controls.Add($TabControl)

#Default page
$DefaultPage = New-Object System.Windows.Forms.TabPage
$DefaultPage.DataBindings.DefaultDataSourceUpdateMode = 0
$DefaultPage.Name = "Default Settings"
$DefaultPage.Text = "Default Settings"
$TabControl.Controls.Add($DefaultPage)

#Add default page controls that we defined earlier to help define sizes
$DefaultPage.Controls.Add($AdditionalSoftwareBox)
$DefaultPage.Controls.Add($SoftwareBuildBox)

#CM Page
$CMPage = New-Object System.Windows.Forms.TabPage
$CMPage.DataBindings.DefaultDataSourceUpdateMode = 0
$CMPage.Name = "CM Settings"
$CMPage.Text = "CM Settings"
$TabControl.Controls.Add($CMPage)

#Credentials Button
$CredentialsButton = New-Object System.Windows.Forms.Button
$CredentialsButton.Size = New-Object System.Drawing.Size(120,75)
$CredentialsButton.Location = New-Object System.Drawing.Size(0,5)
$CredentialsButton.Text = "Credentials"
$CredentialsButton.Add_Click($CredentialsButton_Click)
$CMPage.Controls.Add($CredentialsButton)

#SCCMReport Button
$SCCMReportButton = New-Object System.Windows.Forms.Button
$SCCMReportButton.Size = $CredentialsButton.Size
$SCCMReportButton.Location = New-Object System.Drawing.Size(0,($CredentialsButton.Location.Y+$CredentialsButton.Size.Height+5))
$SCCMReportButton.Text = "SCCM Information Report"
$SCCMReportButton.Add_Click($SCCMReportButton_Click)
$CMPage.Controls.Add($SCCMReportButton)
$SCCMReportButton.Enabled = $false

#SCCMDelete Button
$SCCMDeleteButton = New-Object System.Windows.Forms.Button
$SCCMDeleteButton.Size = $CredentialsButton.Size
$SCCMDeleteButton.Location = New-Object System.Drawing.Size(0,($SCCMReportButton.Location.Y+$SCCMReportButton.Size.Height+5))
$SCCMDeleteButton.Text = "Delete SCCM Record"
$SCCMDeleteButton.Add_Click($SCCMDeleteButton_Click)
$CMPage.Controls.Add($SCCMDeleteButton)
$SCCMDeleteButton.Enabled = $false

#Container for SCCM info display
$CMBox = New-Object System.Windows.Forms.RichTextBox
$CMBox.Size = New-Object System.Drawing.Size(($WindowWidth-$CredentialsButton.Size.Width-25),(25*$SoftwareBuildList.Length+30))
$CMBox.Location = New-Object System.Drawing.Size(($CredentialsButton.Size.Width+5),5)
$CMBox.ReadOnly = $true
$CMBox.Text = "Please enter your credentials."
$CMBox.Font = New-Object System.Drawing.Font("Consolas",8,[System.Drawing.FontStyle]::Regular)
$CMPage.Controls.Add($CMBox)

#PCMover Page
$PCMoverPage = New-Object System.Windows.Forms.TabPage
$PCMoverPage.DataBindings.DefaultDataSourceUpdateMode = 0
$PCMoverPage.Name = "PCMover Settings"
$PCMoverPage.Text = "PCMover Settings"
$TabControl.Controls.Add($PCMoverPage)

#PCMover Label
$PCMInstructions = New-Object System.Windows.Forms.Label
$PCMInstructions.Location = New-Object System.Drawing.Size(5,5)
$PCMInstructions.Size = New-Object System.Drawing.Size(($WindowWidth-20),17)
$PCMInstructions.TextAlign = "TopCenter"
$PCMInstructions.AutoSize = $false
$PCMInstructions.Text = "If this is a PC for an existing user that has had their data backed up using PCMover, please enter the name of the old PC."
$PCMoverPage.Controls.Add($PCMInstructions)

#PCMover Old PC Name
$PCMOldPCNameInput = New-Object System.Windows.Forms.TextBox
$PCMOldPCNameInput.Location = New-Object System.Drawing.Size((($WindowWidth/2-100)),50)
$PCMOldPCNameInput.Size = New-Object System.Drawing.Size(200,25)
$PCMOldPCNameInput.Font = New-Object System.Drawing.Font("Consolas",8,[System.Drawing.FontStyle]::Regular)
$PCMOldPCNameInput.Text = $TSEnv.Value("OSDOldPCName")
$PCMoverPage.Controls.Add($PCMOldPCNameInput)

#Time zone dropdown list
$TimeZoneDropDown = New-Object System.Windows.Forms.ComboBox
$TimeZoneDropDown.Location = New-Object System.Drawing.Size(0,($SoftwareBuildBox.Location.Y+$SoftwareBuildBox.Size.Height))
$TimeZoneDropDown.Size = New-Object System.Drawing.Size(($WindowWidth-20),50)
$TimeZoneDropDown.DropDownStyle = "DropDownList"
$DefaultPage.Controls.Add($TimeZoneDropDown)

#Confirm button
$ConfirmButton = New-Object System.Windows.Forms.Button
$ConfirmButton.Size = New-Object System.Drawing.Size(120,120)
$ConfirmButton.Text = "Confirm"
$ConfirmButton.Add_Click($ConfirmButton_Click)
$MainWindow.Controls.Add($ConfirmButton)

#Container for machine info display
$InfoBox = New-Object System.Windows.Forms.RichTextBox
$InfoBox.ReadOnly = $true
$InfoBox.Text = $InfoBoxText
$InfoBox.Font = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Regular)
$InfoBox.Size = New-Object System.Drawing.Size(($WindowWidth-$ConfirmButton.Size.Width-15),$ConfirmButton.Size.Height)
$MainWindow.Controls.Add($InfoBox)

#Default select the Base build at the 0 index
$SoftwareBuildBox.Controls[0].Checked = $true

#Populate TimeZoneDropDown
$TimeZoneList | ForEach-Object {
    $TimeZoneDropDown.Items.Add($_) | Out-Null
}

#Default select Eastern Time at the 0 index
$TimeZoneDropDown.SelectedIndex = 0

#Redefine sizes/locations as defined by the size of the component pieces
#This is a wreck right now because of the dependencies for dynamic sizing... need to figure out a better way to do this so I can order things appropriately eventually (tm)
$MainWindow.Size = New-Object System.Drawing.Size($WindowWidth,($SoftwareBuildBox.Size.Height+$CCILogo_Embed.Size.Height+$Instructions.Size.Height+$ConfirmButton.Size.Height+$TimeZoneDropDown.Size.Height+55))
$TabControl.Size = New-Object System.Drawing.Size(($WindowWidth-10),($SoftwareBuildBox.Size.Height+55))

$AdditionalSoftwareBox.Location = New-Object System.Drawing.Size(($TabControl.Size.Width/2),($SoftwareBuildBox.Location.Y))

$InfoBox.Location = New-Object System.Drawing.Size(5,($MainWindow.Size.Height-$InfoBox.Size.Height-5))
$ConfirmButton.Location = New-Object System.Drawing.Size(($MainWindow.Size.Width-$ConfirmButton.Size.Width-5),($MainWindow.Size.Height-$ConfirmButton.Size.Height-5))

#Set pages to enabled or not?
$DefaultPage.Enabled = $true
$CMPage.Enabled = $true
$PCMoverPage.Enabled = $true

#Show the UI
$MainWindow.TopMost = $true
$MainWindow.Add_Shown({$MainWindow.Activate()})
[void] $MainWindow.ShowDialog()