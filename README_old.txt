#--------------------------
# Name: CCICOSDFrontEnd.ps1
# Author: Cameron Wisniewski
# Date: 9/28/2018
# Version: 4.180928
# Comment: Basic frontend UI for OSD through SCCM
# Link: N/A
#--------------------------

- Purpose -
In an attempt to make the OSD process as simple as possible for our techs, this was developed to
allow them to set build parameters for each machine as part of the OSD process rather than having
to add the PC to various collections in SCCM after the fact (though that may still be necessary
in some instances). In general, to simply change options available within the script, no changes
to the script itself are necessary. Everything in the general settings tab is defined by text
files in the Files directory or within the task sequence itself. Simply adding a line to those 
text files and/or making the requisite changes to the OSD task sequence is enough to add 
that additional functionality. Files you may need to edit are:

TimeZoneList.txt
For reference, time zone documentation can be found here:
https://docs.microsoft.com/en-us/previous-versions/orphan-topics/ws.10/cc772783(v=ws.10) 
This text file contains the GMT+/-X values for readability and time zone accuracy, and this is
used to populate the time zone selection drop down. When the tech is prompted to confirm their 
selections, this value is then switched for the center column value (e.g. "(GMT-05:00) Eastern 
Time (US & Canada)" > "Eastern Standard Time") which is used to define the final value written 
to the TSV OSDTimeZone. The values contained in this file only define what's in the dropdown. 
The csv used for value comparison and for the final OSDTimeZone value is a "master list" as per 
the link above. 

- Script/Documentation -

#Load in Windows forms components
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

Self explanatory, this loads components necessary to display windows forms used for the UI

#Declare some vars
$SiteServer = "CCICUSSCCM1.us.crowncastle.com"
$SiteCode = "CCI"
$SoftwareBuildGroupName = "Title-Specific Software Builds"
$AdditionalSoftwareGroupName = "Additional Software Selections"
$IPAddressPrefix = "10.*"

These are variables you may want to change depending on the task sequence this is used in. 
$SiteServer = SCCM site server
$SiteCode = SCCM site code
$SoftwareBuildGroupName = Task sequence group name for Software Builds
$AdditionalSoftwareGroupName = Task sequence group name for Additional Software
$IPAddressPrefix = Expected IP address prefix for active network adapter

#Other declarations
$WindowWidth = 800
$OSDTSVarName = $null
$TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
[xml]$TaskSequenceXML = $TSEnv.Value("_SMSTSTaskSequence")

Don't edit this with the exception of maybe WindowWidth to resize things a bit. This sets
a OSDTSVarName to null to stop a script error I ran into, sets the TS environment, then 
writes the TS XML from that to a variable.

#Load in SoftwareBuildList from _SMSTSTaskSequence
$SoftwareBuildList = (($TaskSequenceXML | Select-XML -XPath "//group" | Select-Object -ExpandProperty "Node" | Where-Object -Property "Name" -Like "$SoftwareBuildGroupName").SubTaskSequence | Select-Object -Property "Name").Name

This defines the Software Build List zone based on the content of the task sequence itself.
Adding/removing "Run Task Sequence" steps from the group defined in $SoftwareBuildGroupName
will cause the same changes to occur in the UI. Loading it directly out of the task sequence
allows the list to be "live" (and not require any kind of authentication) so that any changes
made to that section of the task sequence are immediately reflected in the frontend. 

#Load in TimeZoneList.txt and MasterTimeZoneList.csv
$TimeZoneCSV = Import-CSV -Path "$PSScriptRoot\Files\MasterTimeZoneList.csv"
$TimeZoneList = Get-Content -Path "$PSScriptRoot\Files\TimeZoneList.txt"

This defines the 'master' time zone list from a CSV file packaged with the script, and the 
text file defines which of those are actually displayed. Adding/removing entries to the text
file will add/remove that functionality from the frontend. It's best just to copy/paste from
the CSV when doing this to make sure there aren't any spelling errors or something. 

#Load in AdditionalSoftwareList from _SMSTSTaskSequence
$AdditionalSoftwareList = (($TaskSequenceXML | Select-XML -XPath "//group" | Select-Object -ExpandProperty "Node" | Where-Object -Property "Name" -Like "$AdditionalSoftwareGroupName").Step | Select-Object -Property "Name").Name

Similar to the software build list, this constructs the Addtional Software Selections
group based on the group defined in $AdditionalSoftwareGroupName.

#InfoBox Text Stuff
$SerialNumber = (Get-WMIObject -Class win32_bios).SerialNumber
$ComputerName = "CCIC"+$SerialNumber
$Manufacturer = (Get-WMIObject -Class win32_computersystem).Manufacturer
$Model = (Get-WMIObject -Class win32_computersystem).Model
$Version = (Get-WMIObject -Class win32_computersystemproduct).Version
$ActiveAdapter = (Get-WMIObject -Class win32_networkadapterconfiguration -Property * | Where-Object -Property "IPAddress" -like $IPAddressPrefix)
$IPAddress = $ActiveAdapter.IPAddress[0]
$MACAddress = $ActiveAdapter.MACAddress
$UUID = (Get-WMIObject -Class win32_computersystemproduct -Property *).UUID
$InfoBoxText = "Computer Name: $ComputerName `nManufacturer: $Manufacturer`nModel: $Model`nVersion: $Version`nIP Address: $IPAddress`nMAC Address: $MACAddress`nUUID: $UUID"

This defines basic information about the computer for the technican from WMI. 

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
    $AdditionalSoftwareBox.Controls | ForEach-Object {
        if($_.Checked) {
            $AdditionalSoftwareSelection += $AdditionalSoftwareList[$AdditionalSoftwareBox.Controls.IndexOf($_)]
            $AdditionalSoftwareSelectionDisplay = "`nAdditional Software: `n"
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


This script block defines the functionality of the Confirm button in the frontend. It
will loop through each entry in the software build selection to determine which is checked
and write that to a variable, match the time zone dropdown selection to the corresponding
OSDTimeZone compatible value in the imported CSV file, and loop through the additional software
selections to write anything checked there to a string (there's probably a better way to handle
that one that I didn't think of yet, but this works well enough for my purposes.), before finally
writing the OldPCName (for use with PCMover) to a variable. This is all shown in the confirmation
prompt, and if the result of that is "OK", those values will be written to task sequence variables
so that the relating steps can be run later on in the task sequence or the correct config is
defined. 

#Credentials Button
$CredentialsButton_Click = {
    #Get credentials
    $Credential = Get-Credential -Message "Please enter your credentials prefixed by CCIC\"

    #Try to load in SCCM info       
    $script:CMObject = Get-WMIObject -ComputerName "$SiteServer" -NameSpace "root\sms\site_$SiteCode" -Query "SELECT * FROM sms_r_system WHERE smbiosguid = '$UUID'" -Credential $Credential
    
    #Reenable buttons if it actually retrieved an object
    if(($error[0] -like "*Access is denied.*") -or ($Credential -eq $null)) {
        $CMBox.Text = "The credentials for user "+$Credential.UserName+" are invalid or may not have the necessary permissions."
    } else {
        $CMBox.Text = "The provided credentials for "+$Credential.UserName+" are valid."
        $SCCMReportButton.Enabled = $true
    }
}

This script block defines the functionality of the Credentials button on the SCCM tab. Get-Credential
is used to collect credentials before being passed through to a credentialed WMI query to the 
site server for the current computer to determine if they are valid or have the right permissions. 
If they work, the Report button is enabled. 

#SCCM Report Button
$SCCMReportButton_Click = {
    if($CMObject) {
        $CMBox.Text = "Generating report. This may take a long time."

        $CMName = $CMObject.Name
        $CMSystemOUName = ($CMObject.SystemOUName -split '[\r\n]')[($CMObject.SystemOUName -split '[\r\n]').Length-1]
        $CMLastLoggedOnUserName = $CMObject.LastLogonUserName
        $CMLastLogonTimeStamp = $CMObject.LastLogonTimestamp.Substring(4,2)+"/"+$CMObject.LastLogonTimestamp.Substring(6,2)+"/"+$CMObject.LastLogonTimestamp.Substring(0,4)
        $CMBoxText = "Last known CM record for this machine:`n`nLast Computer Name: $CMName `nLast User: $CMLastLoggedOnUserName `nLast Logon Time: $CMLastLogonTimeStamp `nLast System OU: $CMSystemOUName `n`nLast known collection memberships: `n"
        
        $Collections = Get-WMIObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Class SMS_Collection
        $Collections | ForEach-Object {
            if(Get-WMIObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Class SMS_CM_RES_COLL_$($_.CollectionID) -Filter "Name='$($ComputerName)'") {
                $CMBoxText += "`n"+$_.Name
                $CMBox.Text = $CMBoxText
                }
            }

        #Enable SCCM delete button
        $SCCMDeleteButton.Enabled = $false
    } else {
        $CMBoxText = "There is no CM record for this machine."
    }
    $CMBox.Text = $CMBoxText
}

This script block defines the SCCM Report button. If CMObject exists (as defined in the credentials step), 
this will pull some info from that object about the machine, and run through each collection available 
to the specified credentials to determine which the PC is a member of. As a separate query is made for
each collection, this takes a really long time (and there's probably a better way to do it that I haven't
found yet.). The functionality is there to enable the delete button, but that's currently disabled as I'm
not sure if I want that to be there or not yet (need to test with permissions)

#SCCM Delete button
$SCCMDeleteButton_Click = {
    $ConfirmDelete = [System.Windows.Forms.MessageBox]::Show("Are you sure you want delete this computer from the SCCM database? This will remove it from all collections. You can not undo this action.", "Confirm Record Deletion", "OkCancel", "Exclamation")
    if($ConfirmDelete -eq "Ok") {
        $CMObject.PSBase.Delete()
    }
}

This script block defines the SCCM delete button, which will remove the current machine's SCCM record
from the database (if the specified credentials have the required permissions I guess?). This is
currently disabled as I need to do more testing to find the limitations of this. 

#Container for SoftwareBuildList radio buttons
$SoftwareBuildBox = New-Object System.Windows.Forms.GroupBox
$SoftwareBuildBox.Location = New-Object System.Drawing.Size(0,5)
$SoftwareBuildBox.Text = "Software Builds"
$SoftwareBuildBox.Size = New-Object System.Drawing.Size(($WindowWidth/2-5),(25*$SoftwareBuildList.Length+5))

#Container for additional software selections
$AdditionalSoftwareBox = New-Object System.Windows.Forms.GroupBox
$AdditionalSoftwareBox.Text = "Additional Software"
$AdditionalSoftwareBox.Size = New-Object System.Drawing.Size(($SoftwareBuildBox.Size.Width-10),($SoftwareBuildBox.Size.Height))

These define the containers for the SoftwareBuildList and AdditionalSoftwareList selections. 

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

This script block dynamically creates and populates the SoftwareBuildList container selections.
For each entry in the list, a radiobutton object is created and added to the SoftwareBuildBox.

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

This script block works similar to the previous one, but with the AdditionalSoftwareList
variable and checkboxes instead of radiobuttons since we may want to select more than one 
of these at any given time. 

THIS NEXT STUFF WILL DEFINE BASIC UI COMPONENTS

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
$CCILogo = [System.Drawing.Image]::FromFile("$PSScriptRoot\Files\logo_header.png")
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

END DECLATATION OF BASIC UI COMPONENTS

#Default select the Base build at the 0 index
$SoftwareBuildBox.Controls[0].Checked = $true

#Populate TimeZoneDropDown
$TimeZoneList | ForEach-Object {
    $TimeZoneDropDown.Items.Add($_) | Out-Null
}

#Default select Eastern Time at the 0 index
$TimeZoneDropDown.SelectedIndex = 0

This defines the default selections for a basic build in the UI.

#Redefine sizes/locations as defined by the size of the component pieces
#This is a wreck right now because of the dependencies for dynamic sizing... need to figure out a better way to do this so I can order things appropriately eventually (tm)
$MainWindow.Size = New-Object System.Drawing.Size($WindowWidth,($SoftwareBuildBox.Size.Height+$CCILogo_Embed.Size.Height+$Instructions.Size.Height+$ConfirmButton.Size.Height+$TimeZoneDropDown.Size.Height+55))
$TabControl.Size = New-Object System.Drawing.Size(($WindowWidth-10),($SoftwareBuildBox.Size.Height+55))

$AdditionalSoftwareBox.Location = New-Object System.Drawing.Size(($TabControl.Size.Width/2),($SoftwareBuildBox.Location.Y))

$InfoBox.Location = New-Object System.Drawing.Size(5,($MainWindow.Size.Height-$InfoBox.Size.Height-5))
$ConfirmButton.Location = New-Object System.Drawing.Size(($MainWindow.Size.Width-$ConfirmButton.Size.Width-5),($MainWindow.Size.Height-$ConfirmButton.Size.Height-5))

Since so much of this is dynamically drawn/sized based on the content of other objects, this 
script block redefines the sizes for critical UI elements so that everything looks okay
in the end. 

#Set pages to enabled or not?
$DefaultPage.Enabled = $true
$CMPage.Enabled = $true
$PCMoverPage.Enabled = $true

Mostly for administrative purposes, this allows you to enable/disable the individual pages within the UI.

#Show the UI
$MainWindow.TopMost = $true
$MainWindow.Add_Shown({$MainWindow.Activate()})
[void] $MainWindow.ShowDialog()

Finally, this actually shows everything we've done to the technician.
