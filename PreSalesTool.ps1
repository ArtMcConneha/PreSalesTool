#region ScriptForm Designer

#region Constructor

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#endregion

#region Post-Constructor Custom Code

#endregion

#region Form Creation
#Warning: It is recommended that changes inside this region be handled using the ScriptForm Designer.
#When working with the ScriptForm designer this region and any changes within may be overwritten.
#~~< PresalesTools >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PresalesTools = New-Object System.Windows.Forms.Form
$PresalesTools.ClientSize = New-Object System.Drawing.Size(506, 384)
$PresalesTools.Text = "PreSales Tools"
#~~< grp_IISConfig >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$grp_IISConfig = New-Object System.Windows.Forms.GroupBox
$grp_IISConfig.Location = New-Object System.Drawing.Point(296, 184)
$grp_IISConfig.Size = New-Object System.Drawing.Size(128, 88)
$grp_IISConfig.TabIndex = 8
$grp_IISConfig.TabStop = $false
$grp_IISConfig.Text = "IIS"
#~~< btn_IISSiteBinding >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_IISSiteBinding = New-Object System.Windows.Forms.Button
$btn_IISSiteBinding.Location = New-Object System.Drawing.Point(8, 48)
$btn_IISSiteBinding.Size = New-Object System.Drawing.Size(112, 23)
$btn_IISSiteBinding.TabIndex = 1
$btn_IISSiteBinding.Text = "IIS Site Binding"
$btn_IISSiteBinding.UseVisualStyleBackColor = $true
$btn_IISSiteBinding.add_Click({Btn_IISSiteBindingClick($btn_IISSiteBinding)})
#~~< btn_IISAppPool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_IISAppPool = New-Object System.Windows.Forms.Button
$btn_IISAppPool.Location = New-Object System.Drawing.Point(8, 16)
$btn_IISAppPool.Size = New-Object System.Drawing.Size(112, 23)
$btn_IISAppPool.TabIndex = 0
$btn_IISAppPool.Text = "IIS App Pool"
$btn_IISAppPool.UseVisualStyleBackColor = $true
$btn_IISAppPool.add_Click({Btn_IISAppPoolClick($btn_IISAppPool)})
$grp_IISConfig.Controls.Add($btn_IISSiteBinding)
$grp_IISConfig.Controls.Add($btn_IISAppPool)
#~~< btn_ServicesStatus >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_ServicesStatus = New-Object System.Windows.Forms.GroupBox
$btn_ServicesStatus.Location = New-Object System.Drawing.Point(152, 24)
$btn_ServicesStatus.Size = New-Object System.Drawing.Size(128, 136)
$btn_ServicesStatus.TabIndex = 7
$btn_ServicesStatus.TabStop = $false
$btn_ServicesStatus.Text = "Services Status"
#~~< btn_ServiceUser >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_ServiceUser = New-Object System.Windows.Forms.Button
$btn_ServiceUser.Location = New-Object System.Drawing.Point(8, 64)
$btn_ServiceUser.Size = New-Object System.Drawing.Size(112, 23)
$btn_ServiceUser.TabIndex = 1
$btn_ServiceUser.Text = "Services User"
$btn_ServiceUser.UseVisualStyleBackColor = $true
$btn_ServiceUser.add_Click({Btn_ServiceUserClick($btn_ServiceUser)})
#~~< btn_Services >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_Services = New-Object System.Windows.Forms.Button
$btn_Services.Location = New-Object System.Drawing.Point(8, 32)
$btn_Services.Size = New-Object System.Drawing.Size(112, 23)
$btn_Services.TabIndex = 0
$btn_Services.Text = "Services Status"
$btn_Services.UseVisualStyleBackColor = $true
$btn_Services.add_Click({Btn_ServicesClick($btn_Services)})
$btn_ServicesStatus.Controls.Add($btn_ServiceUser)
$btn_ServicesStatus.Controls.Add($btn_Services)
#~~< grp_IPAddressing >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$grp_IPAddressing = New-Object System.Windows.Forms.GroupBox
$grp_IPAddressing.Location = New-Object System.Drawing.Point(152, 176)
$grp_IPAddressing.Size = New-Object System.Drawing.Size(128, 96)
$grp_IPAddressing.TabIndex = 6
$grp_IPAddressing.TabStop = $false
$grp_IPAddressing.Text = "IP Addressing"
#~~< btn_EditHostFile >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_EditHostFile = New-Object System.Windows.Forms.Button
$btn_EditHostFile.Location = New-Object System.Drawing.Point(8, 56)
$btn_EditHostFile.Size = New-Object System.Drawing.Size(112, 23)
$btn_EditHostFile.TabIndex = 1
$btn_EditHostFile.Text = "Edit HOSTS File"
$btn_EditHostFile.UseVisualStyleBackColor = $true
$btn_EditHostFile.add_Click({Btn_EditHostFileClick($btn_EditHostFile)})
#~~< btn_IPAddressing >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_IPAddressing = New-Object System.Windows.Forms.Button
$btn_IPAddressing.Location = New-Object System.Drawing.Point(8, 24)
$btn_IPAddressing.Size = New-Object System.Drawing.Size(112, 23)
$btn_IPAddressing.TabIndex = 0
$btn_IPAddressing.Text = "IP Address"
$btn_IPAddressing.UseVisualStyleBackColor = $true
$btn_IPAddressing.add_Click({Btn_IPAddressingClick($btn_IPAddressing)})
$grp_IPAddressing.Controls.Add($btn_EditHostFile)
$grp_IPAddressing.Controls.Add($btn_IPAddressing)
#~~< grp_GroupMembershipButtons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$grp_GroupMembershipButtons = New-Object System.Windows.Forms.GroupBox
$grp_GroupMembershipButtons.Location = New-Object System.Drawing.Point(8, 176)
$grp_GroupMembershipButtons.Size = New-Object System.Drawing.Size(128, 96)
$grp_GroupMembershipButtons.TabIndex = 5
$grp_GroupMembershipButtons.TabStop = $false
$grp_GroupMembershipButtons.Text = "Group Members"
#~~< btn_IISIUSRMembers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_IISIUSRMembers = New-Object System.Windows.Forms.Button
$btn_IISIUSRMembers.Location = New-Object System.Drawing.Point(8, 56)
$btn_IISIUSRMembers.Size = New-Object System.Drawing.Size(104, 23)
$btn_IISIUSRMembers.TabIndex = 1
$btn_IISIUSRMembers.Text = "IIS_IUSR Group"
$btn_IISIUSRMembers.UseVisualStyleBackColor = $true
$btn_IISIUSRMembers.add_Click({Btn_IISIUSRMembersClick($btn_IISIUSRMembers)})
#~~< btn_AdminGroupMembers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_AdminGroupMembers = New-Object System.Windows.Forms.Button
$btn_AdminGroupMembers.Location = New-Object System.Drawing.Point(8, 24)
$btn_AdminGroupMembers.Size = New-Object System.Drawing.Size(104, 24)
$btn_AdminGroupMembers.TabIndex = 0
$btn_AdminGroupMembers.Text = "Admin Group"
$btn_AdminGroupMembers.UseVisualStyleBackColor = $true
$btn_AdminGroupMembers.add_Click({Btn_AdminGroupMembersClick($btn_AdminGroupMembers)})
$grp_GroupMembershipButtons.Controls.Add($btn_IISIUSRMembers)
$grp_GroupMembershipButtons.Controls.Add($btn_AdminGroupMembers)
#~~< grp_FirewallButtons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$grp_FirewallButtons = New-Object System.Windows.Forms.GroupBox
$grp_FirewallButtons.Location = New-Object System.Drawing.Point(8, 24)
$grp_FirewallButtons.Size = New-Object System.Drawing.Size(128, 136)
$grp_FirewallButtons.TabIndex = 4
$grp_FirewallButtons.TabStop = $false
$grp_FirewallButtons.Text = "Firewall Settings"
#~~< FirewallStatus >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FirewallStatus = New-Object System.Windows.Forms.Button
$FirewallStatus.Location = New-Object System.Drawing.Point(8, 96)
$FirewallStatus.Size = New-Object System.Drawing.Size(112, 23)
$FirewallStatus.TabIndex = 3
$FirewallStatus.Text = "Firewall Status"
$FirewallStatus.UseVisualStyleBackColor = $true
$FirewallStatus.add_Click({FirewallStatusClick($FirewallStatus)})
#~~< btn_FirewallOff >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_FirewallOff = New-Object System.Windows.Forms.Button
$btn_FirewallOff.Location = New-Object System.Drawing.Point(8, 64)
$btn_FirewallOff.Size = New-Object System.Drawing.Size(112, 23)
$btn_FirewallOff.TabIndex = 2
$btn_FirewallOff.Text = "Firewall Off"
$btn_FirewallOff.UseVisualStyleBackColor = $true
$btn_FirewallOff.add_Click({Btn_FirewallOffClick($btn_FirewallOff)})
#~~< btn_FirewallOn >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_FirewallOn = New-Object System.Windows.Forms.Button
$btn_FirewallOn.Location = New-Object System.Drawing.Point(8, 32)
$btn_FirewallOn.Size = New-Object System.Drawing.Size(112, 23)
$btn_FirewallOn.TabIndex = 1
$btn_FirewallOn.Text = "Firewall On"
$btn_FirewallOn.UseVisualStyleBackColor = $true
$btn_FirewallOn.add_Click({Btn_FirewallOnClick($btn_FirewallOn)})
$grp_FirewallButtons.Controls.Add($FirewallStatus)
$grp_FirewallButtons.Controls.Add($btn_FirewallOff)
$grp_FirewallButtons.Controls.Add($btn_FirewallOn)
#~~< grp_Testing >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$grp_Testing = New-Object System.Windows.Forms.GroupBox
$grp_Testing.Location = New-Object System.Drawing.Point(296, 24)
$grp_Testing.Size = New-Object System.Drawing.Size(128, 136)
$grp_Testing.TabIndex = 9
$grp_Testing.TabStop = $false
$grp_Testing.Text = "Testing"
#~~< btn_TestRSPortCSV >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_TestRSPortCSV = New-Object System.Windows.Forms.Button
$btn_TestRSPortCSV.Location = New-Object System.Drawing.Point(8, 93)
$btn_TestRSPortCSV.Size = New-Object System.Drawing.Size(112, 23)
$btn_TestRSPortCSV.TabIndex = 2
$btn_TestRSPortCSV.Text = "Test RS Port - CSV"
$btn_TestRSPortCSV.UseVisualStyleBackColor = $true
$btn_TestRSPortCSV.add_Click({Btn_TestRSPortCSVClick($btn_TestRSPortCSV)})
#~~< btn_TestRSPort >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_TestRSPort = New-Object System.Windows.Forms.Button
$btn_TestRSPort.Location = New-Object System.Drawing.Point(8, 64)
$btn_TestRSPort.Size = New-Object System.Drawing.Size(112, 23)
$btn_TestRSPort.TabIndex = 1
$btn_TestRSPort.Text = "Test RS Port"
$btn_TestRSPort.UseVisualStyleBackColor = $true
$btn_TestRSPort.add_Click({Btn_TestRSPortClick($btn_TestRSPort)})
#~~< btn_TestMSPort >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_TestMSPort = New-Object System.Windows.Forms.Button
$btn_TestMSPort.Location = New-Object System.Drawing.Point(8, 32)
$btn_TestMSPort.Size = New-Object System.Drawing.Size(112, 23)
$btn_TestMSPort.TabIndex = 0
$btn_TestMSPort.Text = "Test MS Port"
$btn_TestMSPort.UseVisualStyleBackColor = $true
$btn_TestMSPort.add_Click({Btn_TestMSPortClick($btn_TestMSPort)})
$grp_Testing.Controls.Add($btn_TestRSPortCSV)
$grp_Testing.Controls.Add($btn_TestRSPort)
$grp_Testing.Controls.Add($btn_TestMSPort)
$PresalesTools.Controls.Add($grp_IISConfig)
$PresalesTools.Controls.Add($btn_ServicesStatus)
$PresalesTools.Controls.Add($grp_IPAddressing)
$PresalesTools.Controls.Add($grp_GroupMembershipButtons)
$PresalesTools.Controls.Add($grp_FirewallButtons)
$PresalesTools.Controls.Add($grp_Testing)

#endregion

#region Custom Code

#endregion

#region Event Loop

function Main{
	[System.Windows.Forms.Application]::EnableVisualStyles()
	[System.Windows.Forms.Application]::Run($PresalesTools)
	}

#endregion

#endregion

#region Event Handlers

function Btn_FirewallOffClick( $object ){

	set-NetFirewallProfile -Profile domain, Public, Private -Enabled False
	
}

function Btn_FirewallOnClick( $object ){

	set-NetFirewallProfile -Profile domain, Public, Private -Enabled True
	
}

function FirewallStatusClick( $object ){
		
	Get-NetFirewallProfile | Out-GridView
	
}

function Btn_IISIUSRMembersClick( $object ){
	
	Get-LocalGroupMember -group IIS_IUSRS | Out-GridView
	
}

function Btn_AdminGroupMembersClick( $object ){
	
	Get-LocalGroupMember -group administrators | Out-GridView
	
}

function Btn_ServicesClick( $object ){
	
	Get-Service "*milestone*" | Out-GridView
	
}

function Btn_IPAddressingClick( $object ){
	
	ipconfig /all | Out-GridView
	
}

function Btn_ServiceUserClick( $object ){
	
	Get-CimInstance -ClassName CIM_Service -filter "Name like 'Milestone%'" | Select-Object Name, StartMode, StartName | Out-GridView
	
}

function Btn_IISSiteBindingClick( $object ){
	
	Get-IISSite | Out-GridView
}

function Btn_IISAppPoolClick( $object ){
	
	Get-IISAppPool | Out-GridView
	
}

function Btn_TestRSPortClick( $object ){
	# add input methods for user entry, csv, and xml lookup
		
		Add-Type -AssemblyName Microsoft.VisualBasic
		$Rtitle = 'RS Hostname or IP'
		$Rmsg = 'Enter Hostname or IP of Recording Server'
		$RServer = [Microsoft.VisualBasic.Interaction]::InputBox($Rmsg, $Rtitle)
		
	
	Test-netconnection $Rserver -port 9001 | Out-GridView
	
}

function Btn_TestMSPortClick( $object ){
	
		
		Add-Type -AssemblyName Microsoft.VisualBasic
		$Mtitle = 'MS Hostname or IP'
		$Mmsg = 'Enter Hostname or IP of Management Server'
		$MServer = [Microsoft.VisualBasic.Interaction]::InputBox($Mmsg, $Mtitle)
		
	
	Test-netconnection $Mserver -port 9000 | Out-GridView
	
}

function Btn_EditHostFileClick( $object ){
		

	
	notepad C:\Windows\System32\drivers\etc\HOSTS
}

function Btn_TestRSPortCSVClick( $object ){
        
        function Get-FileName($initialDirectory)
        {
	        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	    
	        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	        $OpenFileDialog.initialDirectory = $initialDirectory
	        $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
	        $OpenFileDialog.ShowDialog() | Out-Null
	        $OpenFileDialog.filename
        }

        $path = Get-FileName

        if ($path -ne '')
        {
	        $Servers = Import-Csv -Path $path
	
            $jobs = New-Object System.Collections.Generic.List[object]
	        foreach ($ip in $Servers.ip)
	        {
                $jobs.Add((Invoke-Command -AsJob -ComputerName . -ArgumentList $ip, 9001 -ScriptBlock {
                    param([string]$ip, [int]$port)
                    Test-netconnection $ip -port $port
                }))
                $jobs.Add((Invoke-Command -AsJob -ComputerName . -ArgumentList $ip, 7563 -ScriptBlock {
                    param([string]$ip, [int]$port)
                    Test-netconnection $ip -port $port
                }))
	        }
            Receive-Job -Job $jobs -Wait -AutoRemoveJob | Out-GridView
        }
        else
        {
	        Write-Host -ForegroundColor Red "ERROR: Invalid File Path"
        }        
}

Main # This call must remain below all other event functions

#endregion
#region Script Settings
#<ScriptSettings xmlns="http://tempuri.org/ScriptSettings.xsd">
#  <ScriptPackager>
#    <process>powershell.exe</process>
#    <arguments />
#    <extractdir>%TEMP%</extractdir>
#    <files />
#    <usedefaulticon>true</usedefaulticon>
#    <showinsystray>false</showinsystray>
#    <altcreds>false</altcreds>
#    <efs>true</efs>
#    <ntfs>true</ntfs>
#    <local>false</local>
#    <abortonfail>true</abortonfail>
#    <product />
#    <version>1.0.0.1</version>
#    <versionstring />
#    <comments />
#    <company />
#    <includeinterpreter>false</includeinterpreter>
#    <forcecomregistration>false</forcecomregistration>
#    <consolemode>false</consolemode>
#    <EnableChangelog>false</EnableChangelog>
#    <AutoBackup>false</AutoBackup>
#    <snapinforce>false</snapinforce>
#    <snapinshowprogress>false</snapinshowprogress>
#    <snapinautoadd>2</snapinautoadd>
#    <snapinpermanentpath />
#    <cpumode>1</cpumode>
#    <hidepsconsole>false</hidepsconsole>
#  </ScriptPackager>
#</ScriptSettings>
#endregion

