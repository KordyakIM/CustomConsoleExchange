###########################################################
# Разработчик: Кордяк Иван Михайлович kordyakim@gmail.com #
###########################################################
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")  
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Tabl")
[void] [System.Windows.Forms.ComboBoxStyle]::DropDown
Import-Module ActiveDirectory
#Import-Module groupPolicy
#---------------------CheckBox---------------------#
$Font0 = New-Object Drawing.Font("Microsoft Sans Serif",8.25, [Drawing.FontStyle]::Bold)
#connected to Exchange
if (-not ($RemoteEx2013Session)){
$RemoteEx2013Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://server/PowerShell/" -Authentication Kerberos
Import-PSSession $RemoteEx2013Session -AllowClobber
}
$username = $Env:username

#-------------------------------------------------------#
#---------------------Создаём форму---------------------#

#$form = new-object System.Windows.Forms.form
$Exbody = New-Object System.Windows.Forms.Form;
$count = New-Object System.Windows.Forms.Label;
#$frmMain.icon =[system.drawing.icon]::ExtractAssociatedIcon("C:\Windows\System32\mmc.exe")   
$Exbody.ClientSize = New-Object System.Drawing.Size(1000, 700);    
$Exbody.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink;    
$Exbody.AutoScaleDimensions = New-Object System.Drawing.SizeF(200, 100);   
#$Exbody.FormBorderStyle = 'Fixed3D'
$Exbody.MaximizeBox = $true;    
$Exbody.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
$versionNow = "3.3"
$Exbody.Text = “Custom Console Exchange - v$versionNow”; 
$tmp = New-TemporaryFile
$tooltip = New-Object system.Windows.Forms.ToolTip
$Exbody.FormBorderStyle =[System.Windows.Forms.FormBorderStyle]::FixedSingle

#проверка версии программы
$pathvERSION = "\\SMBPath\vERSION.v"
$version = Get-Content $pathvERSION
if ($version -gt $versionNow)
{
	Start-Process -filepath "$pwd\Copy.exe" #-NoNewWindow
	$ExBody.Close()
}
#---------------------------------
#разделение на закладки----------------------------------------------------------------------
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 2
$System_Drawing_Point.Y = 189
$tabControl.Location = $System_Drawing_Point
$tabControl.Name = "tabControl1"
$tabControl.SelectedIndex = 0
$tabControl.ShowToolTips = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 700
#$System_Drawing_Size.Width = 383
$tabControl.Size = $System_Drawing_Size
$tabControl.TabIndex = 4
$tabControl.Dock = "fill"
$tabControl.Sty
$tabControl.add_MouseClick({
})
$ExBody.Controls.Add($TabControl);

$tabControlP1 = New-Object System.Windows.Forms.TabPage
$tabControlP1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabControlP1.Location = $System_Drawing_Point
$tabControlP1.Name = "tabControl"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabControlP1.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 205
$System_Drawing_Size.Width = 445
$tabControlP1.Size = $System_Drawing_Size
$tabControlP1.TabIndex = 0
$tabControlP1.Text = "General"
$tabControlP1.UseVisualStyleBackColor = $True
$tabControl.Controls.Add($tabControlP1)

$tabControlP2 = New-Object System.Windows.Forms.TabPage
$tabControlP2.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabControlP2.Location = $System_Drawing_Point
$tabControlP2.Name = "tabControl"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabControlP2.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 205
$System_Drawing_Size.Width = 445
$tabControlP2.Size = $System_Drawing_Size
$tabControlP2.TabIndex = 0
$tabControlP2.Text = "Message Tracking Log"
$tabControlP2.UseVisualStyleBackColor = $True
$tabControl.Controls.Add($tabControlP2)


. "$PSScriptRoot\MessageTrackingLog.ps1"
. "$PSScriptRoot\ADDPermission.ps1"
#. "$PSScriptRoot\ConnectAccountToSkype.ps1"
. "$PSScriptRoot\InfUser.ps1"

#окно вывода Пользователей
$ListMB = New-Object System.Windows.Forms.ListView
$ListMB.dock = "Fill"
$ListMB.Height = 200
$ListMB.View = "Details"
$ListMB.MultiSelect = $true
$ListMB.FullRowSelect = $True
$ListMB.AutoSize = $true
$ListMB.LabelEdit = $True
$ListMB.AllowColumnReorder = $True
$ListMB.GridLines = $true
$ListMB.Columns.Add("ID", 30)
$ListMB.Columns.Add("DisplayName", 250)
$ListMB.Columns.Add("Samaccountname", 100)
$ListMB.Columns.Add("PrimarySmtpAddress", 150)
$ListMB.Columns.Add("RecipientTypeDetails", 120)
$ListMB.Columns.Add("Identity", 200)
$ListMB.Columns.Add("ItemCount", 100)
$ListMB.Columns.Add("Totalitemsize", 150)
$ListMB.Columns.Add("Totaldeleteditemsize", 150)
$ListMB.Columns.Add("TotalmailboxsizeMR", 150)
$ListMB.Columns.Add("ProhibitSendQuota", 150)
$ListMB.Columns.Add("mDBUseDefaults", 180)
$ListMB.Columns.Add("TargetDataBase", 150)
$ListMB.Columns.Add("SourceDataBase", 150)
$ListMB.Columns.Add("StatusDetailMR", 150)
$ListMB.Columns.Add("PercentCompleteMR", 180)
$ListMB.Columns.Add("StartTimestampMR", 180)
$ListMB.Columns.Add("CompletionTimestampMR", 180)
#При двойном клике список доступа к ПЯ
$ListMB.add_DoubleClick({
 if ($listMB.SelectedItems.Count -cge 1)
 {
Get-MailboxPermission $ListMB.Items[$ListMB.FocusedItem.Index].Subitems[1].Text | Out-GridView -Title "Список пользователей, имеющих доступ к почтовому ящику"
 }
})
$ListMB.add_ItemSelectionChanged({
	if($listMB.SelectedItems.Count -cge 1){
		$menucopyMB.Enabled = $true
		$menumInfUser.Enabled = $true
		$menususpendMB.Enabled = $true
		$menuresumeMB.Enabled = $true
		$menumDBUseDefaultsFalseMB.enabled = $true
		$menumDBUseDefaultsTrueMB.enabled = $true
		$menumDBAccessFromInternet.Enabled = $true
		$menumDBADDPermissionFullAccess.Enabled = $true
		$menumDBADPermissionSend_AS.Enabled = $true
		$menumDBDeletePermission.Enabled = $true
		$menumDBADDPermissionFullAccess.Enabled = $true
	    $menumDBADPermissionSend_AS.Enabled = $true
	    $menumDBDeletePermission.Enabled = $true
	    $menumAccessCalendar.Enabled = $true
	    $menumDeleteCalendar.Enabled = $true
		$menumALLAccesscalendar.Enabled = $true
	
	if ($ListMB.Items[$ListMB.FocusedItem.Index].forecolor -eq "Red")
	{
	  $menumDBAccessFromInternet.Text = "Удаление пользователя из группы предоставления доступа из Интернета"
	}
    else
	{
	$menumDBAccessFromInternet.Text = "Добавление пользователя в группу предоставления доступа из Интернета"
	
	}
#		if ($ListMB.Items[$ListMB.FocusedItem.Index].Subitems[4].Text -eq "SharedMailbox")
#			{
#            $menumDBADDPermissionFullAccess.Enabled = $true
#			 $menumDBADPermissionSend_AS.Enabled = $true
#			 $menumDBDeletePermission.Enabled = $true
#			}
#			else
#			{
#		   
#			 $menumDBADDPermissionFullAccess.Enabled = $false
#			 $menumDBADPermissionSend_AS.Enabled = $false
#			 $menumDBDeletePermission.Enabled = $false
#			 }
		
			
		
#			if ($ListMB.Items[$ListMB.FocusedItem.Index].forecolor -eq "red")
#			{
#		        $menumDBAccessFromInternet.Enabled =$false
#			    $menumDBDeleteAccessFromInternet.Enabled = $true 
#			}
#			else
#			{
#			
#				$menumDBAccessFromInternet.Enabled = $true
#			    $menumDBDeleteAccessFromInternet.Enabled = $false 
#			}
		
	}else{
		$menucopyMB.Enabled = $false
		$menususpendMB.Enabled = $false
		$menuresumeMB.Enabled = $false
		$menumDBUseDefaultsFalseMB.enabled = $false
		$menumDBUseDefaultsTrueMB.enabled = $false
		$menumDBAccessFromInternet.Enabled = $false
		$menumDBADDPermissionFullAccess.Enabled = $false
		$menumDBADPermissionSend_AS.Enabled = $false
		$menumDBDeletePermission.Enabled = $false
		$menumInfUser.Enabled = $false
		$menumDBADDPermissionFullAccess.Enabled = $false
		$menumDBADPermissionSend_AS.Enabled = $false
		$menumDBDeletePermission.Enabled = $false
		$menumAccessCalendar.Enabled = $false
	    $menumDeleteCalendar.Enabled = $false
		$menumALLAccesscalendar.Enabled = $false
	}
	if($list.CheckedItems.Count -eq 1){
		$btn5.Enabled = $true
	}else{
		$btn5.Enabled = $false
	}
})
$ListMB.add_ColumnClick({
	if ($ListMB.Items.Count -gt 1){
		SortListWithID $_.Column
	}
})
$ListMB.add_KeyDown({
	param($sender, $e)
	if ($_.KeyCode -eq "C" -and $e.Control){
		Set-ClipBoard
	}
	if ($_.keycode -eq "A" -and $e.Control){
		foreach ($ListItem in $listMB.Items){
		    $ListItem.selected = $true
		}
	}
})
$ListMB.add_MouseDoubleClick({
#MassegeTracking
})
$tabControlP1.Controls.Add($ListMB);
#---------------------------------------------------------------



#описание программ
			$MenuBar = New-Object System.Windows.Forms.MenuStrip
			#$MenuBar.Size = New-Object System.Drawing.Size(20,20);
			#$MenuBar.Location = New-Object System.Drawing.Point(400,620);
			$MenuBar.Dock = "bottom"
			$Exbody.Controls.Add($MenuBar);
			$UserMenu = New-Object System.Windows.Forms.ToolStripMenuItem
			$UserMenu.Text = "Информация"
			$UserMenu.Name = "openToolStripMenuItem"
			$UserMenu.Alignment = 'Right'
			$MenuBar.Items.Add($UserMenu)
			$UserMenu.add_Click({$output = [System.Windows.Forms.MessageBox]::Show("Кастомная консоль для сервиса Exchange.

	Разработчик:
	-  Кордяк Иван		kordyakim@gmail.com
	Есть инструкция по программе, хотите просмотреть?","Информация",4)
if ($OUTPUT -eq "YES")
{
#функция Обзора
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
	if($FolderBrowser.ShowDialog() -eq "OK") {
		Copy-Item "$PWD\ExConsole.docx" -Destination $FolderBrowser.SelectedPath
		$dir = $FolderBrowser.SelectedPath
		Start-Process -filepath "$dir\ExConsole.docx"
		}
}
			})
			#-----------------------------------------------------  

#строка запроса
$SearchOnType = $false
$textBox = New-Object System.Windows.Forms.TextBox;    
$textBox.dock = "top"
$textBox.Location = New-Object System.Drawing.Point(10, 10);    
$textBox.Size = New-Object System.Drawing.Size(365, 10);    
$textBox.Name = "textBox0";      
#$textBox.Text = "user";
$textBoxname = 'Введите пожалуйста имя Пользователя!'
$textBox.ForeColor = 'LightGray'
$textBox.Text = $textBoxname
$textBox_AddGM = 0;
$textBox.add_Click({
	if($textBox.Text -eq $textBoxname)
    {
        #Clear the text
        $textBox.Text = ""
        $textBox.ForeColor = 'WindowText'
    }
	if($textBox.Text -eq $textBox.Tag)
    {
        #Clear the text
        $textBox.Text = ""
        $textBox.ForeColor = 'WindowText'
    }
	})
$textBox.add_KeyPress({
if($textBox.Visible -and $textBox.Tag -eq $null)
    {
        #Initialize the watermark and save it in the Tag property
        $textBox.Tag = $textBox.Text;
        $textBox.ForeColor = 'LightGray'
        #If we have focus then clear out the text
        if($textBox.Focused)
        {
            $textBox.Text = ""
            $textBox.ForeColor = 'WindowText'
        }
    }
})
$textBox.add_Leave({
if($textBox.Text -eq "")
    {
        #Display the watermark
        $textBox.Text = $textBoxname
        $textBox.ForeColor = 'LightGray'
    }
	if($textBox.Text -eq "")
    {
        #Display the watermark
        $textBox.Text = $textBox.Tag
        $textBox.ForeColor = 'LightGray'
    }
		})
$tabControlP1.Controls.Add($textBox);
#------------------------------------------------------------------

#Место для выгрузки информации по пользователю----------------
$txtInfo = New-Object System.Windows.Forms.TextBox
$txtInfo.dock = "Top"
$txtInfo.Height = 180
$txtInfo.ReadOnly = $true
$txtInfo.Multiline = $True
$txtInfo.ScrollBars = "Vertical"
$tabControlP1.Controls.add($txtInfo)
#-----------------------------------

					#создание запроса--------------------------------------
					function Fill-List ($Mask = "*") {
						$ListMB.Items.Clear()
						$s = $textBox.Text + "*"
						$mail =  [scriptblock]::create("alias -like `"$s`"")
						$mailn =  [scriptblock]::create("displayname -like `"$s`"")
					    $str = Get-mailbox -Filter $mail | select PrimarySmtpAddress,Alias,DisplayName,ProhibitSendQuota,Database,Identity,RecipientTypeDetails,SamAccountName
							if(!$str){
								$str = Get-mailbox -Filter $mailn | select PrimarySmtpAddress,Alias,DisplayName,ProhibitSendQuota,Database,Identity,RecipientTypeDetails,SamAccountName
							}
							    if(!$str){
									$I = $ListMB.Items.Add("1") 
									$I.SubItems.Add("Пользователь не найден...") | Out-Null
								}else{
									$id=1
								foreach ( $item in $str ) 
							    	{
									$mDBUseDefaults=(Get-ADUser $item.SamAccountName -Properties *).mDBUseDefaults
									$str1 = Get-MailboxStatistics $item.PrimarySmtpAddress | select itemcount,totalitemsize,totaldeleteditemsize
									$saved=$global:ErrorActionPreference
    								$global:ErrorActionPreference='stop'
									try{
									$moves=Get-MoveRequest -identity $item.PrimarySmtpAddress | Get-MoveRequestStatistics | Select-Object statusdetail,percentcomplete,totalmailboxsize,StartTimestamp,CompletionTimestamp
									}Catch{
								         Write-Warning $_
								    }
								    Finally{
								        $global:ErrorActionPreference=$saved
									}
									$sizeIS=$str1.totalitemsize -replace ".*[(]",""
									$sizeIS1="$sizeIS" -replace "[by].*",""
									$sizeIDS=$str1.totaldeleteditemsize -replace ".*[(]",""
									$sizeIDS1="$sizeIDS" -replace "[by].*",""
									$sizeMB=$str1.totalmailboxsize -replace ".*[(]",""
									$sizeMB1="$sizeMB" -replace "[by].*",""
									$ProhibitSendQuota=$item.ProhibitSendQuota -replace ".*[(]",""
									$ProhibitSendQuota1=$ProhibitSendQuota -replace "[by].*",""
									$ser=$moves.StatusDetail
									$moveSDB = $move.SourceDatabase
									$StartTS=$move.StartTimestamp
									$CompletionTS=$move.CompletionTimestamp
							    	#Добавляем элемент в список
									$I = $ListMB.Items.Add($id)
									if($item.DisplayName){$I.SubItems.Add($item.DisplayName)}else{$I.SubItems.Add("null")}
									if($item.samaccountname){$I.SubItems.Add($item.samaccountname)}else{$I.SubItems.Add("null")}
									if($item.PrimarySmtpAddress){$I.SubItems.Add($item.PrimarySmtpAddress)}else{$I.SubItems.Add("null")}
									if($item.RecipientTypeDetails){$I.SubItems.Add($item.RecipientTypeDetails)}else{$I.SubItems.Add("null")}
									if($item.Identity){$I.SubItems.Add($item.Identity)}else{$I.SubItems.Add("null")}
									if($str1.itemcount){$I.SubItems.Add($str1.itemcount)}else{$I.SubItems.Add("0")}
									if("$sizeIS1"){$I.SubItems.Add("$sizeIS1")}else{$I.SubItems.Add("0")}
									if("$sizeIDS1"){$I.SubItems.Add("$sizeIDS1")}else{$I.SubItems.Add("0")}
									if("$sizeMB1"){$I.SubItems.Add("$sizeMB1")}else{$I.SubItems.Add("0")}
									if("$ProhibitSendQuota1"){$I.SubItems.Add("$ProhibitSendQuota1")}else{$I.SubItems.Add("null")}
										if("$mDBUseDefaults"){$I.SubItems.Add("$mDBUseDefaults")}else{$I.SubItems.Add("null")}
									if($item.Database){$I.SubItems.Add($item.Database)}else{$I.SubItems.Add("null")}
									if("$moveSDB"){$I.SubItems.Add("$moveSDB")}else{$I.SubItems.Add("null")}
									if($moves.StatusDetail){$I.SubItems.Add($moves.StatusDetail)}else{$I.SubItems.Add("null")}
									if($moves.percentcomplete){$I.SubItems.Add($moves.percentcomplete)}else{$I.SubItems.Add("0")}
									if("$StartTS"){$I.SubItems.Add("$StartTS")}else{$I.SubItems.Add("0")}
									if("$CompletionTS"){$I.SubItems.Add("$CompletionTS")}else{$I.SubItems.Add("0")}
										if((((Get-ADUser $item.SamAccountName -Properties MemberOf).memberof | get-adgroup).SID.value -like "SID group" ) `
										-or ( ((Get-ADUser $item.SamAccountName -Properties MemberOf).memberof | get-adgroup).sid.value -like "SID group"))
										{
												$ListMB.Items[$id-1].forecolor = "red"
											}
#										if(((Get-ADUser $item.SamAccountName -Properties memberof).memberof -like $FullAccessFromInternetZF)-or((Get-ADUser $item.SamAccountName -Properties memberof).memberof -like $FullAccessFromInternetGO))
#										{
#											$ListMB.Items[$id-1].forecolor = "red"
#										}
					### сравнение с квотой						
					
						     if($ListMB.Items[$id-1].Subitems[12].Text  -eq $item.Database )
							 {
							   [string]$stringtext = Get-MailboxDatabase $item.Database -Status | select ProhibitSendQuota
				
									$stringlist = $stringtext.Split("(",2) -replace "[bytes]","" -replace "[)]","" -replace ",","" -replace "[}]",""
									Write-Host ($stringlist)
									$kvota = [int64]$stringlist[1]
		                            $total =  [int64]($ListMB.Items[$id-1].SubItems[7].Text -replace ",","")

									if ($total -ge $kvota)
										{
										  $ListMB.Items[$id-1].backcolor = "gray"
										}
							}
						$id++
						}
									
									}
								}
						
					if ($SearchOnType)
					{
					    #Добавляем обработчик на событие TextChanged, который выполняет функцию Fill-List
					    $textBox.add_TextChanged({Fill-List ("*" + $textBox.Text + "*")})
					}
					else #Ищем только при нажатии Enter
					{
					    #Скриптблок (кусок исполняемого кода) который будет выполнен при нажатии клавиши в поле поиска
					    $SB_KeyPress = {
					        #Если была нажата клавиша Enter (13) то...
					        if (13 -eq $_.keychar)
					        {
					            #Вызываем функцию Fill-List
					            Fill-List ("*" + $textBox.Text + "*")
					        }
					    }
					    #Добавляем обработчик на событие KeyPress, указав в качестве выполняемого кода $SB_KeyPress
					    $textBox.add_KeyPress($SB_KeyPress)
					}
					#-----------------------------------------------------------



#окно запроса баз данных---------------------------------------
$List = New-Object System.Windows.Forms.ListView
$List.dock = "top"
$List.Height = 200
#$List.Width = 200
$List.View = "Details"
$List.MultiSelect = $true
$List.FullRowSelect = $True
#$List.StateImageList = $true
$List.AutoSize = $true
$List.LabelEdit = $True
$List.AllowColumnReorder = $True
$List.CheckBoxes = $true
$List.Scrollable = $true
$List.GridLines = $true
#$List.HoverSelection = $true #тащится выделение за мышкой
$List.Columns.Add("Name", -1)
$List.Columns.Add("Description",100)
$List.Columns.Add("MountedOnServer", -1)
$List.Columns.Add("Mounted", 60)
$List.Columns.Add("Servers",-1)
$List.Columns.Add("DatabaseSize", -1)
$List.Columns.Add("AvailableNewMailboxSpace", -1)
$List.Columns.Add("AdminDisplayVersion", -1)
$List.Columns.Add("ProhibitSendQuota", -1)
$List.Columns.Add("LogFileSize", -1)
$List.add_ItemChecked({
		if($list.CheckedIndices.Count -ige 1){
			$btn1.Enabled = $true
			$btn2.Enabled = $true
			$btn3.Enabled = $true
			$btn4.Enabled = $true
			$btn6.Enabled = $true
			$btn5.Enabled = $true
			$btn8.Enabled = $true
			$menucopy.Enabled = $true
		}else{
			$btn1.Enabled = $false
			$btn2.Enabled = $false
			$btn3.Enabled = $false
			$btn4.Enabled = $false
			$btn5.Enabled = $false
			$btn6.Enabled = $false
			$btn8.Enabled = $false
			$menucopy.Enabled = $false
		}
		})

						
$List.add_ItemSelectionChanged({
	$List.add_ItemMouseHover({
		if($list.SelectedItems.Count -cge 1){
			$menucopy.Enabled = $true
		}
		if($list.CheckedItems.Count -cge 1){
			$menucopy.Enabled = $true
		}
	})
		if($list.SelectedItems.Count -cge 1){
			$menucopy.Enabled = $true
		}
		if($list.CheckedItems.Count -cge 1){
			$menucopy.Enabled = $true
		}
})


$List.add_ColumnClick({
	if ($List.Items.Count -gt 1){
	SortListTwoviewDB $_.Column
	}
})
$List.add_KeyDown({
	param($sender, $e)
	if ($_.KeyCode -eq "C" -and $e.Control){
	Set-ClipBoard
	}
	if ($_.keycode -eq "A" -and $e.Control){
		foreach ($ListItem in $list.Items){
		    $ListItem.selected = $true
		}
	}
})
$tabControlP1.Controls.add($List)
#----------------------------------------
				#Сортировать только по возрастания---------------------------------------------------------------------------------------------
				$LastColumnAscendingOne = $false # отслеживает направление из последних сортировки этого столбца (tracks the direction of the last sort of this column)
				$LastColumnClickedTwo = 0 # треки номер последнего столбца, который был выбран (tracks the last column number that was clicked)
				$LastColumnAscendingTwo = $false # отслеживает направление из последних сортировки этого столбца (tracks the direction of the last sort of this column)
				#Сортировать только по возрастания---------------------------------------------------------------------------------------------
				function SortListOneviewDB {
				param([parameter(Position=0)][UInt32]$Column)
				$Numeric = $true # определить, как сортировать (determine how to sort)
				foreach($ListItem in $List.Items)
				{
				    #если все элементы являются числовыми, могут использовать числовую сортировку (if all items are numeric, can use a numeric sort)
				    if($Numeric -ne $false){ # ничто не может установить значение True, поэтому не процесс излишне (nothing can set this back to true, so don't process unnecessarily)
				        try
				        {
				            $Test = [Double]$ListItem.SubItems[[int]$Column].Text
				        }
				 		catch
				 		{
				            $Numeric = $false #найден нечисловых элементов, так что сортировка будет происходить в виде строки (a non-numeric item was found, so sort will occur as a string)
				        }
				    }
				    $ListItems += ,@($ListItem.SubItems[[int]$Column].Text,$ListItem)
				}
				#создать выражение, которое будет вычисляться для сортировки (create the expression that will be evaluated for sorting)
				$EvalExpression = {
				    if($Numeric)
				    { return [Double]$_[0] }
				    else
				    { return [String]$_[0] }
				}
				#вся информация собрана; выполнения сортировки (all information is gathered; perform the sort)
				$ListItems = $ListItems | Sort-Object -Property @{Expression=$EvalExpression; Descending=$Script:LastColumnAscendingOne}
				#список отсортирован, вывести в list (the list is sorted; display it in the listview)
				$List.BeginUpdate()
				$List.Items.Clear()
				foreach($ListItem in $ListItems)
				{
				    $List.Items.Add($ListItem[1])
				}
				$List.EndUpdate()
				}
				#------------------------------------------------------------------------------------------------------------------------------
				#Сортироватm в два вида (возрастания и убывание)-------------------------------------------------------------------------------
				function SortListTwoviewDB {
				 param([parameter(Position=0)][UInt32]$Column)
				$Numeric = $true # определить, как сортировать (determine how to sort)
				#если пользователь нажал тот же столбец, который был выбран последний раз, его обратный порядок сортировки. в противном случае, сброс для нормальной сортировки по возрастанию
				#if the user clicked the same column that was clicked last time, reverse its sort order. otherwise, reset for normal ascending sort
				if($Script:LastColumnClickedTwo -eq $Column-or$Script:LastColumnClickedOne -eq $Column){
				    $Script:LastColumnAscendingTwo = -not $Script:LastColumnAscendingTwo
				}else{
				    $Script:LastColumnAscendingTwo = $true
				}
				$Script:LastColumnClickedTwo = $Column
				#трехмерный массив; колонке 1 индексы других столбцов, столбец 2 является значением, которое будет отсортирован, и колонка 3 является System.Windows.Forms.ListViewItem object
				#three-dimensional array; column 1 indexes the other columns, column 2 is the value to be sorted on, and column 3 is the System.Windows.Forms.ListViewItem object
				$ListItems = @(@(@()))
				foreach($ListItem in $List.Items)
				{
				    #если все элементы являются числовыми, могут использовать числовую сортировку (if all items are numeric, can use a numeric sort)
				    if($Numeric -ne $false) #ничто не может установить значение True, поэтому не процесс излишне (nothing can set this back to true, so don't process unnecessarily)
				    {
				        try
				        {
				            $Test = [Double]$ListItem.SubItems[[int]$Column].Text
				        }
				        catch
				        {
				            $Numeric = $false #найден нечисловых элементов, так что сортировка будет происходить в виде строки (a non-numeric item was found, so sort will occur as a string)
						}
				    }
				    $ListItems += ,@($ListItem.SubItems[[int]$Column].Text,$ListItem)
				}
				#создать выражение, которое будет вычисляться для сортировки (create the expression that will be evaluated for sorting)
				$EvalExpression = {
				    if($Numeric)
				    { return [double]$_[0] } #{ return [double]$_[0] } #[double]$_[0] }
				    else
				    { return [String]$_[0] }
				}
				#вся информация собрана; выполнения сортировки (all information is gathered; perform the sort)
				$ListItems = $ListItems | Sort-Object -Property @{Expression=$EvalExpression; Ascending=$Script:LastColumnAscendingTwo}
				#список отсортирован, вывести в list (the list is sorted; display it in the listview)
				$List.BeginUpdate()
				$List.Items.Clear()
				foreach($ListItem in $ListItems)
				{
				    $List.Items.Add($ListItem[1])
				}
				$List.EndUpdate()
				}
				#-------------------------------------------------------------------------------------------------------------------------------
				
					#заполнение таблицы
					function DB {
					$DBs = @()
					$DBs = Get-MailboxDatabase N* -Status | select name,DatabaseSize,MountedOnServer,availablenewmailboxspace,ProhibitSendQuota,Servers,AdminDisplayVersion,mounted,LogFileSize,description
					foreach ($DB in $DBs)
					{
						$DBse= $DB.Servers
						$DBm= $DB.mounted
						if($DB.name){$item = $List.Items.Add($DB.name)}else{$item = $List.Items.Add("null")}
						if($DB.description){$item.SubItems.Add($DB.description)}else{$item.SubItems.Add("null")}
						if($DB.MountedOnServer){$item.SubItems.Add($DB.MountedOnServer)}else{$item.SubItems.Add("null")}
						if("$DBm"){$item.SubItems.Add("$DBm")}else{$item.SubItems.Add("null")}
						if("$DBse"){$item.SubItems.Add("$DBse")}else{$item.SubItems.Add("null")}
						if($DB.DatabaseSize){$item.SubItems.Add($DB.DatabaseSize)}else{$item.SubItems.Add("null")}
						if($DB.availablenewmailboxspace){$item.SubItems.Add($DB.availablenewmailboxspace)}else{$item.SubItems.Add("null")}
						if($DB.AdminDisplayVersion){$item.SubItems.Add($DB.AdminDisplayVersion)}else{$item.SubItems.Add("null")}
						if($DB.ProhibitSendQuota){$item.SubItems.Add($DB.ProhibitSendQuota)}else{$item.SubItems.Add("null")}
						if($DB.LogFileSize){$item.SubItems.Add($DB.LogFileSize)}else{$item.SubItems.Add("0")}
						
					}
					}
					DB
					SortListTwoviewDB
					for ($k =0; $k -lt $list.items.count; $k++)
					{
					  if ($list.items[$k].text -eq "NRDB04")
					  {
					    $List.Items[$k].Backcolor = "green"
					  }
					}
					
		
					#----------------------------------------------------------------

function info1 {
$nDB = $List.SelectedItems.text
Write-Host $nDB 
$r = Get-MailboxDatabase "$nDB" | select name
#Заполняем txtInfo данными из объекта-------------------------------------------------------------
    $txtInfo.Text =
@"
$($r)`r
"@
}



#group for combobox
$GroupOU = New-Object System.Windows.Forms.GroupBox
$GroupOU.dock = "top"
$GroupOU.Height = 60
$tabControlP1.Controls.Add($GroupOU);

$labelOU = New-Object System.Windows.Forms.Label
$labelOU.Location = New-Object System.Drawing.Point(10,10) 
$labelOU.Size = New-Object System.Drawing.Size(400,20)
$labelOU.Text = "Выберите OU:"
$groupOU.Controls.Add($labelOU)

$comboBoxOU = New-Object System.Windows.Forms.comboBox
$comboBoxOU.Location = New-Object System.Drawing.Point(10,30) 
$comboBoxOU.Size = New-Object System.Drawing.Size(300,8)
$comboBoxOU.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboBoxOU.BackColor = "white"
$comboBoxOU.Visible = $true

Function combobox (){
	foreach($OU in $OUs)
		{
  		$comboBoxOU.Items.add($OU) | Out-Null
		}
	}
	
$groupOU.Controls.Add($comboBoxOU)
	

#количество мигрированных баз
$checkboxMove = New-Object System.Windows.Forms.CheckBox
$checkboxMove.Location = New-Object System.Drawing.Point(790,30) 
$checkboxMove.Size = New-Object System.Drawing.Size(300,20)
$checkboxMove.Text = "Move MB"
$groupOU.Controls.Add($checkboxMove)

#checkbox все почтовые базы
$checkboxALLDB = New-Object System.Windows.Forms.CheckBox
$checkboxALLDB.Location = New-Object System.Drawing.Point(630,30) 
$checkboxALLDB.Size = New-Object System.Drawing.Size(300,20)
$checkboxALLDB.Text = "Select all mailbox database"
$groupOU.Controls.Add($checkboxALLDB)

#checkbox OU--------------------------------------------------------------------
$checkbox2 = New-Object System.Windows.Forms.CheckBox
$checkbox2.Location = New-Object System.Drawing.Point(500,30) 
$checkbox2.Size = New-Object System.Drawing.Size(300,20)
$checkbox2.Text = "Select OU-ZF"
$groupOU.Controls.Add($checkbox2)

$checkbox1 = New-Object System.Windows.Forms.CheckBox
$checkbox1.Location = New-Object System.Drawing.Point(330,30) 
$checkbox1.Size = New-Object System.Drawing.Size(300,20)
$checkbox1.Text = "Select OU-npr.nornick.ru"
$groupOU.Controls.Add($checkbox1)



$checkboxALLDB.add_checkedChanged({
   if ($checkboxALLDB.Checked)
   {
     for ($i=0; $i -lt $List.Items.Count; $i++)
	 {
	    $List.Items[$i].Checked = $true
	 }
	
   }
   
   if ($checkboxALLDB.Checked -eq $false)
   {
       for ($i=0; $i -lt $List.Items.Count; $i++)
	 {
	    $List.Items[$i].Checked = $false
	 }
   }
})

$checkbox1.add_CheckedChanged({
    if ($checkbox1.Checked)
	{
	  $comboBoxOU.Items.Clear()
	  $checkbox2.Checked = $false
	  $addressOU ="DC=domain,DC=ru"
	  
	  if($comboBoxOU.Disposing -contains $comboBoxOU.Items){$OUs = get-ADOrganizationalUnit -SearchBase $addressOU -Filter * -SearchScope OneLevel}
	  combobox
	
	}
	else
	{
	$comboBoxOU.Items.Clear()
	}
})

$checkbox2.add_CheckedChanged({
    if ($checkbox2.Checked)
	{
	  $comboBoxOU.Items.Clear()
	  $checkbox1.Checked = $false
	  $addressOU ="OU=domain,DC=com"
	  if($comboBoxOU.Disposing -contains $comboBoxOU.Items){$OUs = get-ADOrganizationalUnit -SearchBase $addressOU -Filter * -SearchScope OneLevel}
	  combobox

	}
	else
	{
	$comboBoxOU.Items.Clear()
	}
})


	
	
#group for button--------------------------------------------------------------------
$GroupInside = New-Object System.Windows.Forms.GroupBox
$GroupInside.dock = "top"
$GroupInside.Height = 40
$tabControlP1.Controls.Add($GroupInside);
#-----------------------------------------------------------------------------------------

####Completed - очистить запросы						
$btn8 = New-Object System.Windows.Forms.Button
$btn8.dock = "left"
$btn8.Height = 40
$btn8.Size = New-Object System.Drawing.Size(140, 25);    
$btn8.Name = "btn8";    
$btn8.Text = "Clean query MB"; 
$btn8.Enabled = $false;  
$btn8.add_click({
$txtInfo.Text="Подождите..." 
		Start-Transcript -Path $tmp.fullname
			$databases = $List.CheckedItems.text 
			foreach($database in $databases){
				Get-MoveRequest -TargetDatabase $database -MoveStatus completed | Remove-MoveRequest -Confirm:$false
			}
			count-move
		$log = Get-Content $tmp.fullname
		$txtInfo.Text=$log+"`n"+"`n"
		Stop-Transcript
		$txtInfo.AppendText("Готово!")
		
})
$GroupInside.Controls.Add($btn8)
$btn8.add_MouseHover({
	$tooltip.SetToolTip($btn8, "Очищает запросы выполненные при миграции почтовых ящиков.")
})


$btn1 = New-Object System.Windows.Forms.Button; 
$btn1.dock = "left"
$btn1.Height = 40  
$btn1.Size = New-Object System.Drawing.Size(100, 25);    
$btn1.Name = "btn1";    
$btn1.Text = "Count MB"; 
$btn1.Enabled = $false;    
$btn1.Add_Click({
			count-MB
	})
$btn1.add_MouseMove({
	$tooltip.SetToolTip($btn1, "Считает количество почтовых ящиков.")
})
$GroupInside.Controls.Add($btn1);

								function count-MB {
									$txtInfo.Text = "Подождите..."
											Start-Transcript -Path $tmp.fullname
									$count=0;
									$databases = $List.CheckedItems.text
                                if ($checkboxMove.checked)
								{
									foreach($database in $databases)
									{
									$moves = Get-MoveRequestStatistics -MoveRequestQueue $database | Select-Object displayname,statusdetail,totalmailboxsize,percentcomplete,TargetDatabase,Alias,SourceDatabase,StartTimestamp,CompletionTimestamp,Identity,RecipientTypeDetails
									$allCountDBCompleted = 0;
									$allCountDBCompleted =(Get-MoveRequest -TargetDatabase $database | measure).count; 
									$CountDBCompleted = (Get-MoveRequest -TargetDatabase $database -MoveStatus completed | measure ).count; 
									$allCountDBCompleted -=(Get-MoveRequest -TargetDatabase $database -MoveStatus completed | measure ).count; 
									$txtInfo.AppendText("Completed "+$database+"="+$CountDBCompleted+"`n")
									$txtInfo.AppendText("NotCompleted "+$database+"="+$allCountDBCompleted+"`n"+"`n")
								
									}
									$checkboxMove.Checked = $false
									$log = Get-Content $tmp.fullname
									$txtInfo.AppendText($log+"`n"+"`n") 
									Stop-Transcript
									$txtInfo.AppendText("Готово!")
								}
									else
									{
                                     

									foreach($database in $databases){
									   if ($comboBoxOU.Text)
									   {
									     $DBCount=(Get-Mailbox -ResultSize unlimited -OrganizationalUnit $comboBoxOU.Text -Database $database | Measure-Object).count
									   }
							           else
									   {
										 $DBCount=(Get-Mailbox -ResultSize unlimited -Database $database | Measure-Object).count
									   }
								
										$count+=$DBCount
										if($txtinfo.Text -eq "Подождите..."){
											$txtInfo.Text=""
										}
										$txtInfo.AppendText($database+"="+$DBcount+"`n")
										}
										
#										}
											$txtInfo.AppendText("All="+$count+"`n"+"`n")
											$log = Get-Content $tmp.fullname
											Stop-Transcript
											$txtInfo.AppendText($log+"`n"+"`n")
											$txtInfo.AppendText("Готово!")
											}
								}
							

$btn6 = New-Object System.Windows.Forms.Button;  
$btn6.dock = "left"
$btn6.Height = 40
$btn6.Size = New-Object System.Drawing.Size(100, 25);    
$btn6.Name = "btn6";    
$btn6.Text = "Export Move MB"; 
$btn6.Enabled = $false;  
$btn6.add_DoubleClick({write "asdad"})
$btn6.Add_Click({
	count-move
})
$btn6.add_MouseHover({
	$tooltip.SetToolTip($btn6, "Считает количество мигрируемых или смигрированных почтовых ящиков.")
})
$GroupInside.Controls.Add($btn6);

					function count-move {
					$ListMB.Items.Clear()
						$txtInfo.Text="Подождите..." 
						Start-Transcript -Path $tmp.fullname
#						if($list.CheckedItems){
							$databases = $List.CheckedItems.text
							foreach($database in $databases){
								$moves = Get-MoveRequestStatistics -MoveRequestQueue $database | Select-Object -Property *
										$id=1
										foreach ( $move in $moves )  
									    	{
											try{
											$Alias = $move.Alias -replace "[.].*",""
											$mDBUseDefaults=(Get-ADUser $Alias -Properties *).mDBUseDefaults
											$str1 =	Get-MailboxStatistics -Identity $Alias | select itemcount,totalitemsize,totaldeleteditemsize
											$sizeIS=$str1.totalitemsize -replace ".*[(]",""
											$sizeIS1="$sizeIS" -replace "[by].*",""
											$sizeIDS=$str1.totaldeleteditemsize -replace ".*[(]",""
											$sizeIDS1="$sizeIDS" -replace "[by].*",""
											$sizeMB=$move.totalmailboxsize -replace ".*[(]",""
											$sizeMB1="$sizeMB" -replace "[by].*",""
											$ProhibitSendQuota=(Get-Mailbox -Identity $Alias).ProhibitSendQuota
											$ProhibitSendQuota1=$ProhibitSendQuota -replace "[by].*",""
											$moveTDB = $move.TargetDatabase
											$moveSDB = $move.SourceDatabase
											$StartTS=$move.StartTimestamp
											$CompletionTS=$move.CompletionTimestamp
											$identity = $move.Identity
											$smtp = (Get-Mailbox -Identity $Alias).PrimarySMtpAddress
											}catch{}
									    	#Добавляем элемент в список
											try{
											$I = $ListMB.Items.Add($id)
											if($move.DisplayName){$I.SubItems.Add($move.DisplayName)}else{$I.SubItems.Add("null")}
											if($move.Alias){$I.SubItems.Add($move.Alias)}else{$I.SubItems.Add("null")}
											if("$smtp"){$I.SubItems.Add("$smtp")}else{$I.SubItems.Add("null")}
											if($move.RecipientTypeDetails){$I.SubItems.Add($move.RecipientTypeDetails)}else{$I.SubItems.Add("null")}
											if("$identity"){$I.SubItems.Add("$identity")}else{$I.SubItems.Add("null")}
											if($str1.itemcount){$I.SubItems.Add($str1.itemcount)}else{$I.SubItems.Add("0")}
											if("$sizeIS1"){$I.SubItems.Add("$sizeIS1")}else{$I.SubItems.Add("0")}
											if("$sizeIDS1"){$I.SubItems.Add("$sizeIDS1")}else{$I.SubItems.Add("0")}
											if("$sizeMB1"){$I.SubItems.Add("$sizeMB1")}else{$I.SubItems.Add("0")}
											if("$ProhibitSendQuota1"){$I.SubItems.Add("$ProhibitSendQuota1")}else{$I.SubItems.Add("null")}
												if("$mDBUseDefaults"){$I.SubItems.Add("$mDBUseDefaults")}else{$I.SubItems.Add("null")}
											if("$moveTDB"){$I.SubItems.Add("$moveTDB")}else{$I.SubItems.Add("null")}
											if("$moveSDB"){$I.SubItems.Add("$moveSDB")}else{$I.SubItems.Add("null")}
											if($move.statusdetail){$I.SubItems.Add($move.statusdetail)}else{$I.SubItems.Add("null")}
											if($move.PercentComplete){$I.SubItems.Add($move.PercentComplete)}else{$I.SubItems.Add("0")}
											if("$StartTS"){$I.SubItems.Add("$StartTS")}else{$I.SubItems.Add("0")}
											if("$CompletionTS"){$I.SubItems.Add("$CompletionTS")}else{$I.SubItems.Add("0")}
											}catch{}
											
										
							##############Сравнение с квотой	
						$FullAccessFromInternetZF = Get-ADGroup "group" | Select-Object -ExpandProperty DistinguishedName
										$FullAccessFromInternetGO = Get-ADGroup "group" | Select-Object -ExpandProperty DistinguishedName
										if(((Get-ADUser $ListMB.Items[$id-1].Subitems[2].Text -Properties memberof).memberof -like $FullAccessFromInternetZF)-or((Get-ADUser $ListMB.Items[$id-1].Subitems[2].Text -Properties memberof).memberof -like $FullAccessFromInternetGO)){
												$ListMB.Items[$id-1].forecolor = "red"
											}
		                 
							   [string]$stringtext = Get-MailboxDatabase $ListMB.Items[$id-1].SubItems[12].Text -Status | select ProhibitSendQuota
	
							   
									$stringlist = $stringtext.Split("(",2) -replace "[bytes]","" -replace "[)]","" -replace ",","" -replace "[}]",""
									$kvota = [int64]$stringlist[1]
		                            $total =  [int64]($ListMB.Items[$id-1].SubItems[7].Text -replace ",","")
								
								
									if ($total -ge $kvota)
										{
										  $ListMB.Items[$id-1].backcolor = "gray"
										}
							
										
									    	$id+=1
									}

								
							}
									$log = Get-Content $tmp.fullname
									$txtInfo.AppendText($log+"`n"+"`n") 
									Stop-Transcript
									$txtInfo.AppendText("Готово!")
					}

$btn2 = New-Object System.Windows.Forms.Button;  
$btn2.dock = "left"
$btn2.Height = 40
$btn2.Size = New-Object System.Drawing.Size(100, 25);    
$btn2.Name = "btn2";    
$btn2.Text = "Update Store MB"; 
$btn2.Enabled = $false;    
$btn2.Add_Click({
		$txtInfo.Text="Подождите..." 
		Start-Transcript -Path $tmp.fullname
		$databases = $List.CheckedItems.text
	foreach($database in $databases){
	Get-MailboxStatistics -Database $database | ForEach { Update-StoreMailboxState -Database $_.Database -Identity $_.MailboxGuid -Confirm:$false }
	}
			$log = Get-Content $tmp.fullname
			Stop-Transcript
			$txtInfo.text=$log+"`n"+"`n"
			$txtInfo.AppendText("Готово!")
})
$btn2.add_MouseHover({
	$tooltip.SetToolTip($btn2, "Очистка истории почтовых ящиков.")
})
$GroupInside.Controls.Add($btn2);

$btn3 = New-Object System.Windows.Forms.Button;  
$btn3.dock = "left"
$btn3.Height = 40
$btn3.Size = New-Object System.Drawing.Size(100, 25);    
$btn3.Name = "btn3";    
$btn3.Text = "Unlimited MB"; 
$btn3.Enabled = $false;    
$btn3.Add_Click({
	UnlimitedMB
	})
$btn3.add_MouseHover({
	$tooltip.SetToolTip($btn3, "Выгрузка списка почтовых ящиков с персональной квотой.")
})
$GroupInside.Controls.Add($btn3);

function UnlimitedMB ($Mask = "*") {
	$txtInfo.Text="Подождите..." 
	Start-Transcript -Path $tmp.fullname
	$ListMB.Items.Clear()
	$databases = $List.CheckedItems.text
	foreach($database in $databases){
    $str = Get-mailbox -Database $database -ResultSize unlimited | ? {$_.ProhibitSendQuota -eq "Unlimited"} | select PrimarySmtpAddress,Alias,DisplayName,ProhibitSendQuota,database,identity,RecipientTypeDetails,SamAccountName
		if($str -eq $null){
		$str = Get-mailbox -Database $database -ResultSize unlimited | ? {$_.ProhibitSendQuota -eq "Unlimited"} | select PrimarySmtpAddress,Alias,DisplayName,ProhibitSendQuota,database,identity,RecipientTypeDetails,SamAccountName}
		    if($str -eq $null)
			{
			$I = $ListMB.Items.Add("Пользователи не найден...") | Out-Null
			}
			else
			{
			$id=1
			foreach ( $item in $str ) 
		    	{
				
					$mDBUseDefaults=(Get-ADUser $item.SamAccountName -Properties *).mDBUseDefaults
					if ($mDBUseDefaults -eq $true)
				{
									$str1 = Get-MailboxStatistics $item.PrimarySmtpAddress | select itemcount,totalitemsize,totaldeleteditemsize
									$saved=$global:ErrorActionPreference
    								$global:ErrorActionPreference='stop'
									try{
									$moves=Get-MoveRequest -identity $item.PrimarySmtpAddress | Get-MoveRequestStatistics | Select-Object statusdetail,percentcomplete,totalmailboxsize,StartTimestamp,CompletionTimestamp
									}Catch{
								         Write-Warning $_
								    }
								    Finally{
								        $global:ErrorActionPreference=$saved
									}
				$sizeIS=$str1.totalitemsize -replace ".*[(]",""
				$sizeIS1="$sizeIS" -replace "[by].*",""
				$sizeIDS=$str1.totaldeleteditemsize -replace ".*[(]",""
				$sizeIDS1="$sizeIDS" -replace "[by].*",""
				$sizeMB=$moves.totalmailboxsize -replace ".*[(]",""
				$sizeMB1="$sizeMB" -replace "[by].*",""
				$ProhibitSendQuota=$item.ProhibitSendQuota -replace ".*[(]",""
				$ProhibitSendQuota1=$ProhibitSendQuota -replace "[by].*",""
				$moveSDB = $move.SourceDatabase
				$StartTS=$move.StartTimestamp
				$CompletionTS=$move.CompletionTimestamp
		    	#Добавляем элемент в список
				$I = $ListMB.Items.Add($id)
				if($item.DisplayName){$I.SubItems.Add($item.DisplayName)}else{$I.SubItems.Add("null")}
				if($item.Alias){$I.SubItems.Add($item.Alias)}else{$I.SubItems.Add("null")}
				if($item.PrimarySmtpAddress){$I.SubItems.Add($item.PrimarySmtpAddress)}else{$I.SubItems.Add("null")}
				if($item.RecipientTypeDetails){$I.SubItems.Add($item.RecipientTypeDetails)}else{$I.SubItems.Add("null")}
				if($item.Identity){$I.SubItems.Add($item.Identity)}else{$I.SubItems.Add("null")}
				if($str1.itemcount){$I.SubItems.Add($str1.itemcount)}else{$I.SubItems.Add("0")}
				if("$sizeIS1"){$I.SubItems.Add("$sizeIS1")}else{$I.SubItems.Add("0")}
				if("$sizeIDS1"){$I.SubItems.Add("$sizeIDS1")}else{$I.SubItems.Add("0")}
				if("$sizeMB1"){$I.SubItems.Add("$sizeMB1")}else{$I.SubItems.Add("0")}
				if("$ProhibitSendQuota1"){$I.SubItems.Add("$ProhibitSendQuota1")}else{$I.SubItems.Add("null")}
					if("$mDBUseDefaults"){$I.SubItems.Add("$mDBUseDefaults")}else{$I.SubItems.Add("null")}
				if($item.Database){$I.SubItems.Add($item.Database)}else{$I.SubItems.Add("null")}
				if("$moveSDB"){$I.SubItems.Add("$moveSDB")}else{$I.SubItems.Add("null")}
				if($moves.StatusDetail){$I.SubItems.Add($moves.StatusDetail)}else{$I.SubItems.Add("null")}
				if($moves.percentcomplete){$I.SubItems.Add($moves.percentcomplete)}else{$I.SubItems.Add("0")}
				if("$StartTS"){$I.SubItems.Add("$StartTS")}else{$I.SubItems.Add("0")}
				if("$CompletionTS"){$I.SubItems.Add("$CompletionTS")}else{$I.SubItems.Add("0")}
				$id+=1
				}
		      }
			}
				if ($ListMB.Items.Count -gt 1){
				SortListWithID -Column 1
				}
			$log = Get-Content $tmp.fullname
			Stop-Transcript
			$txtInfo.Text=$log+"`n"+"`n"
			$txtInfo.AppendText("Готово!")
	
	}}

$btn4 = New-Object System.Windows.Forms.Button;  
$btn4.dock = "left"
$btn4.Height = 40
$btn4.Size = New-Object System.Drawing.Size(100, 25);    
$btn4.Name = "btn4";    
$btn4.Text = "Export MB"; 
$btn4.Enabled = $false;  
#Функция поиска ПЯ по OU и БД
function AllMBOU
{
$ListMB.Items.Clear()
					foreach ($baza in $List.CheckedItems.Text)
					{
   							 $mailboxOU = Get-Mailbox -Database $baza -OrganizationalUnit $comboBoxOU.Text -ResultSize unlimited | select PrimarySmtpAddress,Alias,DisplayName,ProhibitSendQuota,database,identity,RecipientTypeDetails,SamAccountName 
									if($mailboxOU){
										$id=1
									foreach ($item in $mailboxOU) 
								    	{
										$mDBUseDefaults=(Get-ADUser $item.SamAccountName -Properties *).mDBUseDefaults
										$item1=$item.PrimarySmtpAddress
										$str1 = Get-MailboxStatistics $item1 | select itemcount,totalitemsize,totaldeleteditemsize
											$saved=$global:ErrorActionPreference
		    								$global:ErrorActionPreference='stop'
											try{
												$moves = Get-MoveRequest -identity $item1 | Get-MoveRequestStatistics | Select-Object statusdetail,percentcomplete,sourcedatabase,totalmailboxsize,StartTimestamp,CompletionTimestamp
											}Catch{
										         Write-Warning $_
										    }Finally{
										        $global:ErrorActionPreference=$saved
											}
										$sizeIS=$str1.totalitemsize -replace ".*[(]",""
										$sizeIS1="$sizeIS" -replace "[by].*",""
										$sizeIDS=$str1.totaldeleteditemsize -replace ".*[(]",""
										$sizeIDS1="$sizeIDS" -replace "[by].*",""
										$sizeMB=$moves.totalmailboxsize -replace ".*[(]",""
										$sizeMB1="$sizeMB" -replace "[by].*",""
										$ProhibitSendQuota=$item.ProhibitSendQuota -replace ".*[(]",""
										$ProhibitSendQuota1=$ProhibitSendQuota -replace "[by].*",""
										$ser=$moves.StatusDetail
										$moveSDB = $move.SourceDatabase
										$StartTS=$move.StartTimestamp
										$CompletionTS=$move.CompletionTimestamp
								    	#Добавляем элемент в список
										$I = $ListMB.Items.Add($id)
										if($item.DisplayName){$I.SubItems.Add($item.DisplayName)}else{$I.SubItems.Add("null")}
										if($item.samaccountname){$I.SubItems.Add($item.samaccountname)}else{$I.SubItems.Add("null")}
										if($item.PrimarySmtpAddress){$I.SubItems.Add($item.PrimarySmtpAddress)}else{$I.SubItems.Add("null")}
										if($item.RecipientTypeDetails){$I.SubItems.Add($item.RecipientTypeDetails)}else{$I.SubItems.Add("null")}
										if($item.Identity){$I.SubItems.Add($item.Identity)}else{$I.SubItems.Add("null")}
										if($str1.itemcount){$I.SubItems.Add($str1.itemcount)}else{$I.SubItems.Add("0")}
										if("$sizeIS1"){$I.SubItems.Add("$sizeIS1")}else{$I.SubItems.Add("0")}
										if("$sizeIDS1"){$I.SubItems.Add("$sizeIDS1")}else{$I.SubItems.Add("0")}
										if("$sizeMB1"){$I.SubItems.Add("$sizeMB1")}else{$I.SubItems.Add("0")}
										if("$ProhibitSendQuota1"){$I.SubItems.Add("$ProhibitSendQuota1")}else{$I.SubItems.Add("null")}
											if("$mDBUseDefaults"){$I.SubItems.Add("$mDBUseDefaults")}else{$I.SubItems.Add("null")}
										if($item.Database){$I.SubItems.Add($item.Database)}else{$I.SubItems.Add("null")}
										if("$moveSDB"){$I.SubItems.Add("$moveSDB")}else{$I.SubItems.Add("null")}
										if($moves.statusdetail){$I.SubItems.Add($moves.statusdetail)}else{$I.SubItems.Add("null")}
										if($moves.percentcomplete){$I.SubItems.Add($moves.percentcomplete)}else{$I.SubItems.Add("0")}
										if("$StartTS"){$I.SubItems.Add("$StartTS")}else{$I.SubItems.Add("0")}
										if("$CompletionTS"){$I.SubItems.Add("$CompletionTS")}else{$I.SubItems.Add("0")}

										if(( ((Get-ADUser $ListMB.Items[$i].Subitems[2].Text -Properties MemberOf).memberof | get-adgroup).SID.value -like "SID group" ) -or ( ((Get-ADUser $ListMB.Items[$i].Subitems[2].Text -Properties MemberOf).memberof | get-adgroup).sid.value -like "SID group")){
												$ListMB.Items[$id-1].forecolor = "red"
											}
##Сравнение с квотой
 if($ListMB.Items[$id-1].Subitems[12].Text  -eq $item.Database )
							 {
							   [string]$stringtext = Get-MailboxDatabase $item.Database -Status | select ProhibitSendQuota
						
									$stringlist = $stringtext.Split("(",2) -replace "[bytes]","" -replace "[)]","" -replace ",","" -replace "[}]",""
								
									$kvota = [int64]$stringlist[1]
		                            $total =  [int64]($ListMB.Items[$id-1].SubItems[7].Text -replace ",","")
								
								
									if ($total -ge $kvota)
										{
										  $ListMB.Items[$id-1].backcolor = "gray"
										}
							}
								    	$id++
										}
								}
									else
								{
									$I = $ListMB.Items.Add("1") 
									$I.SubItems.Add("Пользователь не найден") | Out-Null
										$I.SubItems.Add("null")
										$I.SubItems.Add("null")
									$I.SubItems.Add("null")
										$I.SubItems.Add("null")
										$I.SubItems.Add("0")
										$I.SubItems.Add("0")
										$I.SubItems.Add("0")
										$I.SubItems.Add("0")
										$I.SubItems.Add("null")
											$I.SubItems.Add("null")
										$I.SubItems.Add($baza)
										$I.SubItems.Add("null")
								$I.SubItems.Add("null")
										$I.SubItems.Add("0")
										$I.SubItems.Add("0")
										$I.SubItems.Add("0")
										
									
									
								}	
								}
									if ($ListMB.Items.Count -gt 1){
									SortListWithID -Column 1
									}
									$log = Get-Content $tmp.fullname
									$txtInfo.Text = $log+"`n"+"`n"
									Stop-Transcript
									$txtInfo.AppendText("Готово!")
									$comboBoxOU.Text = $null
									for ($k =0; $k -lt $list.items.count; $k++)
					{
					  if ($list.items[$k].text -eq "NRDB04")
					  {
					    $List.Items[$k].Backcolor = "green"
					  }
					}
					

}


$btn4.Add_Click({
  if ($comboBoxOU.Text)
					{
						AllMBOU
					}
			else
		{
		AllMB
		}
		})
$btn4.add_MouseHover({
	$tooltip.SetToolTip($btn4, "Экспорт списка почтовых ящиков.")
})
$GroupInside.Controls.Add($btn4);

					function AllMB ($Mask = "*") {
					$txtInfo.Text="Подождите..." 
					Start-Transcript -Path $tmp.fullname
						$ListMB.Items.Clear()
						$databases = $List.CheckedItems.text
						foreach($database in $databases){
					    $str = Get-mailbox -Database $database | select PrimarySmtpAddress,Alias,DisplayName,ProhibitSendQuota,database,identity,RecipientTypeDetails,SamAccountName
							    if($str){
								$id=1
									foreach ($item in $str) 
								    	{
										$mDBUseDefaults=(Get-ADUser $item.SamAccountName -Properties *).mDBUseDefaults
										$item1=$item.PrimarySmtpAddress
										$str1 = Get-MailboxStatistics $item1 | select itemcount,totalitemsize,totaldeleteditemsize
											$saved=$global:ErrorActionPreference
		    								$global:ErrorActionPreference='stop'
											try{
												$moves = Get-MoveRequest -identity $item1 | Get-MoveRequestStatistics | Select-Object statusdetail,percentcomplete,sourcedatabase,totalmailboxsize,StartTimestamp,CompletionTimestamp 
											}Catch{
										         Write-Warning $_
										    }Finally{
										        $global:ErrorActionPreference=$saved
											}
										$sizeIS=$str1.totalitemsize -replace ".*[(]",""
										$sizeIS1="$sizeIS" -replace "[by].*",""
										$sizeIDS=$str1.totaldeleteditemsize -replace ".*[(]",""
										$sizeIDS1="$sizeIDS" -replace "[by].*",""
										$sizeMB=$moves.totalmailboxsize -replace ".*[(]",""
										$sizeMB1="$sizeMB" -replace "[by].*",""
										$ProhibitSendQuota=$item.ProhibitSendQuota -replace ".*[(]",""
										$ProhibitSendQuota1=$ProhibitSendQuota -replace "[by].*",""
										$ser=$moves.StatusDetail
										$moveSDB = $move.SourceDatabase
										$StartTS=$move.StartTimestamp
										$CompletionTS=$move.CompletionTimestamp
								    	#Добавляем элемент в список
										$I = $ListMB.Items.Add($id)
										if($item.DisplayName){$I.SubItems.Add($item.DisplayName)}else{$I.SubItems.Add("null")}
										if($item.samaccountname){$I.SubItems.Add($item.samaccountname)}else{$I.SubItems.Add("null")}
										if($item.PrimarySmtpAddress){$I.SubItems.Add($item.PrimarySmtpAddress)}else{$I.SubItems.Add("null")}
										if($item.RecipientTypeDetails){$I.SubItems.Add($item.RecipientTypeDetails)}else{$I.SubItems.Add("null")}
										if($item.Identity){$I.SubItems.Add($item.Identity)}else{$I.SubItems.Add("null")}
										if($str1.itemcount){$I.SubItems.Add($str1.itemcount)}else{$I.SubItems.Add("0")}
										if("$sizeIS1"){$I.SubItems.Add("$sizeIS1")}else{$I.SubItems.Add("0")}
										if("$sizeIDS1"){$I.SubItems.Add("$sizeIDS1")}else{$I.SubItems.Add("0")}
										if("$sizeMB1"){$I.SubItems.Add("$sizeMB1")}else{$I.SubItems.Add("0")}
										if("$ProhibitSendQuota1"){$I.SubItems.Add("$ProhibitSendQuota1")}else{$I.SubItems.Add("null")}
											if("$mDBUseDefaults"){$I.SubItems.Add("$mDBUseDefaults")}else{$I.SubItems.Add("null")}
										if($item.Database){$I.SubItems.Add($item.Database)}else{$I.SubItems.Add("null")}
										if("$moveSDB"){$I.SubItems.Add("$moveSDB")}else{$I.SubItems.Add("null")}
										if($moves.statusdetail){$I.SubItems.Add($moves.statusdetail)}else{$I.SubItems.Add("null")}
										if($moves.percentcomplete){$I.SubItems.Add($moves.percentcomplete)}else{$I.SubItems.Add("0")}
										if("$StartTS"){$I.SubItems.Add("$StartTS")}else{$I.SubItems.Add("0")}
										if("$CompletionTS"){$I.SubItems.Add("$CompletionTS")}else{$I.SubItems.Add("0")}

										if((((Get-ADUser $item.SamAccountName -Properties MemberOf).memberof | get-adgroup).SID.value -like "SID group" ) `
										-or ( ((Get-ADUser $item.SamAccountName -Properties MemberOf).memberof | get-adgroup).sid.value -like "SID group")){
												$ListMB.Items[$id-1].forecolor = "red"
											}
										##Сравнение с квотой
 										if($ListMB.Items[$id-1].Subitems[12].Text  -eq $item.Database )
							 				{
											   [string]$stringtext = Get-MailboxDatabase $item.Database -Status | select ProhibitSendQuota
										
													$stringlist = $stringtext.Split("(",2) -replace "[bytes]","" -replace "[)]","" -replace ",","" -replace "[}]",""
												
													$kvota = [int64]$stringlist[1]
						                            $total =  [int64]($ListMB.Items[$id-1].SubItems[7].Text -replace ",","")
												
												
													if ($total -ge $kvota)
														{
														  $ListMB.Items[$id-1].backcolor = "gray"
														}
											}
								    	$id++
										}
								}else{
									$I = $ListMB.Items.Add("1") 
									$I.SubItems.Add("Пользователь не найден...") | Out-Null
								}
									if ($ListMB.Items.Count -gt 1){
									SortListWithID -Column 1
									}
									$log = Get-Content $tmp.fullname
									$txtInfo.Text =$log+"`n"+"`n"
									Stop-Transcript
									$txtInfo.AppendText("Готово!")
						}
					}

$btn5 = New-Object System.Windows.Forms.Button;  
$btn5.dock = "left"
$btn5.Height = 40
$btn5.Size = New-Object System.Drawing.Size(100, 25);    
$btn5.Name = "btn5";    
$btn5.Text = "Migrate MB"; 
$btn5.Enabled = $false;    
$btn5.Add_Click({
	$txtInfo.Text="Подождите..." 
	Start-Transcript -Path $tmp.fullname
			$smtpaddress = @()
			$n = "`n"
			$ListMB.SelectedItems | % {$smtpaddress += $_.subitems[3].text+$n}
			ForEach ($smtpaddres in $smtpaddress){
				$smtpaddres1 = $smtpaddres -replace "[`n]",""
				New-MoveRequest -Identity $smtpaddres1 -TargetDatabase $List.CheckedItems.text -BadItemLimit "200" -priority emergency
			}
				$log = Get-Content $tmp.fullname
				$txtInfo.Text=$log+"`n"+"`n"
				Stop-Transcript
				$txtInfo.AppendText("Готово!")
})
$btn5.add_MouseHover({
	$tooltip.SetToolTip($btn5, "Миграция почтовых ящиков между базами. Сначала выюерите нужную базу, а потом почтовый ящик.")
})
$GroupInside.Controls.Add($btn5);



$btn7 = New-Object System.Windows.Forms.Button
$btn7.dock = "left"
$btn7.Height = 40
$btn7.Size = New-Object System.Drawing.Size(100, 25);    
$btn7.Name = "btn7";    
$btn7.Text = "Refresh DB"; 
$btn7.Enabled = $true;    
$btn7.Add_Click({
	$List.Items.Clear()
	DB
	SortListOneviewDB
	for ($k =0; $k -lt $list.items.count; $k++)
					{
					  if ($list.items[$k].text -eq "NRDB04")
					  {
					    $List.Items[$k].Backcolor = "green"
					  }
					}
	})
$btn7.add_MouseHover({
	$tooltip.SetToolTip($btn7, "Обновление списка баз данных.")
})
$GroupInside.Controls.Add($btn7);
#------------------------------------------------------------

						#Сортироватm в два вида по умолчанию (возрастания и убывание)-------------------------------------------------------------------------------
						function SortListWithID {
						 param([parameter(Position=0)][UInt32]$Column)
						$Numeric = $true # определить, как сортировать (determine how to sort)
						#если пользователь нажал тот же столбец, который был выбран последний раз, его обратный порядок сортировки. в противном случае, сброс для нормальной сортировки по возрастанию
						#if the user clicked the same column that was clicked last time, reverse its sort order. otherwise, reset for normal ascending sort
						if($Script:LastColumnClickedTwo -eq $Column-or$Script:LastColumnClickedOne -eq $Column){
						    $Script:LastColumnAscendingTwo = -not $Script:LastColumnAscendingTwo
						}else{
						    $Script:LastColumnAscendingTwo = $true
						}
						$Script:LastColumnClickedTwo = $Column
						#трехмерный массив; колонке 1 индексы других столбцов, столбец 2 является значением, которое будет отсортирован, и колонка 3 является System.Windows.Forms.ListViewItem object
						#three-dimensional array; column 1 indexes the other columns, column 2 is the value to be sorted on, and column 3 is the System.Windows.Forms.ListViewItem object
						$ListItems = @(@(@()))
						foreach($ListItem in $ListMB.Items){
						    #если все элементы являются числовыми, могут использовать числовую сортировку (if all items are numeric, can use a numeric sort)
						    if($Numeric -ne $false) #ничто не может установить значение True, поэтому не процесс излишне (nothing can set this back to true, so don't process unnecessarily)
						    {
						        try
						        {
						            $Test = [Double]$ListItem.SubItems[[int]$Column].Text
						        }
						        catch
						        {
						            $Numeric = $false #найден нечисловых элементов, так что сортировка будет происходить в виде строки (a non-numeric item was found, so sort will occur as a string)
								}
						    }
						    $ListItems += ,@($ListItem.SubItems[[int]$Column].Text,$ListItem)
						}
							
						#создать выражение, которое будет вычисляться для сортировки (create the expression that will be evaluated for sorting)
						$columntext=$listMB.Columns[[int]$Column].text
						$EvalExpression = {
						    if($Numeric)
						    { return [double]$_[0] }
						    else
						    { return [String]$_[0] }
						}
							if(!$Numeric){
								#создание массива из таблицы
								$Arrays = ""
								$Arrays = @()
								$ArraysForeColor = ""
								$ArraysForeColor = @()
								$ArraysBackColor = ""
								$ArraysBackColor = @()
							    $listMB.Items | %{
							        $Object = New-Object PSObject
									$ObjForeColor = New-Object PSObject
									$ObjBackColor = New-Object PSObject
							        $lvItem = $_
							        $listMB.Columns | %{
							            $Object | add-member Noteproperty -Name $_.Text -Value $lvItem.SubItems[$_.index].text -Force
										$ObjForeColor | add-member Noteproperty -Name $_.text -Value $lvItem.SubItems[$_.index].forecolor -Force
										$ObjBackColor | add-member Noteproperty -Name $_.text -Value $lvItem.SubItems[$_.index].backcolor -Force
							        }
									$Arrays += $Object
									$ArraysForeColor += $Object.displayname + ";ForeColor=" + $ObjForeColor.id.name
									$ArraysBackColor += $Object.displayname + ";BackColor=" + $ObjBackColor.id.name
							        Remove-Variable Object
							    }
								$Arrays = $Arrays | Sort-Object -Property @{Expression=$columntext; Ascending=$Script:LastColumnAscendingTwo}
								#список отсортирован, вывести в list (the list is sorted; display it in the listview)
								$ListMB.BeginUpdate()
								$ListMB.Items.Clear()
								$o=""
								$o=1
								foreach($Array in $Arrays){
								    $I = $ListMB.Items.Add($o)
									$I.SubItems.Add($Array.displayname)
									$I.SubItems.Add($Array.Samaccountname)
									$I.SubItems.Add($Array.PrimarySMTPAddress)
									$I.SubItems.Add($Array.RecipientTypeDetails)
									$I.SubItems.Add($Array.Identity)
									$I.SubItems.Add($Array.ItemCount)
									$I.SubItems.Add($Array.totalitemsize)
									$I.SubItems.Add($Array.totaldeleteditemsize)
									$I.SubItems.Add($Array.totalmailboxsizeMR)
									$I.SubItems.Add($Array.ProhibitSendQuota)
									$I.SubItems.Add($Array.mDBUseDefaults)
									$I.SubItems.Add($Array.TargetDataBase)
									$I.SubItems.Add($Array.SourceDataBase)
									$I.SubItems.Add($Array.StatusDetailMR)
									$I.SubItems.Add($Array.PercentCompleteMR)
									$I.SubItems.Add($Array.StartTimestampMR)
									$I.SubItems.Add($Array.CompletionTimestampMR)
									if($ArraysForeColor -like ($Array.displayname + ";ForeColor=Red")){
											$ListMB.Items[$o-1].forecolor = "red"
										}
									if($ArraysBackColor -like ($Array.displayname + ";BackColor=Gray")){
											$ListMB.Items[$o-1].backcolor = "gray"
										}
									$o+=1
								}					
								$ListMB.EndUpdate()
							}else{
								#вся информация собрана; выполнения сортировки (all information is gathered; perform the sort)
								$ListItems = $ListItems | Sort-Object -Property @{Expression=$EvalExpression; Ascending=$Script:LastColumnAscendingTwo}
								#список отсортирован, вывести в list (the list is sorted; display it in the listview)
#								$ListMB.BeginUpdate()
								$ListMB.Items.Clear()
								foreach($ListItem in $ListItems){
								    $ListMB.Items.Add($ListItem[1])
								}
								#создание массива из таблицы
								$Arrays = ""
								$Arrays = @()
								$ArraysForeColor = ""
								$ArraysForeColor = @()
								$ArraysBackColor = ""
								$ArraysBackColor = @()
							    $listMB.Items | %{
							        $Object = New-Object PSObject
									$ObjForeColor = New-Object PSObject
									$ObjBackColor = New-Object PSObject
							        $lvItem = $_
							        $listMB.Columns | %{
							            $Object | add-member Noteproperty -Name $_.Text -Value $lvItem.SubItems[$_.index].text -Force
										$ObjForeColor | add-member Noteproperty -Name $_.text -Value $lvItem.SubItems[$_.index].forecolor -Force
										$ObjBackColor | add-member Noteproperty -Name $_.text -Value $lvItem.SubItems[$_.index].backcolor -Force
							        }
									$Arrays += $Object
									$ArraysForeColor += $Object.displayname + ";ForeColor=" + $ObjForeColor.id.name
									$ArraysBackColor += $Object.displayname + ";BackColor=" + $ObjBackColor.id.name
							        Remove-Variable Object
							    }
								#список отсортирован, вывести в list (the list is sorted; display it in the listview)
								$ListMB.BeginUpdate()
								$ListMB.Items.Clear()
								$o=""
								$o=1
								foreach($Array in $Arrays){
								    $I = $ListMB.Items.Add($o)
									$I.SubItems.Add($Array.displayname)
									$I.SubItems.Add($Array.Samaccountname)
									$I.SubItems.Add($Array.PrimarySMTPAddress)
									$I.SubItems.Add($Array.RecipientTypeDetails)
									$I.SubItems.Add($Array.Identity)
									$I.SubItems.Add($Array.ItemCount)
									$I.SubItems.Add($Array.totalitemsize)
									$I.SubItems.Add($Array.totaldeleteditemsize)
									$I.SubItems.Add($Array.totalmailboxsizeMR)
									$I.SubItems.Add($Array.ProhibitSendQuota)
									$I.SubItems.Add($Array.mDBUseDefaults)
									$I.SubItems.Add($Array.TargetDataBase)
									$I.SubItems.Add($Array.SourceDataBase)
									$I.SubItems.Add($Array.StatusDetailMR)
									$I.SubItems.Add($Array.PercentCompleteMR)
									$I.SubItems.Add($Array.StartTimestampMR)
									$I.SubItems.Add($Array.CompletionTimestampMR)
									if($ArraysForeColor -like ($Array.displayname + ";ForeColor=Red")){
											$ListMB.Items[$o-1].forecolor = "red"
										}
									if($ArraysBackColor -like ($Array.displayname + ";BackColor=Gray")){
											$ListMB.Items[$o-1].backcolor = "gray"
										}
									$o+=1
								}					
							$ListMB.EndUpdate()
							}
						}
						


						#-------------------------------------------------------------------------------------------------------------------------------

#контекстное меню вызывается в Listview
$menumInfUser = New-Object System.Windows.Forms.MenuItem
$menumInfUser.Text = "Информация о пользователе"
$menumInfUser.Enabled = $false
$menumInfUser.Add_Click({
 FunctionUser
 $FormInfUser.ShowDialog()
})

$menucopy = New-Object System.Windows.Forms.MenuItem
$menucopy.Text = "Копировать содержимое"
$menucopy.Enabled = $false
$menucopy.Add_Click({
Set-ClipBoard
})
$menucopyMB = New-Object System.Windows.Forms.MenuItem
$menucopyMB.Text = "Копировать содержимое"
$menucopyMB.Enabled = $false
$menucopyMB.Add_Click({
Set-ClipBoardMB
})
$menususpendMB = New-Object System.Windows.Forms.MenuItem
$menususpendMB.Text = "Приостоновить перемещение"
$menususpendMB.Enabled = $false
$menususpendMB.Add_Click({
Suspend-Move
})
$menuresumeMB = New-Object System.Windows.Forms.MenuItem
$menuresumeMB.Text = "Возобновить перемещение"
$menuresumeMB.Enabled = $false
$menuresumeMB.Add_Click({
Resume-Move
})
$menumDBUseDefaultsFalseMB = New-Object System.Windows.Forms.MenuItem
$menumDBUseDefaultsFalseMB.Text = "Отключить квоту"
$menumDBUseDefaultsFalseMB.Enabled = $false
$menumDBUseDefaultsFalseMB.Add_Click({
mDBUseDefaults-false
})
$menumDBUseDefaultsTrueMB = New-Object System.Windows.Forms.MenuItem
$menumDBUseDefaultsTrueMB.Text = "Включить квоту"
$menumDBUseDefaultsTrueMB.Enabled = $false
$menumDBUseDefaultsTrueMB.Add_Click({
mDBUseDefaults-true
})
$menumDBADDPermissionFullAccess = New-Object System.Windows.Forms.MenuItem
$menumDBADDPermissionFullAccess.Text = "Предоставить права FullAccess на почтовый ящик"
$menumDBADDPermissionFullAccess.Enabled = $false
$menumDBADDPermissionFullAccess.Add_Click({
$comboboxCalendar.Visible = $false
AddPermissionFullAccess
})

$menumDBADPermissionSend_AS = New-Object System.Windows.Forms.MenuItem
$menumDBADPermissionSend_AS.Text = "Предоставить права Send-AS на почтовый ящик"
$menumDBADPermissionSend_AS.Enabled = $false
$menumDBADPermissionSend_AS.Add_Click({
$comboboxCalendar.Visible = $false
ADDPermessionSend-As
})

$menumDBDeletePermission = New-Object System.Windows.Forms.MenuItem
$menumDBDeletePermission.Text = "Удаленить доступ к почтовому ящику"
$menumDBDeletePermission.Enabled = $false
$menumDBDeletePermission.Add_Click({
$comboboxCalendar.Visible = $false
DeletePermission 
})

$menumDBAccessFromInternet = New-Object System.Windows.Forms.MenuItem
$menumDBAccessFromInternet.Text = "Добавить или удалить пользователя из группы доступа копоротивной почты из Интернета"
$menumDBAccessFromInternet.Enabled = $false
$menumDBAccessFromInternet.Add_Click({
$SID = "SID goupr"
    $group = Get-ADGroup $SID
	$user = Get-ADUser $listMB.Items[$listMB.FocusedItem.Index].Subitems[2].Text 
 				if ($listMB.Items[$listMB.FocusedItem.Index].forecolor -eq "red")
 					{
         			Remove-ADGroupMember $group -Member $user -Server "npr.nornick.ru" -Confirm:$false
					$ListMB.Items[$ListMB.FocusedItem.Index].forecolor = "black"
					[System.Windows.Forms.MessageBox]::Show("Пользователь удален из группы!")
					

 					}
					else
					{
     				Add-ADGroupMember  $group  $user -Server "npr.nornick.ru"
					$listMB.Items[$listMB.FocusedItem.Index].forecolor = "red"
					[System.Windows.Forms.MessageBox]::Show("Данный пользователь добавлен в группу")
    				}

})

$menumAccessCalendar = New-Object System.Windows.Forms.MenuItem
$menumAccessCalendar.Text = "Предоставить доступ к календарю"
$menumAccessCalendar.Enabled = $false
$menumAccessCalendar.Add_Click({
$indexform = 3
$comboboxCalendar.Visible = $true
$fromAddPermission.Text = "Предоставить доступ к календарю"
$fromAddPermission.Update()
$fromAddPermission.Refresh()
$fromAddPermission.ShowDialog()
})

$menumDeleteCalendar = New-Object System.Windows.Forms.MenuItem
$menumDeleteCalendar.Text = "Удалить доступ к календарю"
$menumDeleteCalendar.Enabled = $false
$menumDeleteCalendar.Add_Click({
$indexform = 4
$comboboxCalendar.Visible = $false
$fromAddPermission.Text = "Удалить доступ к календарю"
$fromAddPermission.Update()
$fromAddPermission.Refresh()
$buttonadd.Text = "Удалить учетную запись"
$fromAddPermission.ShowDialog()

})

$menumALLAccesscalendar = New-Object System.Windows.Forms.MenuItem
$menumALLAccesscalendar.Text = "Доступ к календарю"
$menumALLAccesscalendar.Enabled = $false
$menumALLAccesscalendar.Add_Click({
$calendar2 = $null
$calendar2 = $ListMB.Items[$ListMB.Focuseditem.Index].SubItems[3].Text+":\Календарь"
Get-MailboxFolderPermission -Identity $calendar2 | Out-GridView -Title "Доступ к календарю"

})

$ContextMenu = New-Object System.Windows.Forms.ContextMenu
$ContextMenu.MenuItems.AddRange(@($menucopy))
$List.ContextMenu = $ContextMenu

$ContextMenu1 = New-Object System.Windows.Forms.ContextMenu
$ContextMenu1.MenuItems.AddRange(@($menumInfUser))
$ContextMenu1.MenuItems.AddRange(@($menucopyMB))
$ContextMenu1.MenuItems.AddRange(@($menususpendMB))
$ContextMenu1.MenuItems.AddRange(@($menuresumeMB))
$ContextMenu1.MenuItems.AddRange(@($menumDBUseDefaultsFalseMB))
$ContextMenu1.MenuItems.AddRange(@($menumDBUseDefaultsTrueMB))
$ContextMenu1.MenuItems.AddRange(@($menumDBADDPermissionFullAccess))
$ContextMenu1.MenuItems.AddRange(@($menumDBADPermissionSend_AS))
$ContextMenu1.MenuItems.AddRange(@($menumDBDeletePermission))
$ContextMenu1.MenuItems.AddRange(@($menumDBAccessFromInternet))
$ContextMenu1.MenuItems.AddRange(@($menumAccessCalendar))
$ContextMenu1.MenuItems.AddRange(@($menumDeleteCalendar)) 
$ContextMenu1.MenuItems.AddRange(@($menumALLAccesscalendar))
$listMB.ContextMenu = $ContextMenu1
#------------------------------------------------------------------------------------------------------------------------------

#копирует содержимое listMB--------------------------------------------
Function Set-ClipBoardMB {
$CopyTexts = @()
$n = "`n"
$listMB.SelectedItems | % {$CopyTexts+=$_.subitems.text+$n}
ForEach ($CopyText in $CopyTexts){
$CopyText1 += ";$CopyText"
}
[System.Windows.Forms.Clipboard]::SetText($CopyText1)
}
#------------------------------------------------------------------------

#копирует содержимое listview--------------------------------------------
Function Set-ClipBoard {
$CopyTexts = @()
$n = "`n"
$list.SelectedItems | % {$CopyTexts+=$_.subitems.text+$n}
ForEach ($CopyText in $CopyTexts){
$CopyText1 += ";$CopyText"
}
[System.Windows.Forms.Clipboard]::SetText($CopyText1)
}
#------------------------------------------------------------------------

#копирует содержимое listMT--------------------------------------------
Function Set-ClipBoardMT {
$CopyTexts = @()
$n = "`n"
$listMT.SelectedItems | % {$CopyTexts+=$_.subitems.text+$n}
ForEach ($CopyText in $CopyTexts){
$CopyText1 += ";$CopyText"
}
[System.Windows.Forms.Clipboard]::SetText($CopyText1)
}
#------------------------------------------------------------------------

#приостанавливание перемещения
function Suspend-Move {
$txtInfo.Text="Подождите..." 
	Start-Transcript -Path $tmp.fullname
			$aliass = @()
			$n = "`n"
			$ListMB.SelectedItems | % {$aliass += $_.subitems[2].text+$n}
			ForEach ($alias in $aliass){
				$alias1 = $alias -replace "[`n]",""
				Suspend-MoveRequest -Identity $alias1 -Confirm:$false
			}
				$log = Get-Content $tmp.fullname
				$txtInfo.Text=$log+"`n"+"`n"
				Stop-Transcript
				$txtInfo.AppendText("Готово!")
}
#---------------------------------------------------------------------------------------------

#возобновление перемещения
function Resume-Move {
$txtInfo.Text="Подождите..." 
	Start-Transcript -Path $tmp.fullname
			$aliass = @()
			$n = "`n"
			$ListMB.SelectedItems | % {$aliass += $_.subitems[2].text+$n}
			ForEach ($alias in $aliass){
				$alias1 = $alias -replace "[`n]",""
				Resume-MoveRequest -Identity $alias1 -Confirm:$false
			}
				$log = Get-Content $tmp.fullname
				$txtInfo.Text=$log+"`n"+"`n"
				Stop-Transcript
				$txtInfo.AppendText("Готово!")
}
#---------------------------------------------------------------------------------------------

#Отключение квот
function mDBUseDefaults-false {
$txtInfo.Text="Подождите..." 
	Start-Transcript -Path $tmp.fullname
			$aliass = @()
			$n = "`n"
			$ListMB.SelectedItems | % {$aliass += $_.subitems[2].text+$n}
			ForEach ($alias in $aliass){
				$alias1 = $alias -replace "[`n]",""
		
#				Write-Host $alias1
				Set-ADUser $alias1 -Replace @{mDBUseDefaults = $false}
			}
				$log = Get-Content $tmp.fullname
				$txtInfo.Text=$log+"`n"+"`n"
				Stop-Transcript
				$txtInfo.AppendText("Готово!")
}
#---------------------------------------------------------------------------------------------

#Включение квот
function mDBUseDefaults-true {
$txtInfo.Text="Подождите..." 
	Start-Transcript -Path $tmp.fullname
			$aliass = @()
			$n = "`n"
			$ListMB.SelectedItems | % {$aliass += $_.subitems[2].text+$n}
			ForEach ($alias in $aliass){
				$alias1 = $alias -replace "[`n]",""
#				Write-Host $alias1
				Set-ADUser $alias1 -Replace @{mDBUseDefaults = $true}
			}
				$log = Get-Content $tmp.fullname
				$txtInfo.Text=$log+"`n"+"`n"
				Stop-Transcript
				$txtInfo.AppendText("Готово!")
}
#---------------------------------------------------------------------------------------------

#Предоставление прав на общий ящик как FullAccess
function ADDPermissionFullAccess
{
$indexform = 1
$fromAddPermission.Text = "Предоставление полного доступа к общему ящику"
$fromAddPermission.Update()
$fromAddPermission.Refresh()
$fromAddPermission.ShowDialog()
}

#Предоставление права на общий ящик как Send-AS
function ADDPermessionSend-As
{
$indexform = 0
$fromAddPermission.Text = "Предоставление Send-AS доступа к общему ящику"
$fromAddPermission.Update()
$fromAddPermission.Refresh()
$fromAddPermission.ShowDialog()

}

#Удаление доступа к общему ящику
function DeletePermission
{
$indexform = 2
$fromAddPermission.Text = "Удаление доступа к общему ящику"
$buttonadd.Text = "Удалить доступ"
$fromAddPermission.Update()
$fromAddPermission.Refresh()
$fromAddPermission.ShowDialog()
}

$Exbody.ShowDialog()