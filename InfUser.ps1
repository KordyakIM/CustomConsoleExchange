###########################################################
# Разработчик: Кордяк Иван Михайлович kordyakim@gmail.com #
###########################################################
$FormInfUser = New-Object System.Windows.Forms.Form;
#$frmMain.icon =[system.drawing.icon]::ExtractAssociatedIcon("C:\Windows\System32\mmc.exe")   
$FormInfUser.ClientSize = New-Object System.Drawing.Size(500, 400);    
$FormInfUser.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink;    
$FormInfUser.AutoScaleDimensions = New-Object System.Drawing.SizeF(200, 100);   
#$Exbody.FormBorderStyle = 'Fixed3D'
$FormInfUser.MaximizeBox = $true;    
$FormInfUser.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
$FormInfUser.Text = “Информация о пользователе”; 
$FormInfUser.FormBorderStyle =[System.Windows.Forms.FormBorderStyle]::FixedSingle

$buttonaddFunction = New-Object System.Windows.Forms.Button
$buttonaddFunction.Dock = "bottom"
#$buttonaddFunction.Location = New-Object System.Drawing.Point(10,400)
$buttonaddFunction.Size = New-Object System.Drawing.Size(140, 50)
$buttonaddFunction.Text = "Добавить функции"
$buttonaddFunction.Enabled = $false
$FormInfUser.Controls.Add($buttonaddFunction)

$buttondelFunction = New-Object System.Windows.Forms.Button
$buttondelFunction.Dock = "bottom"
#$buttonaddFunction.Location = New-Object System.Drawing.Point(10,400)
$buttondelFunction.Size = New-Object System.Drawing.Size(140, 50)
$buttondelFunction.Text = "Удалить функции"
$buttondelFunction.Enabled = $false
$FormInfUser.Controls.Add($buttondelFunction)

$labelFIO = New-Object System.Windows.Forms.Label
$labelFIO.Size = New-Object System.Drawing.Size(200, 20)
$labelFIO.Location = New-Object System.Drawing.Point(10,10)
$labelFIO.Text = "ФИО:"
$FormInfUser.Controls.Add($labelFIO)

$labelUser = New-Object System.Windows.Forms.Textbox
$labelUser.Multiline = $true
$labelUser.BorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$labelUser.ReadOnly = $true
$labelUser.ScrollBars = "Vertical"
$labelUser.Size = New-Object System.Drawing.Size(200, 20)
$labelUser.Location = New-Object System.Drawing.Point(10,30)
$FormInfUser.Controls.Add($labelUser)

$labeltitle = New-Object System.Windows.Forms.Label
$labeltitle.Size = New-Object System.Drawing.Size(200, 20)
$labeltitle.Location = New-Object System.Drawing.Point(10,50)
$labeltitle.Text = "Должность:"
$FormInfUser.Controls.Add($labeltitle)

$labelUser2 = New-Object System.Windows.Forms.textbox
$labelUser2.Multiline = $true
$labelUser2.BorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$labelUser2.ReadOnly = $true
$labelUser2.ScrollBars = "Vertical"
$labelUser2.Size = New-Object System.Drawing.Size(200, 20)
$labelUser2.Location = New-Object System.Drawing.Point(10,70)
$FormInfUser.Controls.Add($labelUser2)
#
$labelUser3 = New-Object System.Windows.Forms.textbox
$labelUser3.Multiline = $true
$labelUser3.BorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$labelUser3.ReadOnly = $true
$labelUser3.ScrollBars = "Vertical"
$labelUser3.Size = New-Object System.Drawing.Size(200, 20)
$labelUser3.Location = New-Object System.Drawing.Point(10,110)
$FormInfUser.Controls.Add($labelUser3)
#
$labelphone = New-Object System.Windows.Forms.Label
$labelphone.Size = New-Object System.Drawing.Size(200, 20)
$labelphone.Location = New-Object System.Drawing.Point(10,90)
$labelphone.Text = "Рабочий телефон:"
$FormInfUser.Controls.Add($labelphone)

$labelUser4 = New-Object System.Windows.Forms.textbox
$labelUser4.Multiline = $true
$labelUser4.BorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$labelUser4.ReadOnly = $true
$labelUser4.ScrollBars = "Vertical"
$labelUser4.Size = New-Object System.Drawing.Size(200, 20)
$labelUser4.Location = New-Object System.Drawing.Point(10,150)
$FormInfUser.Controls.Add($labelUser4)
#

$labeldepart = New-Object System.Windows.Forms.Label
$labeldepart.Size = New-Object System.Drawing.Size(200, 20)
$labeldepart.Location = New-Object System.Drawing.Point(10,130)
$labeldepart.Text = "Отдел:"
$FormInfUser.Controls.Add($labeldepart)

$labelUser5 = New-Object System.Windows.Forms.textbox
$labelUser5.Multiline = $true
$labelUser5.BorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$labelUser5.ReadOnly = $true
$labelUser5.ScrollBars = "Vertical"
$labelUser5.Size = New-Object System.Drawing.Size(200, 20)
$labelUser5.Location = New-Object System.Drawing.Point(10,190)
$FormInfUser.Controls.Add($labelUser5)

$labelcompany = New-Object System.Windows.Forms.Label
$labelcompany.Size = New-Object System.Drawing.Size(200, 20)
$labelcompany.Location = New-Object System.Drawing.Point(10,170)
$labelcompany.Text = "Компания:"
$FormInfUser.Controls.Add($labelcompany)

$groupboxInfUser = New-Object System.Windows.Forms.GroupBox
$groupboxInfUser.Size = New-Object System.Drawing.Size(250, 250)
$groupboxInfUser.Location = New-Object System.Drawing.Point(240,10)
$groupboxInfUser.BackColor = "white"
$groupboxInfUser.BackColor = [System.Drawing.Color]::WhiteSmoke
$groupboxInfUser.Text = "Функции:"
$FormInfUser.Controls.Add($groupboxInfUser)

$checkboxActiveSync = New-Object System.Windows.Forms.CheckBox
$checkboxActiveSync.Size = New-Object System.Drawing.Size(200, 20)
$checkboxActiveSync.Text = "Exchange ActiveSync"
$checkboxActiveSync.Location = New-Object System.Drawing.Point(10,30)
$groupboxInfUser.Controls.Add($checkboxActiveSync)


$checkboxOWA = New-Object System.Windows.Forms.CheckBox
$checkboxOWA.Size = New-Object System.Drawing.Size(200, 20)
$checkboxOWA.Text = "Outlook Web App"
$checkboxOWA.Location = New-Object System.Drawing.Point(10,70)
$groupboxInfUser.Controls.Add($checkboxOWA)


$checkboxOWAdevice = New-Object System.Windows.Forms.CheckBox
$checkboxOWAdevice.Size = New-Object System.Drawing.Size(200, 20)
$checkboxOWAdevice.Text = "Outlook Web App для устройств"
$checkboxOWAdevice.Location = New-Object System.Drawing.Point(10,110)
$groupboxInfUser.Controls.Add($checkboxOWAdevice)

$checkboxActiveSync.add_CheckedChanged({
    if ($checkboxActiveSync.Checked -or $checkboxOWA.Checked -or $checkboxOWAdevice.Checked)
	{
	
	  $buttonaddFunction.Enabled = $true
	  $buttondelFunction.Enabled = $true
	}
	elseif ($checkboxActiveSync.Checked -eq $false -and $checkboxOWA.Checked -eq $false -and $checkboxOWAdevice.Checked -eq $false)
	{
	  $buttonaddFunction.Enabled = $false
	  $buttondelFunction.Enabled = $false
	}
})

$checkboxOWA.add_CheckedChanged({
   if ($checkboxActiveSync.Checked -or $checkboxOWA.Checked -or $checkboxOWAdevice.Checked)
	{
	  $buttonaddFunction.Enabled = $true
	  $buttondelFunction.Enabled = $true
	}
	elseif ($checkboxActiveSync.Checked -eq $false -and $checkboxOWA.Checked -eq $false -and $checkboxOWAdevice.Checked -eq $false)
	{
	  $buttonaddFunction.Enabled = $false
	  $buttondelFunction.Enabled = $false
	}
	
})

$checkboxOWAdevice.add_CheckedChanged({
     if ($checkboxActiveSync.Checked -or $checkboxOWA.Checked -or $checkboxOWAdevice.Checked)
	{
	  $buttonaddFunction.Enabled = $true
	  $buttondelFunction.Enabled = $true
	}
	elseif ($checkboxActiveSync.Checked -eq $false -and $checkboxOWA.Checked -eq $false -and $checkboxOWAdevice.Checked -eq $false)
	{
	  $buttonaddFunction.Enabled = $false
	  $buttondelFunction.Enabled = $false
	}
})



####функция проверки функций :)))
function FunctionUser {
if ($listMB.Items[$listMB.FocusedItem.Index].Subitems[2].Text)
{
$userget = Get-ADUser $listMB.Items[$listMB.FocusedItem.Index].subitems[2].Text -Properties * | select department,telephonenumber,title,company


$labelUser.Text = $listMB.Items[$listMB.FocusedItem.Index].Subitems[1].Text
if ($userget.title) {$labelUser2.Text = $userget.title} else {$labelUser2.Text = "Нет данных"}
if ($userget.telephonenumber) {$labelUser3.Text = $userget.telephonenumber} else {$labelUser3.Text = "Нет данных"}
if ($userget.department) {$labelUser4.Text = $userget.department} else {$labelUser4.Text = "Нет данных"}
if ($userget.company) {$labelUser5.Text = $userget.company} else {$labelUser5.Text = "Нет данных"}


$infuser = $listMB.Items[$listMB.FocusedItem.Index].Subitems[2].Text 
 
  
   $infocas =  Get-CASMailbox -Identity $infuser | select activesyncenabled,owaenabled,OWAforDevicesEnabled
   
   if ($infocas.activesyncenabled -eq "true")
   {
   $checkboxActiveSync.BackColor = "red"
   }
   else
   {
      $checkboxActiveSync.BackColor = "white"
   }
   if ($infocas.owaenabled -eq "true")
   {
      $checkboxOWA.BackColor = "red"
   }
   else
   {
      $checkboxOWA.BackColor = "white"
   }
   if ($infocas.OWAforDevicesEnabled -eq "true")
   {
     $checkboxOWAdevice.BackColor = "red"
   }
   else
   {
     $checkboxOWAdevice.BackColor = "white"
   }
  }
 
}

####Кнопки удаления и добавления функций
$buttonaddFunction.add_Click({
$OUTPUT = [System.Windows.Forms.MessageBox]::Show("Вы уверены в данных действиях?","Внимание",4)
				if ($OUTPUT -eq "YES")
{

if ($checkboxActiveSync.Checked)
{
   if ($checkboxActiveSync.BackColor -ne "Red")
   { 
    Set-CASMailbox -Identity $listMb.Items[$ListMB.FocusedItem.Index].subitems[2].Text -ActiveSyncEnabled $true
	$checkboxActiveSync.BackColor = "red"
   }
  
}

 if ($checkboxOWA.Checked)
{
    if ($checkboxOWA.BackColor -ne "Red")
   { 
      Set-CASMailbox -Identity $listMb.Items[$ListMB.FocusedItem.Index].subitems[2].Text -OWAEnabled $true 
	 $checkboxOWA.BackColor = "red"
	   
   }
}

 if ( $checkboxOWAdevice.Checked)
  {
    if ( $checkboxOWAdevice.BackColor -ne "Red")
   { 
      Set-CASMailbox -Identity $listMb.Items[$ListMB.FocusedItem.Index].subitems[2].Text -OWAForDevicesEnabled $true 
	$checkboxOWAdevice.BackColor = "red"
   }
  
       
   }
   $checkboxOWA.Checked = $false
   $checkboxOWAdevice.Checked = $false
   $checkboxActiveSync.Checked = $false
   }
})

$buttondelFunction.add_Click({
$OUTPUT = [System.Windows.Forms.MessageBox]::Show("Вы уверены в данных действиях?","Внимание",4)
				if ($OUTPUT -eq "YES")
{
if ($checkboxActiveSync.Checked)
{
	if ($checkboxActiveSync.BackColor -eq "Red")
  	 { 
      Set-CASMailbox -Identity $listMb.Items[$ListMB.FocusedItem.Index].subitems[2].Text -ActiveSyncEnabled $false
	  $checkboxActiveSync.BackColor = "white"
  	 }
 }  
 
 if ($checkboxOWA.Checked)
{
    if ($checkboxOWA.BackColor -eq "Red")
   { 
      Set-CASMailbox -Identity $listMb.Items[$ListMB.FocusedItem.Index].subitems[2].Text -OWAEnabled $false
	$checkboxOWA.BackColor = "white"
   }
  }
  
 if ($checkboxOWAdevice.Checked) 
 {
    if ( $checkboxOWAdevice.BackColor -eq "Red")
   { 
      Set-CASMailbox -Identity $listMb.Items[$ListMB.FocusedItem.Index].subitems[2].Text -OWAForDevicesEnabled $false
	  $checkboxOWAdevice.BackColor = "white"
	
   }
}


$checkboxActiveSync.Checked = $false
$checkboxOWAdevice.Checked = $false
$checkboxOWA.Checked = $false
}
}
)




