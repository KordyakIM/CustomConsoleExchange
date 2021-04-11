###########################################################
# Разработчик: Кордяк Иван Михайлович kordyakim@gmail.com #
###########################################################
$fromAddPermission = New-Object System.Windows.Forms.Form
$fromAddPermission.ClientSize = New-Object System.Drawing.Size(420, 300);    
$fromAddPermission.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink;    
$fromAddPermission.AutoScaleDimensions = New-Object System.Drawing.SizeF(200, 100);   
#$Exbody.FormBorderStyle = 'Fixed3D'
$fromAddPermission.MaximizeBox = $true;    
$fromAddPermission.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;



$buttonadd = New-Object System.Windows.Forms.Button
$buttonadd.Location = New-Object System.Drawing.Point(10,250)
$buttonadd.AutoSize = $true
$buttonadd.Text = "Добавить учетную запись"
$fromAddPermission.Controls.Add($buttonadd)


$comboboxCalendar = New-Object System.Windows.Forms.ComboBox
$comboboxCalendar.Location = New-Object System.Drawing.Point(10,200) 
$comboboxCalendar.Size = New-Object System.Drawing.Size(300,8)
$comboboxCalendar.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboboxCalendar.Visible = $false
$fromAddPermission.Controls.Add($comboboxCalendar)

          $comboboxCalendar.Items.Add("Owner")
		  $comboboxCalendar.Items.Add("Reviewer")
		  $comboboxCalendar.Items.Add("Editor")


$listSA = New-Object System.Windows.Forms.ListView
$listSA.Location = New-Object System.Drawing.Size(10,50)
$listSA.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$listSA.View = [System.Windows.Forms.View]::Details
$listSA.AutoSize = $true
$listSA.Size = New-Object System.Drawing.Size(400,120)
$fromAddPermission.Controls.Add($listSA)
#$listSA.LabelEdit = $true
$listSA.MultiSelect = $true
$LISTSA.FullRowSelect = $True
$listSA.AllowColumnReorder = $True
$listSA.GridLines = $true
$listSA.Columns.Add("SamAccountName")
$listSA.Columns.Add("Display Name")
$listSA.Columns[0].Width = 150
$listSA.Columns[1].Width = 245
$listSA.BackColor = 'Azure'
#
#$progressbar = New-Object System.Windows.Forms.ProgressBar
#$progressBar.Size = New-Object System.Drawing.Size(105,15)
#$progressbar.Location = New-Object System.Drawing.Point(10,270)
#$progressBar.ForeColor = "green"
#$progressBarName = "|||"
#$progressBar.Text = ""
#$fromaddpermission.Controls.Add($progressBar)


$textBox11 = New-Object System.Windows.Forms.TextBox;    
$textBox11.Location = New-Object System.Drawing.Point(10, 20);    
$textBox11.Size = New-Object System.Drawing.Size(400, 10);    
$textBox11.Name = "textBox0";      
$fromAddPermission.Controls.Add($textBox11)
#$textBox0.Text = "user";
$textBox0name = 'Введите пожалуйста имя Пользователя!'
$textbox11.ForeColor = 'LightGray'
$textBox11.Text = $textBox0name
$textBox11.add_Click({
	if($textBox11.Text -eq $textBox0name)
    {
        #Clear the text
        $textBox11.Text = ""
        $textBox11.ForeColor = 'WindowText'
    }
	if($textBox11.Text -eq $textBox11.Tag)
    {
        #Clear the text
        $textBox11.Text = ""
        $textBox11.ForeColor = 'WindowText'
    }
	
	})
$textBox11.add_KeyPress({
if($textBox11.Visible -and $textBox11.Tag -eq $null)
    {
        #Initialize the watermark and save it in the Tag property
        $textBox11.Tag = $textBox11.Text;
        $textBox11.ForeColor = 'LightGray'
        #If we have focus then clear out the text
        if($textBox11.Focused)
        {
            $textBox11.Text = ""
            $textBox11.ForeColor = 'WindowText'
        }
    }
})
$textBox11.add_Leave({
if($textBox1.Text -eq "")
    {
        #Display the watermark
        $textBox11.Text = $textBox0name
        $textBox11.ForeColor = 'LightGray'
    }
	if($textBox11.Text -eq "")
    {
        #Display the watermark
        $textBox11.Text = $textBox1.Tag
        $textBox11.ForeColor = 'LightGray'
    }
		})

function Fill-Listad ($Mask = "*") {
	#$ObjListbox.Items.Clear()
	$listSA.Items.Clear()
	$s = $textBox11.Text + "*"
    $str = Get-ADUser -Filter {SamAccountName -like $s } -Searchbase "DC=npr,DC=nornick,DC=ru" | select SamAccountName
		if($str -eq $null){
		$str = Get-ADUser -Filter { DisplayName -like $s } -Searchbase "DC=npr,DC=nornick,DC=ru" | select SamAccountName
	}

	if ($str -eq $null)
	{
#		$I = $objListBox.Items.Add("Пользователь не найден...") | Out-Null
		$listSA.Items.Add("Пользователь не найден...")
		$i = $listSA.Items[0].SubItems.Add("Пользователь не найден...") | Out-Null
	}
	else
	{
			foreach ( $item in $str ) 
		    {
		    	$s1 = $item -split "}"
				#Добавляем элемент в список
			    $string = $s1[0].Substring(17)
				$listSA.Items.Add($string)
	
			    $str1 = (Get-ADUser $string -Properties *).DisplayName
					if ($str1 -ne $null)
					{
			 			 $listSA.Items[$listSA.SelectedItems.Count-1].SubItems.Add($str1)	
					}
			
			}
		
		
	}
	
	
	
	
}

if ($Search)
{
    #Добавляем обработчик на событие TextChanged, который выполняет функцию Fill-List
    $textBox11.add_TextChanged({Fill-Listadd ("*" + $textBox11.Text + "*")})
}
else #Ищем только при нажатии Enter
{
    #Скриптблок (кусок исполняемого кода) который будет выполнен при нажатии клавиши в поле поиска
    $SB_KeyPress1 = {
        #Если была нажата клавиша Enter (13) то...
        if (13 -eq $_.keychar)
        {
		
		 if ((!$textBox11.Text) -or ( $textBox11.Text -eq " ") )
			{
				 #Вызываем функцию Fill-List
			[System.Windows.Forms.MessageBox]::Show("Введите пожалуйста имя пользователя!")
				return
		      
		 	 }
		 else
        	 {  
				Fill-Listad ("*" + $textBox11.Text + "*")
        	 }
		}
    }
    #Добавляем обработчик на событие KeyPress, указав в качестве выполняемого кода $SB_KeyPress
    $textBox11.add_KeyPress($SB_KeyPress1)
}


$listSA.add_ItemSelectionChanged({
$i = 0
if ($indexform -eq 3)
{

if($listSA.SelectedItems.Count -cge 1)
{

  if ($ListSA.FocusedItem)
  {
  
foreach ($a in $getuser)
{
Write-Host $a 
}
 if ($a -match $ListSA.Items[$listSA.FocusedItem.Index].Subitems[1].Text )
   {
       $buttonadd.Enabled = $true
	   $buttonadd.Text = "Удалить учетную запись"
   }
   else
 {
      $buttonadd.Enabled = $true
	  $buttonadd.Text = "Добавить учетную запись"
	  
   }


  
   }
 }
   

}
})


$buttonadd.add_Click({

$OUTPUT = [System.Windows.Forms.MessageBox]::Show("Вы уверены в данных действиях?","Внимание",4)
				if ($OUTPUT -eq "YES")
{

			if (!$listSA.FocusedItem)
			{
					[System.Windows.Forms.MessageBox]::Show("Выберите пожалуйста пользователя")
					return
			}
				try
				{
					if ($indexform -eq 1)
					{

			    		if (Get-Mailbox  $listSA.Items[$listSA.FocusedItem.Index].Text)
						{
				  		 Add-MailboxPermission -Identity $ListMB.Items[$ListMB.Focuseditem.Index].SubItems[2].Text -User $listSA.Items[$listSA.FocusedItem.Index].Text -AccessRights "FullAccess" -InheritanceType all -ErrorAction STOP
						}
						else
						{   
			      		 Add-MailboxPermission -Identity $ListMB.Items[$ListMB.Focuseditem.Index].SubItems[1].Text -User $listSA.Items[$listSA.FocusedItem.Index].Text -AccessRights "FullAccess" -ErrorAction STOP
						}
						[System.Windows.Forms.MessageBox]::Show("Пользователь добавлен!")
						$fromAddPermission.Close()
					}

				}
			catch{
			[System.Windows.Forms.MessageBox]::Show("Ошибка!")
			}

			try
			{
				if ($indexform -eq 0)
					{
					
			     			if (Get-Mailbox  $listSA.Items[$listSA.FocusedItem.Index].Text)
								{
				   				Add-ADPermission -Identity $ListMB.Items[$ListMB.Focuseditem.Index].SubItems[1].Text -User $listSA.Items[$listSA.Focuseditem.Index].Text -ExtendedRights "Send AS" -ErrorAction STOP
								}
							else
								{   
			      				 Add-ADPermission -Identity $ListMB.Items[$ListMB.Focuseditem.Index].SubItems[1].Text -User $listSA.Items[$listSA.Focused.Indexitem].Text -ExtendedRights "Send AS" -ErrorAction STOP
								}
					
				[System.Windows.Forms.MessageBox]::Show("Пользователь добавлен!")
				$fromAddPermission.Close()
					}
			}
					
			catch
			{
				[System.Windows.Forms.MessageBox]::Show("Ошибка!")
			}



			try
			{
				if ($indexform -eq 2)
			{
			if ((Get-MailboxPermission -Identity $ListMB.Items[$ListMB.Focuseditem.Index].SubItems[1].Text -User $listSA.Items[$listSA.FocusedItem.Index].Text  | Format-Table -Property * -AutoSize | Out-String -Width 1024) -like "*FullAccess*")
											   {
											       Remove-MailboxPermission -Identity $ListMB.Items[$ListMB.Focuseditem.Index].SubItems[1].Text -User $listSA.Items[$listSA.FocusedItem.Index].Text -AccessRights FullAccess  -Confirm:$false
												   [System.Windows.Forms.MessageBox]::Show("Пользователь удален!")	
											       
											   }
											   else
											   {
											       Remove-ADPermission -Identity $ListMB.Items[$ListMB.Focuseditem.Index].SubItems[1].Text -User $listSA.Items[$listSA.FocusedItem.Index].Text -ExtendedRights "Send As","Receive AS"  -Confirm:$false
											   [System.Windows.Forms.MessageBox]::Show("Пользователь удален!")	
											   }
									  					
									        
									$fromAddPermission.Close()  		 
									
			}


				}
				catch 
				{
				  [System.Windows.Forms.MessageBox]::Show("Ошибка!")
				}
	
	
	####Доступ к календарю
			try
			{
				if ($indexform -eq 3)
			{
			       if ($comboboxCalendar.Text)
				{
#				$calendar2 = $ListMB.Items[$ListMB.Focuseditem.Index].SubItems[2].Text
				$calendar = $ListMB.Items[$ListMB.Focuseditem.Index].SubItems[3].Text+":\Календарь"
				   
					Add-MailboxFolderPermission -Identity $calendar -User $listSA.Items[$listSA.FocusedItem.Index].Text -AccessRights $comboboxCalendar.Text
					[System.Windows.Forms.MessageBox]::Show("Пользователь добавлен!")
					
				}
					else
					{
					[System.Windows.Forms.MessageBox]::Show("Не выбрана роль доступа!")
					}
			
			}


				}
				catch 
				{
				  [System.Windows.Forms.MessageBox]::Show("$error")
				}
	
	try
			{
				if ($indexform -eq 4)
			{
			   $calendar = $ListMB.Items[$ListMB.Focuseditem.Index].SubItems[3].Text+":\Календарь"
			Remove-MailboxFolderPermission -Identity $calendar -User $listSA.Items[$listSA.FocusedItem.Index].Text 	-Confirm:$false
			[System.Windows.Forms.MessageBox]::Show("Пользователь удален!")
			
			}


				}
				catch 
				{
				  [System.Windows.Forms.MessageBox]::Show("$Error")
				}

}





})


