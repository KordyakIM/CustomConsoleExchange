###########################################################
# Разработчик: Кордяк Иван Михайлович kordyakim@gmail.com #
###########################################################
$excel = New-Object -com excel.application
$ListMTL = New-Object System.Windows.Forms.ListView
$ListMTL.dock = "Fill"
$ListMTL.Height = 200  
$ListMTL.View = "Details"
$ListMTL.MultiSelect = $true
$ListMTL.FullRowSelect = $True
$ListMTL.AutoSize = $true
$ListMTL.LabelEdit = $True
$ListMTL.AllowColumnReorder = $True
$ListMTL.Scrollable = $true
$ListMTL.GridLines = $true
$ListMTL.Columns.Add("TimeStamp",160)
$ListMTL.Columns.Add("TotalMilliseconds",160)
$ListMTL.Columns.Add("MessageLatency",160)
$ListMTL.Columns.Add("TotalBytes",90)
$ListMTL.Columns.Add("TransportService",90)
$ListMTL.Columns.Add("Sender",90)
$ListMTL.Columns.Add("Recipients",90)
$ListMTL.Columns.Add("MessageSubject",90)
$ListMTL.Columns.Add("ClientHostname",90)
$ListMTL.Columns.Add("ServerHostname",90)
$ListMTL.Columns.Add("ConnectorId",90)
$ListMTL.Columns.Add("Source",90)
$ListMTL.Columns.Add("EventID",90)
$ListMTL.Columns.Add("SourceContext",90)
$ListMTL.Columns.Add("RecipientStatus",90)
$ListMTL.Columns.Add("OriginalClientIP",90)
$ListMTL.add_ColumnClick({
	if ($ListMTL.Items.Count -gt 1){
		SortListTwoviewMTL $_.Column
	}
})

[array]$SelectItems = "timestamp","MessageLatency","TotalBytes","messageid","sender","recipients","messagesubject","ClientHostname","ServerHostname","ConnectorId","Source","eventid","SourceContext","RecipientStatus","OriginalClientIp"
[array]$SelectItemsEasy = "timestamp","TotalBytes","sender","recipients","messagesubject","Source","eventid"

$ListMTL.add_ItemSelectionChanged({
$ListMTL.add_ItemMouseHover({
						if($ListMTL.SelectedItems.Count -cge 1){
							$menucopyMTL.Enabled = $true
							$menucopyMTL1.Enabled = $true
							$menucopyMTL2.Enabled = $true
							$menucopyMTL01.Enabled = $true
							$menucopyMTL02.Enabled = $true
								if($listMB1.CheckedItems.Count -eq 1){
									$menucopyMTL1.Enabled = $true
									$menucopyMTL01.Enabled = $true
								}else{
									$menucopyMTL1.Enabled = $false
									$menucopyMTL01.Enabled = $false
								}
						}else{
							$menucopyMTL.Enabled = $false
							$menucopyMTL1.Enabled = $false
							$menucopyMTL2.Enabled = $false
							$menucopyMTL01.Enabled = $false
							$menucopyMTL02.Enabled = $false
							$menucopyMTL1.Enabled = $false
							$menucopyMTL01.Enabled = $false
						}
					})
		if($ListMTL.SelectedItems.Count -cge 1){
					
			$menucopyMTL.Enabled = $true
			$menucopyMTL1.Enabled = $true
			$menucopyMTL2.Enabled = $true
			$menucopyMTL01.Enabled = $true
			$menucopyMTL02.Enabled = $true
				if($listMB1.CheckedItems.Count -eq 1){
					$menucopyMTL1.Enabled = $true
					$menucopyMTL01.Enabled = $true
				}else{
					$menucopyMTL1.Enabled = $false
					$menucopyMTL01.Enabled = $false
				}
		}else{
			$menucopyMTL.Enabled = $false
			$menucopyMTL1.Enabled = $false
			$menucopyMTL2.Enabled = $false
			$menucopyMTL01.Enabled = $false
			$menucopyMTL02.Enabled = $false
			$menucopyMTL1.Enabled = $false
			$menucopyMTL01.Enabled = $false
		}
})
$ListMTL.add_KeyDown({
	param($sender, $e)
	if ($_.KeyCode -eq "C" -and $e.Control){
		Set-ClipBoardMTL
	}
	if ($_.keycode -eq "A" -and $e.Control){
		foreach ($ListItem in $ListMTL.Items){
		    $ListItem.selected = $true
		}
	}
})
$tabControlP2.Controls.add($ListMTL)

#group for button--------------------------------------------------------------------
$GroupInsideMT = New-Object System.Windows.Forms.GroupBox
$GroupInsideMT.dock = "top"
$GroupInsideMT.Height = 40
$tabControlP2.Controls.Add($GroupInsideMT);
#Easy-----------------------------------------------------------------------------------------
$btnExcel = New-Object System.Windows.Forms.Button; 
$btnExcel.dock = "Left"
$btnExcel.Height = 40  
$btnExcel.Size = New-Object System.Drawing.Size(90, 25);    
$btnExcel.Name = "btnExcel";    
$btnExcel.Text = "Excel"; 
$btnExcel.Enabled = $true;    
$btnExcel.Add_Click({
#	if($ListMTL.Items.Count -ige 1){
		$excel.Visible = $True
		$Workbook = $excel.Workbooks.Add()
		#Соединяемся к worksheet, меняем имя и делаем активным
		$serverInfoSheet = $workbook.Worksheets.Item(1)
		$serverInfoSheet.Name = 'Message Tracker info'
		$serverInfoSheet.application.activewindow.splitrow = 1
		$serverInfoSheet.application.activewindow.freezepanes = $true
		$serverInfoSheet.Activate() | Out-Null
		$row = 1
		$Column = 1
		while($ListMTL.get_Columns().text[$column-1]){
			$serverInfoSheet.Cells.Item($row,$column)= $ListMTL.get_Columns().text[$column-1]
			$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex = 48
			$Column++
		}
		$row++
    	foreach ( $ListItem in $ListMTL.Items) {
	    	#Добавляем элемент в список
			$Column = 1
			while($ListItem.subitems[$column-1]){
	    		try{$serverInfoSheet.Cells.Item($row,$column)=($ListItem.subitems[$Column-1].text)}catch{}
				$Column++
			}
			$row++
		}									
		#включения фильтра
     	$excel.Selection.AutoFilter(1)
		$excel.Top
		$excel.DisplayFullScreen = $true
#	}
})
$GroupInsideMT.Controls.Add($btnExcel);
#Easy-----------------------------------------------------------------------------------------
$btnEasy = New-Object System.Windows.Forms.Button; 
$btnEasy.dock = "Left"
$btnEasy.Height = 40  
$btnEasy.Size = New-Object System.Drawing.Size(90, 25);    
$btnEasy.Name = "btn1";    
$btnEasy.Text = "Easy Отчет"; 
$btnEasy.Enabled = $true;    
$btnEasy.Add_Click({
	export-messagetracker-easy
})
$GroupInsideMT.Controls.Add($btnEasy);
#-----------------------------------------------------------------------------------------
$btn = New-Object System.Windows.Forms.Button; 
$btn.dock = "Left"
$btn.Height = 40  
$btn.Size = New-Object System.Drawing.Size(90, 25);    
$btn.Name = "btn1";    
$btn.Text = "Full Отчет"; 
$btn.Enabled = $true;    
$btn.Add_Click({
	export-messagetracker
})
$GroupInsideMT.Controls.Add($btn);
#чек бокс------------------
$CB = New-Object System.Windows.Forms.CheckBox;    
$CB.dock = "Right"  
$CB.Size = New-Object System.Drawing.Size(100, 25);    
$CBname = "Recipients"; 
$CB.Text = $CBname;
$CB.Enabled = $false;
$GroupInsideMT.Controls.Add($CB);

					#строка запроса времени END-------------------------------------------------------
#Надпись временем запроса Start
$dash = New-Object System.Windows.Forms.Label
$dash.dock = "Right"
$dash.Size = New-Object System.Drawing.Size(35,25) 
$dash.ForeColor = "Blue"
$dash.Text = "Start:"
$GroupInsideMT.Controls.Add($dash)
#---------------------------------------------------

			#строка запроса времени START-------------------------------------------------------
			$textBox2 = New-Object System.Windows.Forms.TextBox;    
			$textBox2.dock = "Right"  
			$textBox2.Size = New-Object System.Drawing.Size(55, 10);    
			$textBox2.Name = "textBox2";      
			$textBox2name = 'Месяц'
			$textBox2.ForeColor = 'LightGray'
			$textBox2.Text = $textBox2name
			$textBox2_AddGM = 0;
			$textBox2.add_Click({
				if($textBox2.Text -eq $textBox2name)
			    {
			        #Clear the text
			        $textBox2.Text = ""
			        $textBox2.ForeColor = 'WindowText'
			    }
				if($textBox2.Text -eq $textBox2.Tag)
			    {
			        #Clear the text
			        $textBox2.Text = ""
			        $textBox2.ForeColor = 'WindowText'
			    }
				})
			$textBox2.add_KeyPress({
			if($textBox2.Visible -and $textBox2.Tag -eq $null)
			    {
			        #Initialize the watermark and save it in the Tag property
			        $textBox2.Tag = $textBox2.Text;
			        $textBox2.ForeColor = 'LightGray'
			        #If we have focus then clear out the text
			        if($textBox2.Focused)
			        {
			            $textBox2.Text = ""
			            $textBox2.ForeColor = 'WindowText'
			        }
			    }
			})
			$textBox2.add_Leave({
			if($textBox2.Text -eq "")
			    {
			        #Display the watermark
			        $textBox2.Text = $textBox2name
			        $textBox2.ForeColor = 'LightGray'
			    }
				if($textBox2.Text -eq "")
			    {
			        #Display the watermark
			        $textBox2.Text = $textBox2.Tag
			        $textBox2.ForeColor = 'LightGray'
			    }
					})
			$GroupInsideMT.Controls.Add($textBox2);
			
					$textBox1 = New-Object System.Windows.Forms.TextBox;    
					$textBox1.dock = "Right" 
					$textBox1.Size = New-Object System.Drawing.Size(55, 10);    
					$textBox1.Name = "textBox1";      
					$textBox1name = 'День'
					$textBox1.ForeColor = 'LightGray'
					$textBox1.Text = $textBox1name
					$textBox1_AddGM = 0;
					$textBox1.add_Click({
						if($textBox1.Text -eq $textBox1name)
					    {
					        #Clear the text
					        $textBox1.Text = ""
					        $textBox1.ForeColor = 'WindowText'
					    }
						if($textBox1.Text -eq $textBox1.Tag)
					    {
					        #Clear the text
					        $textBox1.Text = ""
					        $textBox1.ForeColor = 'WindowText'
					    }
						})
					$textBox1.add_KeyPress({
					if($textBox1.Visible -and $textBox1.Tag -eq $null)
					    {
					        #Initialize the watermark and save it in the Tag property
					        $textBox1.Tag = $textBox1.Text;
					        $textBox1.ForeColor = 'LightGray'
					        #If we have focus then clear out the text
					        if($textBox1.Focused)
					        {
					            $textBox1.Text = ""
					            $textBox1.ForeColor = 'WindowText'
					        }
					    }
					})
					$textBox1.add_Leave({
					if($textBox1.Text -eq "")
					    {
					        #Display the watermark
					        $textBox1.Text = $textBox1name
					        $textBox1.ForeColor = 'LightGray'
					    }
						if($textBox1.Text -eq "")
					    {
					        #Display the watermark
					        $textBox1.Text = $textBox1.Tag
					        $textBox1.ForeColor = 'LightGray'
					    }
							})
					$GroupInsideMT.Controls.Add($textBox1);
			
							$textBoxYear = New-Object System.Windows.Forms.TextBox;    
							$textBoxYear.dock = "Right"  
							$textBoxYear.Size = New-Object System.Drawing.Size(55, 10);    
							$textBoxYear.Name = "textBox";      
							$textBoxYearname = 'Год'
							$textBoxYear.ForeColor = 'LightGray'
							$textBoxYear.Text = $textBoxYearname
							$textBoxYear_AddGM = 0;
							$textBoxYear.add_Click({
								if($textBoxYear.Text -eq $textBoxYearname)
							    {
							        #Clear the text
							        $textBoxYear.Text = ""
							        $textBoxYear.ForeColor = 'WindowText'
							    }
								if($textBoxYear.Text -eq $textBoxYear.Tag)
							    {
							        #Clear the text
							        $textBoxYear.Text = ""
							        $textBoxYear.ForeColor = 'WindowText'
							    }
								})
							$textBoxYear.add_KeyPress({
							if($textBoxYear.Visible -and $textBoxYear.Tag -eq $null)
							    {
							        #Initialize the watermark and save it in the Tag property
							        $textBoxYear.Tag = $textBoxYear.Text;
							        $textBoxYear.ForeColor = 'LightGray'
							        #If we have focus then clear out the text
							        if($textBoxYear.Focused)
							        {
							            $textBoxYear.Text = ""
							            $textBoxYear.ForeColor = 'WindowText'
							        }
							    }
							})
							$textBoxYear.add_Leave({
							if($textBoxYear.Text -eq "")
							    {
							        #Display the watermark
							        $textBoxYear.Text = $textBoxYearname
							        $textBoxYear.ForeColor = 'LightGray'
							    }
								if($textBoxYear.Text -eq "")
							    {
							        #Display the watermark
							        $textBoxYear.Text = $textBoxYear.Tag
							        $textBoxYear.ForeColor = 'LightGray'
							    }
									})
							$GroupInsideMT.Controls.Add($textBoxYear);
							
#-----------------------------------------------------------------------------------	
				
#Надпись временем запроса END
$dash = New-Object System.Windows.Forms.Label
$dash.dock = "Right"
$dash.Size = New-Object System.Drawing.Size(35,25) 
$dash.ForeColor = "Blue"
$dash.Text = "END:"
$GroupInsideMT.Controls.Add($dash)

									$textBoxEND2 = New-Object System.Windows.Forms.TextBox;    
									$textBoxEND2.dock = "Right"
									$textBoxEND2.Size = New-Object System.Drawing.Size(55, 10);    
									$textBoxEND2.Name = "$textBoxEND2";      
									$textBoxEND2name = 'Месяц'
									$textBoxEND2.ForeColor = 'LightGray'
									$textBoxEND2.Text = $textBoxEND2name
									$textBoxEND2_AddGM = 0;
									$textBoxEND2.add_Click({
										if($textBoxEND2.Text -eq $textBoxEND2name)
									    {
									        #Clear the text
									        $textBoxEND2.Text = ""
									        $textBoxEND2.ForeColor = 'WindowText'
									    }
										if($textBoxEND2.Text -eq $textBoxEND2.Tag)
									    {
									        #Clear the text
									        $textBoxEND2.Text = ""
									        $textBoxEND2.ForeColor = 'WindowText'
									    }
										})
									$textBoxEND2.add_KeyPress({
									if($textBoxEND2.Visible -and $textBoxEND2.Tag -eq $null)
									    {
									        #Initialize the watermark and save it in the Tag property
									        $textBoxEND2.Tag = $textBoxEND2.Text;
									        $textBoxEND2.ForeColor = 'LightGray'
									        #If we have focus then clear out the text
									        if($textBoxEND2.Focused)
									        {
									            $textBoxEND2.Text = ""
									            $textBoxEND2.ForeColor = 'WindowText'
									        }
									    }
									})
									$textBoxEND2.add_Leave({
									if($textBoxEND2.Text -eq "")
									    {
									        #Display the watermark
									        $textBoxEND2.Text = $textBoxEND2name
									        $textBoxEND2.ForeColor = 'LightGray'
									    }
										if($textBoxEND2.Text -eq "")
									    {
									        #Display the watermark
									        $textBoxEND2.Text = $textBoxEND2.Tag
									        $textBoxEND2.ForeColor = 'LightGray'
									    }
											})
									$GroupInsideMT.Controls.Add($textBoxEND2);
									
							$textBoxEND1 = New-Object System.Windows.Forms.TextBox;    
							$textBoxEND1.dock = "Right"
							$textBoxEND1.Size = New-Object System.Drawing.Size(55, 10);    
							$textBoxEND1.Name = "$textBoxEND1";      
							$textBoxEND1name = 'День'
							$textBoxEND1.ForeColor = 'LightGray'
							$textBoxEND1.Text = $textBoxEND1name
							$textBoxEND1_AddGM = 0;
							$textBoxEND1.add_Click({
								if($textBoxEND1.Text -eq $textBoxEND1name)
							    {
							        #Clear the text
							        $textBoxEND1.Text = ""
							        $textBoxEND1.ForeColor = 'WindowText'
							    }
								if($textBoxEND1.Text -eq $textBoxEND1.Tag)
							    {
							        #Clear the text
							        $textBoxEND1.Text = ""
							        $textBoxEND1.ForeColor = 'WindowText'
							    }
								})
							$textBoxEND1.add_KeyPress({
							if($textBoxEND1.Visible -and $textBoxEND1.Tag -eq $null)
							    {
							        #Initialize the watermark and save it in the Tag property
							        $textBoxEND1.Tag = $textBoxEND1.Text;
							        $textBoxEND1.ForeColor = 'LightGray'
							        #If we have focus then clear out the text
							        if($textBoxEND1.Focused)
							        {
							            $textBoxEND1.Text = ""
							            $textBoxEND1.ForeColor = 'WindowText'
							        }
							    }
							})
							$textBoxEND1.add_Leave({
							if($textBoxEND1.Text -eq "")
							    {
							        #Display the watermark
							        $textBoxEND1.Text = $textBoxEND1name
							        $textBoxEND1.ForeColor = 'LightGray'
							    }
								if($textBoxEND1.Text -eq "")
							    {
							        #Display the watermark
							        $textBoxEND1.Text = $textBoxEND1.Tag
							        $textBoxEND1.ForeColor = 'LightGray'
							    }
									})
							$GroupInsideMT.Controls.Add($textBoxEND1);
									
					$textBoxEND = New-Object System.Windows.Forms.TextBox;    
					$textBoxEND.dock = "Right"
					$textBoxEND.Size = New-Object System.Drawing.Size(55, 10);    
					$textBoxEND.Name = "$textBoxEND";      
					$textBoxENDname = 'Год'
					$textBoxEND.ForeColor = 'LightGray'
					$textBoxEND.Text = $textBoxENDname
					$textBoxEND_AddGM = 0;
					$textBoxEND.add_Click({
						if($textBoxEND.Text -eq $textBoxENDname)
					    {
					        #Clear the text
					        $textBoxEND.Text = ""
					        $textBoxEND.ForeColor = 'WindowText'
					    }
						if($textBoxEND.Text -eq $textBoxEND.Tag)
					    {
					        #Clear the text
					        $textBoxEND.Text = ""
					        $textBoxEND.ForeColor = 'WindowText'
					    }
						})
					$textBoxEND.add_KeyPress({
					if($textBoxEND.Visible -and $textBoxEND.Tag -eq $null)
					    {
					        #Initialize the watermark and save it in the Tag property
					        $textBoxEND.Tag = $textBoxEND.Text;
					        $textBoxEND.ForeColor = 'LightGray'
					        #If we have focus then clear out the text
					        if($textBoxEND.Focused)
					        {
					            $textBoxEND.Text = ""
					            $textBoxEND.ForeColor = 'WindowText'
					        }
					    }
					})
					$textBoxEND.add_Leave({
					if($textBoxEND.Text -eq "")
					    {
					        #Display the watermark
					        $textBoxEND.Text = $textBoxENDname
					        $textBoxEND.ForeColor = 'LightGray'
					    }
						if($textBoxEND.Text -eq "")
					    {
					        #Display the watermark
					        $textBoxEND.Text = $textBoxEND.Tag
					        $textBoxEND.ForeColor = 'LightGray'
					    }
							})
					$GroupInsideMT.Controls.Add($textBoxEND);
				#------------------------------------------------------------------------------------
				
#строка запроса
$SearchOnType = $false
$textBoxMTL = New-Object System.Windows.Forms.TextBox;    
$textBoxMTL.dock = "right"
$textBoxMTL.Location = New-Object System.Drawing.Point(10, 10);    
$textBoxMTL.Size = New-Object System.Drawing.Size(200, 10);    
$textBoxMTL.Name = "textBox0";      
#$textBoxMTL.Text = "user";
$textBoxMTLname = 'Введите пожалуйста тему Письма!'
$textBoxMTL.ForeColor = 'LightGray'
$textBoxMTL.Text = $textBoxMTLname
$textBoxMTL_AddGM = 0;
$textBoxMTL.add_Click({
	if($textBoxMTL.Text -eq $textBoxMTLname)
    {
        #Clear the text
        $textBoxMTL.Text = ""
        $textBoxMTL.ForeColor = 'WindowText'
    }
	if($textBoxMTL.Text -eq $textBoxMTL.Tag)
    {
        #Clear the text
        $textBoxMTL.Text = ""
        $textBoxMTL.ForeColor = 'WindowText'
    }
	})
$textBoxMTL.add_KeyPress({
if($textBoxMTL.Visible -and $textBoxMTL.Tag -eq $null)
    {
        #Initialize the watermark and save it in the Tag property
        $textBoxMTL.Tag = $textBoxMTL.Text;
        $textBoxMTL.ForeColor = 'LightGray'
        #If we have focus then clear out the text
        if($textBoxMTL.Focused)
        {
            $textBoxMTL.Text = ""
            $textBoxMTL.ForeColor = 'WindowText'
        }
    }
})
$textBoxMTL.add_Leave({
if($textBoxMTL.Text -eq "")
    {
        #Display the watermark
        $textBoxMTL.Text = $textBoxMTLname
        $textBoxMTL.ForeColor = 'LightGray'
    }
	if($textBoxMTL.Text -eq "")
    {
        #Display the watermark
        $textBoxMTL.Text = $textBoxMTL.Tag
        $textBoxMTL.ForeColor = 'LightGray'
    }
		})
$GroupInsideMT.Controls.Add($textBoxMTL);
#------------------------------------------------------------------

#вызов окна выбора отчетов
function form {
$OUTPUT = [System.Windows.Forms.MessageBox]::Show("Хотите Full-ьный Отчет?
No - Easy Отчет","Отчет","YesNoCancel")
	if ($OUTPUT -eq "YES"){
		export-messagetracker
	}
	elseif($OUTPUT -eq "NO") {
		export-messagetracker-easy
	}
}

			#действие на нажатие клавиши Enter------------------------------------------------------
			$textBoxEND.Add_KeyDown({
			if ($_.KeyCode -eq "Enter") 
				    {
				    	form
				    }
			})
				$textBoxEND1.Add_KeyDown({
				if ($_.KeyCode -eq "Enter") 
					    {
					    	form
					    }
				})
					$textBoxEND2.Add_KeyDown({
					if ($_.KeyCode -eq "Enter") 
						    {
						    	form
						    }
					})
						$textBoxYear.Add_KeyDown({
						if ($_.KeyCode -eq "Enter") 
							    {
							    	form
							    }
						})
							$textBox1.Add_KeyDown({
							if ($_.KeyCode -eq "Enter") 
								    {
								    	form
								    }
							})
								$textBox2.Add_KeyDown({
								if ($_.KeyCode -eq "Enter") 
									    {
									    	form
									    }
								})
									$textBoxMTL.Add_KeyDown({
									if ($_.KeyCode -eq "Enter") 
										    {
										    	form
										    }
									})
			#--------------------------------------------------------
										function Add-ListMTLmessagetrackinglog {
											foreach ($MT in $MTss){
												#Добавляем элемент в список
												$MTt=$MT.timestamp
												$MTmllscd=$MT.timestamp.TimeOfDay.TotalMilliseconds
												$MTml=$MT.MessageLatency
												$MTb=$MT.TotalBytes
												$MTs=$MT.sender
												$MTr=$MT.recipients
												$MTmm=$MT.messagesubject
												$MTclt=$MT.ClientHostname
												$MTsrv=$MT.ServerHostname
												$MTcid=$MT.ConnectorId
												$MTsrc=$MT.Source
												$MTe=$MT.eventid
												$MTsc=$MT.SourceContext
												$MTrs=$MT.RecipientStatus
												$MTorigIp = $MT.originalClientIP
												if($MTe -notlike "ha*"){
													if("$MTt"){$E = $ListMTL.Items.Add("$MTt")}else{$E = $ListMTL.Items.Add("0")}
													if("$MTmllscd"){$E.SubItems.Add("$MTmllscd")}else{$E.SubItems.Add("0")}
													if("$MTml"){$E.SubItems.Add("$MTml")}else{$E.SubItems.Add("0")}
													if("$MTb"){$E.SubItems.Add("$MTb")}else{$E.SubItems.Add("0")}
													if("$transportservice"){$E.SubItems.Add("$transportservice")}else{$E.SubItems.Add("null")}
													if("$MTs"){$E.SubItems.Add("$MTs")}else{$E.SubItems.Add("null")}	
													if("$MTr"){$E.SubItems.Add("$MTr")}else{$E.SubItems.Add("null")}
													if("$MTmm"){$E.SubItems.Add("$MTmm")}else{$E.SubItems.Add("null")}
													if("$MTclt"){$E.SubItems.Add("$MTclt")}else{$E.SubItems.Add("null")}
													if("$MTsrv"){$E.SubItems.Add("$MTsrv")}else{$E.SubItems.Add("null")}
													if("$MTcid"){$E.SubItems.Add("$MTcid")}else{$E.SubItems.Add("null")}
													if("$MTsrc"){$E.SubItems.Add("$MTsrc")}else{$E.SubItems.Add("null")}
													if("$MTe"){$E.SubItems.Add("$MTe")}else{$E.SubItems.Add("null")}
													if("$MTsc"){$E.SubItems.Add("$MTsc")}else{$E.SubItems.Add("null")}
													if("$MTrs"){$E.SubItems.Add("$MTrs")}else{$E.SubItems.Add("null")}
													if("$MTorigIp"){$E.SubItems.Add("$MTorigIp")}else{$E.SubItems.Add("null")}
												}
											}
										}
				#--------------------------------------------------------
										function Add-ListMTLmessagetrackinglog-Easy {
											foreach ($MT in $MTss){
												#Добавляем элемент в список
												$MTt=$MT.timestamp
												$MTb=$MT.TotalBytes
												$MTs=$MT.sender
												$MTr=$MT.recipients
												$MTmm=$MT.messagesubject
												$MTsrc=$MT.Source
												$MTe=$MT.eventid
												if($MTe -notlike "ha*"){
													if("$MTt"){$E = $ListMTL.Items.Add("$MTt")}else{$E = $ListMTL.Items.Add("0")}
													$E.SubItems.Add("")
													$E.SubItems.Add("")
													if("$MTb"){$E.SubItems.Add("$MTb")}else{$E.SubItems.Add("0")}
													$E.SubItems.Add("")
													if("$MTs"){$E.SubItems.Add("$MTs")}else{$E.SubItems.Add("null")}	
													if("$MTr"){$E.SubItems.Add("$MTr")}else{$E.SubItems.Add("null")}
													if("$MTmm"){$E.SubItems.Add("$MTmm")}else{$E.SubItems.Add("null")}
													$E.SubItems.Add("")
													$E.SubItems.Add("")
													$E.SubItems.Add("")
													if("$MTsrc"){$E.SubItems.Add("$MTsrc")}else{$E.SubItems.Add("null")}
													if("$MTe"){$E.SubItems.Add("$MTe")}else{$E.SubItems.Add("null")}
													$E.SubItems.Add("")
													$E.SubItems.Add("")
													$E.SubItems.Add("")
												}
											}
										}
function export-messagetracker {
		$ListMTL.items.clear()
		$month = $textBox2.text
		$day= $textBox1.Text
		$year = $textBoxYear.text
		$mend = $textBoxEND2.text
		$dend = $textBoxEND1.Text
		$yend = $textBoxEND.text
#1
#почтовый ящик - да
#тема сообщения - да
#временная зона - да
#Recipients - да
		if(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -ne 'Введите пожалуйста тему Письма!')-and($textBox2.text -ne 'Месяц')-and($textBox1.text -ne 'День')-and($textBoxYear.text -ne 'Год')-and($textBoxEND2.text -ne 'Месяц')-and($textBoxEND1.text -ne 'День')-and($textBoxEND.text -ne 'Год')-and($CB.Checked)){
			$transportservices = (Get-TransportService).name
			foreach ($transportservice in $transportservices){
				$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Recipients $ListMB1.CheckedItems.subitems[2].text -start "$month/$day/$year 00:00:00" -end "$mend/$dend/$yend 23:59:00" -MessageSubject $textBoxMTL.Text | select-object $SelectItems
				Add-ListMTLmessagetrackinglog
			}
		}
#2
#почтовый ящик - да
#тема сообщения - да
#временная зона - да
#Recipients - нет
		elseif(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -ne 'Введите пожалуйста тему Письма!')-and($textBox2.text -ne 'Месяц')-and($textBox1.text -ne 'День')-and($textBoxYear.text -ne 'Год')-and($textBoxEND2.text -ne 'Месяц')-and($textBoxEND1.text -ne 'День')-and($textBoxEND.text -ne 'Год')-and(!$CB.Checked)){
			$transportservices = (Get-TransportService).name
			foreach ($transportservice in $transportservices){
				$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Sender $ListMB1.CheckedItems.subitems[2].text -start "$month/$day/$year 00:00:00" -end "$mend/$dend/$yend 23:59:00" -MessageSubject $textBoxMTL.Text | select-object $SelectItems
				Add-ListMTLmessagetrackinglog
			}
		}
#3
#почтовый ящик - да
#тема сообщения - нет
#временная зона - да
#Recipients - нет
		elseif(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -eq 'Введите пожалуйста тему Письма!')-and($textBox2.text -ne 'Месяц')-and($textBox1.text -ne 'День')-and($textBoxYear.text -ne 'Год')-and($textBoxEND2.text -ne 'Месяц')-and($textBoxEND1.text -ne 'День')-and($textBoxEND.text -ne 'Год')-and(!$CB.Checked)){
			$transportservices = (Get-TransportService).name
			foreach ($transportservice in $transportservices){
				$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Sender $ListMB1.CheckedItems.subitems[2].text -start "$month/$day/$year 00:00:00" -end "$mend/$dend/$yend 23:59:00" | select-object $SelectItems
				Add-ListMTLmessagetrackinglog
			}
		}
#4
#почтовый ящик - да
#тема сообщения - нет
#временная зона - нет
#Recipients - нет
		elseif(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -eq 'Введите пожалуйста тему Письма!')-and($textBox2.text -eq 'Месяц')-and($textBox1.text -eq 'День')-and($textBoxYear.text -eq 'Год')-and($textBoxEND2.text -eq 'Месяц')-and($textBoxEND1.text -eq 'День')-and($textBoxEND.text -eq 'Год')-and(!$CB.Checked)){
			$transportservices = (Get-TransportService).name
			foreach ($transportservice in $transportservices){
				$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Sender $ListMB1.CheckedItems.subitems[2].text | select-object $SelectItems
				Add-ListMTLmessagetrackinglog
			}
		}
#5
#почтовый ящик - нет
#тема сообщения - да
#временная зона - нет
#Recipients - нет
		elseif((!$ListMB1.CheckedItems.Count)-and($textBoxMTL.Text -ne 'Введите пожалуйста тему Письма!')-and($textBox2.text -eq 'Месяц')-and($textBox1.text -eq 'День')-and($textBoxYear.text -eq 'Год')-and($textBoxEND2.text -eq 'Месяц')-and($textBoxEND1.text -eq 'День')-and($textBoxEND.text -eq 'Год')-and(!$CB.Checked)){
			$transportservices = (Get-TransportService).name
			foreach ($transportservice in $transportservices){
				$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -MessageSubject $textBoxMTL.Text | select-object $SelectItems
				Add-ListMTLmessagetrackinglog
			}
		}
#6
#почтовый ящик - нет
#тема сообщения - нет
#временная зона - да
#Recipients - нет
		elseif((!$ListMB1.CheckedItems.Count)-and($textBoxMTL.Text -eq 'Введите пожалуйста тему Письма!')-and($textBox2.text -ne 'Месяц')-and($textBox1.text -ne 'День')-and($textBoxYear.text -ne 'Год')-and($textBoxEND2.text -ne 'Месяц')-and($textBoxEND1.text -ne 'День')-and($textBoxEND.text -ne 'Год')-and(!$CB.Checked)){
			$transportservices = (Get-TransportService).name
			foreach ($transportservice in $transportservices){
				$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -start "$month/$day/$year 00:00:00" -end "$mend/$dend/$yend 23:59:00" | select-object $SelectItems
				Add-ListMTLmessagetrackinglog
			}
		}
#7
#почтовый ящик - нет
#тема сообщения - да
#временная зона - да
#Recipients - нет
		elseif((!$ListMB1.CheckedItems.Count)-and($textBoxMTL.Text -ne 'Введите пожалуйста тему Письма!')-and($textBox2.text -ne 'Месяц')-and($textBox1.text -ne 'День')-and($textBoxYear.text -ne 'Год')-and($textBoxEND2.text -ne 'Месяц')-and($textBoxEND1.text -ne 'День')-and($textBoxEND.text -ne 'Год')-and(!$CB.Checked)){
			$transportservices = (Get-TransportService).name
			foreach ($transportservice in $transportservices){
				$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -start "$month/$day/$year 00:00:00" -end "$mend/$dend/$yend 23:59:00" -MessageSubject $textBoxMTL.Text | select-object $SelectItems
				Add-ListMTLmessagetrackinglog
			}
		}
#8
#почтовый ящик - да
#тема сообщения - да
#временная зона - нет
#Recipients - нет
		elseif(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -ne 'Введите пожалуйста тему Письма!')-and($textBox2.text -eq 'Месяц')-and($textBox1.text -eq 'День')-and($textBoxYear.text -eq 'Год')-and($textBoxEND2.text -eq 'Месяц')-and($textBoxEND1.text -eq 'День')-and($textBoxEND.text -eq 'Год')-and(!$CB.Checked)){
			$transportservices = (Get-TransportService).name
			foreach ($transportservice in $transportservices){
				$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Sender $ListMB1.CheckedItems.subitems[2].text -MessageSubject $textBoxMTL.Text | select-object $SelectItems
				Add-ListMTLmessagetrackinglog
			}
		}
#8
#почтовый ящик - да
#тема сообщения - нет
#временная зона - нет
#Recipients - да
		elseif(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -eq 'Введите пожалуйста тему Письма!')-and($textBox2.text -eq 'Месяц')-and($textBox1.text -eq 'День')-and($textBoxYear.text -eq 'Год')-and($textBoxEND2.text -eq 'Месяц')-and($textBoxEND1.text -eq 'День')-and($textBoxEND.text -eq 'Год')-and($CB.Checked)){
			$transportservices = (Get-TransportService).name
			foreach ($transportservice in $transportservices){
				$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Recipients $ListMB1.CheckedItems.subitems[2].text | select-object $SelectItems
				Add-ListMTLmessagetrackinglog
			}
		}
#9
#почтовый ящик - да
#тема сообщения - нет
#временная зона - да
#Recipients - да
		elseif(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -eq 'Введите пожалуйста тему Письма!')-and($textBox2.text -ne 'Месяц')-and($textBox1.text -ne 'День')-and($textBoxYear.text -ne 'Год')-and($textBoxEND2.text -ne 'Месяц')-and($textBoxEND1.text -ne 'День')-and($textBoxEND.text -ne 'Год')-and($CB.Checked)){
			$transportservices = (Get-TransportService).name
			foreach ($transportservice in $transportservices){
				$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Recipients $ListMB1.CheckedItems.subitems[2].text -start "$month/$day/$year 00:00:00" -end "$mend/$dend/$yend 23:59:00" | select-object $SelectItems
				Add-ListMTLmessagetrackinglog
			}
		}
			if(!$listMTL.Items.count){$ListMTL.Items.Add("Сообщение не найдено...")}
			if ($ListMTL.Items.Count -gt 1){SortListTwoviewMTL}
}

		function export-messagetracker-easy {
				$ListMTL.items.clear()
				$month = $textBox2.text
				$day= $textBox1.Text
				$year = $textBoxYear.text
				$mend = $textBoxEND2.text
				$dend = $textBoxEND1.Text
				$yend = $textBoxEND.text
		#1
		#почтовый ящик - да
		#тема сообщения - да
		#временная зона - да
		#Recipients - да
				if(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -ne 'Введите пожалуйста тему Письма!')-and($textBox2.text -ne 'Месяц')-and($textBox1.text -ne 'День')-and($textBoxYear.text -ne 'Год')-and($textBoxEND2.text -ne 'Месяц')-and($textBoxEND1.text -ne 'День')-and($textBoxEND.text -ne 'Год')-and($CB.Checked)){
					$transportservices = (Get-TransportService).name
					foreach ($transportservice in $transportservices){
						$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Recipients $ListMB1.CheckedItems.subitems[2].text -start "$month/$day/$year 00:00:00" -end "$mend/$dend/$yend 23:59:00" -MessageSubject $textBoxMTL.Text | select-object $SelectItemsEasy
						Add-ListMTLmessagetrackinglog-Easy
					}
				}
		#2
		#почтовый ящик - да
		#тема сообщения - да
		#временная зона - да
		#Recipients - нет
				elseif(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -ne 'Введите пожалуйста тему Письма!')-and($textBox2.text -ne 'Месяц')-and($textBox1.text -ne 'День')-and($textBoxYear.text -ne 'Год')-and($textBoxEND2.text -ne 'Месяц')-and($textBoxEND1.text -ne 'День')-and($textBoxEND.text -ne 'Год')-and(!$CB.Checked)){
					$transportservices = (Get-TransportService).name
					foreach ($transportservice in $transportservices){
						$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Sender $ListMB1.CheckedItems.subitems[2].text -start "$month/$day/$year 00:00:00" -end "$mend/$dend/$yend 23:59:00" -MessageSubject $textBoxMTL.Text | select-object $SelectItemsEasy
						Add-ListMTLmessagetrackinglog-Easy
					}
				}
		#3
		#почтовый ящик - да
		#тема сообщения - нет
		#временная зона - да
		#Recipients - нет
				elseif(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -eq 'Введите пожалуйста тему Письма!')-and($textBox2.text -ne 'Месяц')-and($textBox1.text -ne 'День')-and($textBoxYear.text -ne 'Год')-and($textBoxEND2.text -ne 'Месяц')-and($textBoxEND1.text -ne 'День')-and($textBoxEND.text -ne 'Год')-and(!$CB.Checked)){
					$transportservices = (Get-TransportService).name
					foreach ($transportservice in $transportservices){
						$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Sender $ListMB1.CheckedItems.subitems[2].text -start "$month/$day/$year 00:00:00" -end "$mend/$dend/$yend 23:59:00" | select-object $SelectItemsEasy
						Add-ListMTLmessagetrackinglog-Easy
					}
				}
		#4
		#почтовый ящик - да
		#тема сообщения - нет
		#временная зона - нет
		#Recipients - нет
				elseif(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -eq 'Введите пожалуйста тему Письма!')-and($textBox2.text -eq 'Месяц')-and($textBox1.text -eq 'День')-and($textBoxYear.text -eq 'Год')-and($textBoxEND2.text -eq 'Месяц')-and($textBoxEND1.text -eq 'День')-and($textBoxEND.text -eq 'Год')-and(!$CB.Checked)){
					$transportservices = (Get-TransportService).name
					foreach ($transportservice in $transportservices){
						$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Sender $ListMB1.CheckedItems.subitems[2].text | select-object $SelectItemsEasy
						Add-ListMTLmessagetrackinglog-Easy
					}
				}
		#5
		#почтовый ящик - нет
		#тема сообщения - да
		#временная зона - нет
		#Recipients - нет
				elseif((!$ListMB1.CheckedItems.Count)-and($textBoxMTL.Text -ne 'Введите пожалуйста тему Письма!')-and($textBox2.text -eq 'Месяц')-and($textBox1.text -eq 'День')-and($textBoxYear.text -eq 'Год')-and($textBoxEND2.text -eq 'Месяц')-and($textBoxEND1.text -eq 'День')-and($textBoxEND.text -eq 'Год')-and(!$CB.Checked)){
					$transportservices = (Get-TransportService).name
					foreach ($transportservice in $transportservices){
						$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -MessageSubject $textBoxMTL.Text | select-object $SelectItemsEasy
						Add-ListMTLmessagetrackinglog-Easy
					}
				}
		#6
		#почтовый ящик - нет
		#тема сообщения - нет
		#временная зона - да
		#Recipients - нет
				elseif((!$ListMB1.CheckedItems.Count)-and($textBoxMTL.Text -eq 'Введите пожалуйста тему Письма!')-and($textBox2.text -ne 'Месяц')-and($textBox1.text -ne 'День')-and($textBoxYear.text -ne 'Год')-and($textBoxEND2.text -ne 'Месяц')-and($textBoxEND1.text -ne 'День')-and($textBoxEND.text -ne 'Год')-and(!$CB.Checked)){
					$transportservices = (Get-TransportService).name
					foreach ($transportservice in $transportservices){
						$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -start "$month/$day/$year 00:00:00" -end "$mend/$dend/$yend 23:59:00" | select-object $SelectItemsEasy
						Add-ListMTLmessagetrackinglog-Easy
					}
				}
		#7
		#почтовый ящик - нет
		#тема сообщения - да
		#временная зона - да
		#Recipients - нет
				elseif((!$ListMB1.CheckedItems.Count)-and($textBoxMTL.Text -ne 'Введите пожалуйста тему Письма!')-and($textBox2.text -ne 'Месяц')-and($textBox1.text -ne 'День')-and($textBoxYear.text -ne 'Год')-and($textBoxEND2.text -ne 'Месяц')-and($textBoxEND1.text -ne 'День')-and($textBoxEND.text -ne 'Год')-and(!$CB.Checked)){
					$transportservices = (Get-TransportService).name
					foreach ($transportservice in $transportservices){
						$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -start "$month/$day/$year 00:00:00" -end "$mend/$dend/$yend 23:59:00" -MessageSubject $textBoxMTL.Text | select-object $SelectItemsEasy
						Add-ListMTLmessagetrackinglog-Easy
					}
				}
		#8
		#почтовый ящик - да
		#тема сообщения - да
		#временная зона - нет
		#Recipients - нет
				elseif(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -ne 'Введите пожалуйста тему Письма!')-and($textBox2.text -eq 'Месяц')-and($textBox1.text -eq 'День')-and($textBoxYear.text -eq 'Год')-and($textBoxEND2.text -eq 'Месяц')-and($textBoxEND1.text -eq 'День')-and($textBoxEND.text -eq 'Год')-and(!$CB.Checked)){
					$transportservices = (Get-TransportService).name
					foreach ($transportservice in $transportservices){
						$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Sender $ListMB1.CheckedItems.subitems[2].text -MessageSubject $textBoxMTL.Text | select-object $SelectItemsEasy
						Add-ListMTLmessagetrackinglog-Easy
					}
				}
		#8
		#почтовый ящик - да
		#тема сообщения - нет
		#временная зона - нет
		#Recipients - да
				elseif(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -eq 'Введите пожалуйста тему Письма!')-and($textBox2.text -eq 'Месяц')-and($textBox1.text -eq 'День')-and($textBoxYear.text -eq 'Год')-and($textBoxEND2.text -eq 'Месяц')-and($textBoxEND1.text -eq 'День')-and($textBoxEND.text -eq 'Год')-and($CB.Checked)){
					$transportservices = (Get-TransportService).name
					foreach ($transportservice in $transportservices){
						$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Recipients $ListMB1.CheckedItems.subitems[2].text | select-object $SelectItemsEasy
						Add-ListMTLmessagetrackinglog-Easy
					}
				}
		#9
		#почтовый ящик - да
		#тема сообщения - нет
		#временная зона - да
		#Recipients - да
				elseif(($ListMB1.CheckedItems.count)-and($textBoxMTL.Text -eq 'Введите пожалуйста тему Письма!')-and($textBox2.text -ne 'Месяц')-and($textBox1.text -ne 'День')-and($textBoxYear.text -ne 'Год')-and($textBoxEND2.text -ne 'Месяц')-and($textBoxEND1.text -ne 'День')-and($textBoxEND.text -ne 'Год')-and($CB.Checked)){
					$transportservices = (Get-TransportService).name
					foreach ($transportservice in $transportservices){
						$MTss = Get-MessageTrackingLog -Server $transportservice -ResultSize unlimited -Recipients $ListMB1.CheckedItems.subitems[2].text -start "$month/$day/$year 00:00:00" -end "$mend/$dend/$yend 23:59:00" | select-object $SelectItemsEasy
						Add-ListMTLmessagetrackinglog-Easy
					}
				}
					if(!$listMTL.Items.count){$ListMTL.Items.Add("Сообщение не найдено...")}
					if ($ListMTL.Items.Count -gt 1){SortListTwoviewMTL}
		}

$menucopyMTL = New-Object System.Windows.Forms.MenuItem
$menucopyMTL.Text = "Копировать содержимое"
$menucopyMTL.Enabled = $false
$menucopyMTL.Add_Click({
Set-ClipBoardMTL
})
$menucopyMTL01 = New-Object System.Windows.Forms.MenuItem
$menucopyMTL01.Text = "Проверить наличие сообщения в почтовом ящике - CheckBox"
$menucopyMTL01.Enabled = $false
$menucopyMTL01.Add_Click({
Check-mail
})
$menucopyMTL02 = New-Object System.Windows.Forms.MenuItem
$menucopyMTL02.Text = "Проверить наличие сообщения в почтовых ящиках - Sender, Recipients"
$menucopyMTL02.Enabled = $false
$menucopyMTL02.Add_Click({
Check-mailAll
})
$menucopyMTL1 = New-Object System.Windows.Forms.MenuItem
$menucopyMTL1.Text = "Удалить сообщение в почтовом ящике - CheckBox!"
$menucopyMTL1.Enabled = $false
$menucopyMTL1.Add_Click({
Delete-mail
})
$menucopyMTL2 = New-Object System.Windows.Forms.MenuItem
$menucopyMTL2.Text = "Удалить сообщение в почтовых ящиках - Sender, Recipients!"
$menucopyMTL2.Enabled = $false
$menucopyMTL2.Add_Click({
Delete-mailAll
})
$Context = New-Object System.Windows.Forms.ContextMenu
$Context.MenuItems.AddRange(@($menucopyMTL))
$Context.MenuItems.AddRange(@($menucopyMTL01))
$Context.MenuItems.AddRange(@($menucopyMTL02))
$Context.MenuItems.AddRange(@($menucopyMTL1))
$Context.MenuItems.AddRange(@($menucopyMTL2))
$ListMTL.contextmenu = $Context

			function Delete-mail {
				$OUTPUT = [System.Windows.Forms.MessageBox]::Show("Вы уверены?","Внимание",4)
				if ($OUTPUT -eq "YES")
				{
						#создание массива из таблицы
							$CsvArray = @()
						    $listMTL.Items | %{
						        $Object = New-Object PSObject
						        $lvItem = $_
						        $listMTL.Columns | %{
						            $Object | add-member Noteproperty -Name $_.Text -Value  $lvItem.SubItems[$_.index].text -Force
						        }
								$CsvArray += $Object
						        Remove-Variable Object
						    }
						#	Write-Host $CsvArray | ft
						#    Return ,$CsvArray
							$subject=$listMTL.FocusedItem.subitems[6].text
#							$messageid=$ListMTL.FocusedItem.subitems[2].text
								$SM = Search-Mailbox -Identity $ListMB1.CheckedItems.subitems[2].text -SearchQuery Subject:"$subject" -TargetMailbox "AdmExc@nornik.ru" -TargetFolder "MessageTrackingLog" -LogOnly | select ResultItemsCount
								if($SM.ResultItemsCount -eq "0"){
#									$i=0
#									foreach($Obj in $CsvArray){
#										if($subject -eq $Obj.MessageSubject){
										$ListMTL.FocusedItem.backcolor = "green"
										$ListMTL.FocusedItem.forecolor = "black"
#										}
#										$i+=1
#									}
								}else{
#									$i=0
								Search-Mailbox -Identity $ListMB1.CheckedItems.subitems[2].text -SearchQuery Subject:"$subject" -TargetMailbox "AdmExc@nornik.ru" -TargetFolder "MessageTrackingLog" -LogLevel full -DeleteContent -confirm:$false -Force
#							foreach($Obj in $CsvArray){
#										if($subject -eq $Obj.MessageSubject){
                                  [system.Windows.Forms.MessageBox]::Show("Письмо удалено!")
										$ListMTL.FocusedItem.backcolor = "red"
										$ListMTL.FocusedItem.forecolor = "black"
#										}
#										$i+=1
#									}
								}
				}
			}
			#-------------------------------------------------------------------------------------------------------------------------------------
			
					function Delete-mailAll {
						$OUTPUT = [System.Windows.Forms.MessageBox]::Show("Вы уверены?","Внимание",4)
						if ($OUTPUT -eq "YES")
						{
							#создание массива из таблицы
								$CsvArray = @()
							    $listMTL.Items | %{
							        $Object = New-Object PSObject
							        $lvItem = $_
							        $listMTL.Columns | %{
							            $Object | add-member Noteproperty -Name $_.Text -Value  $lvItem.SubItems[$_.index].text -Force
							        }
									$CsvArray += $Object
							        Remove-Variable Object
							    }
							#	Write-Host $CsvArray | ft
							#    Return ,$CsvArray
								$successful = @()
								$recipients = @()
								$subject = $listMTL.FocusedItem.subitems[6].text
#								$messageid = $ListMTL.FocusedItem.subitems[2].text
								$sender = $ListMTL.FocusedItem.subitems[4].text
								$recipients = $ListMTL.FocusedItem.subitems[5].text
								[string[]]$recipients = $recipients.Split(' ',[System.StringSplitOptions]::RemoveEmptyEntries) #создает из строкового значения массив
								$recipients+=$sender
								foreach ($recipient in $recipients){
#								[system.Windows.Forms.MessageBox]::Show("$recipient" , "1")
										$SM = Search-Mailbox -Identity $recipient -SearchQuery Subject:"$subject" -TargetMailbox "AdmExc@nornik.ru" -TargetFolder "MessageTrackingLog" -LogOnly | select ResultItemsCount
									if($SM.ResultItemsCount -eq "0"){
									}else{
										Search-Mailbox -Identity $recipient -SearchQuery Subject:"$subject" -TargetMailbox "AdmExc@nornik.ru" -TargetFolder "MessageTrackingLog" -LogLevel full -DeleteContent -confirm:$false -Force
										$successful = "red"
									}
#									$i=0
#									foreach($Obj in $CsvArray){
#										if($subject -eq $Obj.MessageSubject){
											if($successful){
												[system.Windows.Forms.MessageBox]::Show("$recipient" , "Письмо удалено!")
												$ListMTL.FocusedItem.backcolor = "red"
												$ListMTL.FocusedItem.forecolor = "black"
											}else{
												$ListMTL.FocusedItem.backcolor = "green"
												$ListMTL.FocusedItem.forecolor = "black"
											}
#										}
#									$i+=1
#									}
								}
									
						}
					}
					#-------------------------------------------------------------------------------------------------------------------------------------
					
							function Check-mail {
									#создание массива из таблицы
										$CsvArray = @()
									    $listMTL.Items | %{
									        $Object = New-Object PSObject
									        $lvItem = $_
									        $listMTL.Columns | %{
									            $Object | add-member Noteproperty -Name $_.Text -Value  $lvItem.SubItems[$_.index].text -Force
									        }
											$CsvArray += $Object
									        Remove-Variable Object
									    }
									#	Write-Host $CsvArray | ft
									#    Return ,$CsvArray
										$subject=$listMTL.FocusedItem.subitems[6].text
										$messageid=$ListMTL.FocusedItem.subitems[2].text
											$SM = Search-Mailbox -Identity $ListMB1.CheckedItems.subitems[4].text -SearchQuery Subject:"$subject" -TargetMailbox "AdmExc@nornik.ru" -TargetFolder "MessageTrackingLog" -LogOnly | select ResultItemsCount
											if($SM.ResultItemsCount -eq "0"){
#												$i=0
#												foreach($Obj in $CsvArray){
#													if($subject -eq $Obj.MessageSubject){
													$ListMTL.FocusedItem.backcolor = "black"
													$ListMTL.FocusedItem.forecolor = "white"
#													}
#													$i+=1
#												}
											}else{
#												$i=0
#												foreach($Obj in $CsvArray){
#													if($subject -eq $Obj.MessageSubject){
													$ListMTL.FocusedItem.backcolor = "yellow"
													$ListMTL.FocusedItem.forecolor = "black"
#													}
#													$i+=1
#												}
											}
							}
						#-------------------------------------------------------------------------------------------------------------------------------------
								function Check-mailAll {
										#создание массива из таблицы
											$CsvArray = @()
										    $listMTL.Items | %{
										        $Object = New-Object PSObject
										        $lvItem = $_
										        $listMTL.Columns | %{
										            $Object | add-member Noteproperty -Name $_.Text -Value  $lvItem.SubItems[$_.index].text -Force
										        }
												$CsvArray += $Object
										        Remove-Variable Object
										    }
										#	Write-Host $CsvArray | ft
										#    Return ,$CsvArray
#											$successful = @()
											$recipients = @()
											$subject = $listMTL.FocusedItem.subitems[6].text
											$messageid = $ListMTL.FocusedItem.subitems[2].text
											$sender = $ListMTL.FocusedItem.subitems[4].text
											$recipients = $ListMTL.FocusedItem.subitems[5].text
											[string[]]$recipients = $recipients.Split(' ',[System.StringSplitOptions]::RemoveEmptyEntries) #создает из строкового значения массив
											$recipients+=$sender
											foreach ($recipient in $recipients){
													$SM = Search-Mailbox -Identity $recipient -SearchQuery Subject:"$subject" -TargetMailbox "AdmExc@nornik.ru" -TargetFolder "MessageTrackingLog" -LogOnly | select ResultItemsCount
#													[system.Windows.Forms.MessageBox]::Show("$SM.ResultItemsCount" , "1")
												if($SM.ResultItemsCount -eq "0"){
													[system.Windows.Forms.MessageBox]::Show("$recipient" , "Письма НЕТ!")
#													$i=0
#													foreach($Obj in $CsvArray){
#														if($subject -eq $Obj.MessageSubject){
															$ListMTL.FocusedItem.backcolor = "black"
															$ListMTL.FocusedItem.forecolor = "white"
#														}
#														$i+=1
#													}
												}else{
													[system.Windows.Forms.MessageBox]::Show("$recipient" , "Есть письмо!")
#													$i=0
#													foreach($Obj in $CsvArray){
#														if($subject -eq $Obj.MessageSubject){
															$ListMTL.FocusedItem.backcolor = "yellow"
															$ListMTL.FocusedItem.forecolor = "black"
#														}
#														$i+=1
#													}
												}
											}
											
								}
								#-------------------------------------------------------------------------------------------------------------------------------------


					#Сортироватm в два вида (возрастания и убывание)-------------------------------------------------------------------------------
					function SortListTwoviewMTL {
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
					foreach($ListItem in $ListMTL.Items)
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
					    { return [Double]$_[0] }
					    else
					    { return [String]$_[0] }
					}
					#вся информация собрана; выполнения сортировки (all information is gathered; perform the sort)
					$ListItems = $ListItems | Sort-Object -Property @{Expression=$EvalExpression; Ascending=$Script:LastColumnAscendingTwo}
					#список отсортирован, вывести в list (the list is sorted; display it in the listview)
					$ListMTL.BeginUpdate()
					$ListMTL.Items.Clear()
					foreach($ListItem in $ListItems)
					{
					    $ListMTL.Items.Add($ListItem[1])
					}
					$ListMTL.EndUpdate()
					}
					#------------------------------------------
					
#--------------------------------------------------------

#окно вывода Пользователей
$ListMB1 = New-Object System.Windows.Forms.ListView
$ListMB1.dock = "Top"
$ListMB1.Height = 200
$ListMB1.View = "Details"
$ListMB1.MultiSelect = $false
$ListMB1.FullRowSelect = $True
$ListMB1.AutoSize = $true
$ListMB1.LabelEdit = $True
$ListMB1.AllowColumnReorder = $True
$ListMB1.CheckBoxes = $true
$ListMB1.GridLines = $true
$ListMB1.Columns.Add("DisplayName", 160)
$ListMB1.Columns.Add("Alias", 90)
$ListMB1.Columns.Add("PrimarySmtpAddress", 150)
$ListMB1.Columns.Add("ItemCount", 100)
$ListMB1.Columns.Add("MailboxSize (bytes)", 150)
$ListMB1.Columns.Add("ProhibitSendQuota (bytes)", 150)
$ListMB1.Columns.Add("TargetDataBase", 150)
$ListMB1.add_ItemSelectionChanged({
	if($ListMB1.SelectedItems.Count -cge 1){
		$menucopyMB1.Enabled = $true
	}else{
		$menucopyMB1.Enabled = $false
	}
})
$ListMB1.add_ItemChecked({
	if($listMB1.CheckedItems.Count -eq 1){
		$CB.Enabled = $true
	}else{
		$CB.Enabled = $false
	}
	if($listMB1.CheckedItems.Count -cgt 1){
		$ListMB1.CheckedItems[0].checked = $false
	}
})
$ListMB1.add_ColumnClick({
	if ($ListMB1.Items.Count -gt 1){
		SortListTwoviewMB1 $_.Column
	}
})
$ListMB1.add_KeyDown({
	param($sender, $e)
	if ($_.KeyCode -eq "C" -and $e.Control){
		Set-ClipBoardMB1
	}
	if ($_.keycode -eq "A" -and $e.Control){
		foreach ($ListItem in $ListMB1.Items){
		    $ListItem.selected = $true
		}
	}
})
$tabControlP2.Controls.Add($ListMB1);
#---------------------------------------------------------------

#строка запроса
$SearchOnType = $false
$textBoxMTLMB = New-Object System.Windows.Forms.TextBox;    
$textBoxMTLMB.dock = "top"
$textBoxMTLMB.Location = New-Object System.Drawing.Point(10, 10);    
$textBoxMTLMB.Size = New-Object System.Drawing.Size(365, 10);    
$textBoxMTLMB.Name = "textBox0";      
#$textBoxMTLMB.Text = "user";
$textBoxMTLMBname = 'Введите пожалуйста имя Пользователя!'
$textBoxMTLMB.ForeColor = 'LightGray'
$textBoxMTLMB.Text = $textBoxMTLMBname
$textBoxMTLMB_AddGM = 0;
$textBoxMTLMB.add_Click({
	if($textBoxMTLMB.Text -eq $textBoxMTLMBname)
    {
        #Clear the text
        $textBoxMTLMB.Text = ""
        $textBoxMTLMB.ForeColor = 'WindowText'
    }
	if($textBoxMTLMB.Text -eq $textBoxMTLMB.Tag)
    {
        #Clear the text
        $textBoxMTLMB.Text = ""
        $textBoxMTLMB.ForeColor = 'WindowText'
    }
	})
$textBoxMTLMB.add_KeyPress({
if($textBoxMTLMB.Visible -and $textBoxMTLMB.Tag -eq $null)
    {
        #Initialize the watermark and save it in the Tag property
        $textBoxMTLMB.Tag = $textBoxMTLMB.Text;
        $textBoxMTLMB.ForeColor = 'LightGray'
        #If we have focus then clear out the text
        if($textBoxMTLMB.Focused)
        {
            $textBoxMTLMB.Text = ""
            $textBoxMTLMB.ForeColor = 'WindowText'
        }
    }
})
$textBoxMTLMB.add_Leave({
if($textBoxMTLMB.Text -eq "")
    {
        #Display the watermark
        $textBoxMTLMB.Text = $textBoxMTLMBname
        $textBoxMTLMB.ForeColor = 'LightGray'
    }
	if($textBoxMTLMB.Text -eq "")
    {
        #Display the watermark
        $textBoxMTLMB.Text = $textBoxMTLMB.Tag
        $textBoxMTLMB.ForeColor = 'LightGray'
    }
		})
$tabControlP2.Controls.Add($textBoxMTLMB);
#------------------------------------------------------------------

					#создание запроса textBoxMTLMB--------------------------------------
					function Fill-ListMTLMB ($Mask = "*") {
						$ListMB1.Items.Clear()
						$s = $textBoxMTLMB.Text + "*"
						$mail =  [scriptblock]::create("alias -like `"$s`"")
						$mailn =  [scriptblock]::create("displayname -like `"$s`"")
					    $str = Get-mailbox -Filter $mail | select PrimarySmtpAddress,Alias,DisplayName,ProhibitSendQuota,Database
							if(!$str){
								$str = Get-mailbox -Filter $mailn | select PrimarySmtpAddress,Alias,DisplayName,ProhibitSendQuota,Database}
							    if(!$str){
									$I = $ListMB1.Items.Add("Пользователь не найден...") | Out-Null
								}else{
									foreach ( $item in $str ){
										$str1 = Get-MailboxStatistics $item.PrimarySmtpAddress | select itemcount,totalitemsize
										$size=$str1.totalitemsize -replace ".*[(]",""
										$size1="$size" -replace "[by].*",""
										$ProhibitSendQuota=$item.ProhibitSendQuota -replace ".*[(]",""
										$ProhibitSendQuota1=$ProhibitSendQuota -replace "[by].*",""
										$ser=$moves.StatusDetail
										$moveSDB = $move.SourceDatabase
								    	#Добавляем элемент в список
								    	if($item.DisplayName){$I = $ListMB1.Items.Add($item.DisplayName)}else{$I = $ListMB1.Items.Add("null")}
										if($item.Alias){$I.SubItems.Add($item.Alias)}else{$I.SubItems.Add("null")}
										if($item.PrimarySmtpAddress){$I.SubItems.Add($item.PrimarySmtpAddress)}else{$I.SubItems.Add("null")}
										if($str1.itemcount){$I.SubItems.Add($str1.itemcount)}else{$I.SubItems.Add("0")}
										if("$size1"){$I.SubItems.Add("$size1")}else{$I.SubItems.Add("0")}
										if("$ProhibitSendQuota1"){$I.SubItems.Add("$ProhibitSendQuota1")}else{$I.SubItems.Add("null")}
										if($item.Database){$I.SubItems.Add($item.Database)}else{$I.SubItems.Add("null")}
								    }
								}
					}
					if ($SearchOnType)
					{
					    #Добавляем обработчик на событие TextChanged, который выполняет функцию Fill-List
					    $textBoxMTLMB.add_TextChanged({Fill-ListMTLMB ("*" + $textBoxMTLMB.Text + "*")})
					}
					else #Ищем только при нажатии Enter
					{
					    #Скриптблок (кусок исполняемого кода) который будет выполнен при нажатии клавиши в поле поиска
					    $SB_KeyPress = {
					        #Если была нажата клавиша Enter (13) то...
					        if (13 -eq $_.keychar)
					        {
					            #Вызываем функцию Fill-List
					            Fill-ListMTLMB ("*" + $textBoxMTLMB.Text + "*")
					        }
					    }
					    #Добавляем обработчик на событие KeyPress, указав в качестве выполняемого кода $SB_KeyPress
					    $textBoxMTLMB.add_KeyPress($SB_KeyPress)
					}
					#-----------------------------------------------------------

#контекстное меню вызывается в ListMB1
$menucopyMB1 = New-Object System.Windows.Forms.MenuItem
$menucopyMB1.Text = "Копировать содержимое"
$menucopyMB1.Enabled = $false
$menucopyMB1.Add_Click({
Set-ClipBoardMB1
})

$ContextMenu1 = New-Object System.Windows.Forms.ContextMenu
$ContextMenu1.MenuItems.AddRange(@($menucopyMB1))
$listMB1.ContextMenu = $ContextMenu1
#------------------------------------------------------------------------------------------------------------------------------

#копирует содержимое listMB1--------------------------------------------
Function Set-ClipBoardMB1 {
$CopyTexts = @()
$n = "`n"
$listMB1.SelectedItems | % {$CopyTexts+=$_.subitems.text+$n}
ForEach ($CopyText in $CopyTexts){
$CopyText1 += ";$CopyText"
}
[System.Windows.Forms.Clipboard]::SetText($CopyText1)
}
#------------------------------------------------------------------------
#копирует содержимое listMTL--------------------------------------------
Function Set-ClipBoardMTL {
$CopyTexts = @()
$n = "`n"
$listMTL.SelectedItems | % {$CopyTexts+=$_.subitems.text+$n}
ForEach ($CopyText in $CopyTexts){
$CopyText1 += ";$CopyText"
}
[System.Windows.Forms.Clipboard]::SetText($CopyText1)
}
#------------------------------------------------------------------------

					#Сортироватm в два вида (возрастания и убывание)-------------------------------------------------------------------------------
					function SortListTwoviewMB1 {
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
					foreach($ListItem in $ListMB1.Items)
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
					$ListMB1.BeginUpdate()
					$ListMB1.Items.Clear()
					foreach($ListItem in $ListItems)
					{
					    $ListMB1.Items.Add($ListItem[1])
					}
					$ListMB1.EndUpdate()
					}
					#-------------------------------------------------------------------------------------------------------------------------------