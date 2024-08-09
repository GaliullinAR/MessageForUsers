Function setLogin ($Path) {
    $Users = Get-ADUser -Filter *
    $ExcelObj = New-Object -comobject Excel.Application
    $ExcelWorkBook = $ExcelObj.Workbooks.Open($Path)
    $ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Лист1")
    $MaxRows = ($ExcelWorkSheet.UsedRange.Rows).count

    for ($count = 2; $count -ile $MaxRows; $count++) {
        $user = $ExcelWorkSheet.Rows.Item($count).Columns.Item(1).Text
        $currentUser = $Users | where {$_.Name -eq $user}
        if ($currentUser) {
            if ($currentUser.SamAccountName -is [string]) {
                $ExcelWorkSheet.Rows.Item($count).Columns.Item(2).Value2 = $currentUser.SamAccountName
            }
        } else {
            continue
        }
    }

    $ExcelWorkBook.Save()
    $ExcelWorkBook.close($true)

    $ExcelObj.Quit()
}


Function CreateMsgForUsers($Path) {
    $ExcelObj = New-Object -comobject Excel.Application # Создание объекта Excel
    $ExcelWorkBook = $ExcelObj.Workbooks.Open($Path) # Получение книги
    
    $ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Лист1") # Выбор листа
    
    $MaxRows = ($ExcelWorkSheet.UsedRange.Rows).count; # Получение количества строк листа Excel
    
    $Computers = Get-ADComputer -Filter * -SearchBase "OU=Компьютеры,OU=ТАИФ КОС,OU=SIBUR Holding,DC=sibur,DC=local" -Properties *
    
    $LoggedUsersFromComputers = @()
    $StatusMessage = @()
    
    
    foreach ($computer in $Computers) {
    
      if ($computer.Description -eq $NULL) {
        continue
      }
      if ($computer.Description.Split(" ")[0] -eq "Logged") {
        $ComputerName = $computer.Name
        $UserName = $computer.Description.Split(" ")[2]
        $LoggedDate = $computer.Description.Split(" ")[6]
        $LoggedUsersFromComputers += @{ComputerName=$ComputerName;UserName=$UserName;LoggedDate=$LoggedDate}
      } else {
        $ComputerName = $computer.Name
        $UserName = $computer.Description.Split(" ")[3]
        $LoggedDate = $computer.Description.Split(" ")[7]
        $LoggedUsersFromComputers += @{ComputerName=$ComputerName;UserName=$UserName;LoggedDate=$LoggedDate}
      }
    }
    
    for ($count = 2; $count -ile $MaxRows; $count++) {
        $user = $ExcelWorkSheet.Rows.Item($count).Columns.Item(2).Text
        $fio = $ExcelWorkSheet.Rows.Item($count).Columns.Item(1).Text
        $inventoryNumber = $ExcelWorkSheet.Rows.Item($count).Columns.Item(3).Text
        $inventoryDescription = $ExcelWorkSheet.Rows.Item($count).Columns.Item(4).Text
        $ticketStatus = $ExcelWorkSheet.Rows.Item($count).Columns.Item(5).Text
        $endDate = $ExcelWorkSheet.Rows.Item($count).Columns.Item(6).Text
        $isGivedInventory = $ExcelWorkSheet.Rows.Item($count).Columns.Item(7).Text
        
        $currentDate = Get-Date
    
        if ($ticketStatus -eq 'да') {
            continue
        }

        if ($isGivedInventory -eq "да") {
            continue
        }

        if ($user -eq '') {
            $StatusMessage += @{Login=$NULL;UserName=$fio;StatusMessage="Найдено несколько пользователей по ФИО. Не отправлено";MultiLogin=$True}
            continue
        }

    
        $currentDateUsers = $LoggedUsersFromComputers | where {$_.LoggedDate -eq $currentDate.ToString("dd.MM.yyyy")}
        $currentUsers = $currentDateUsers | where {$_.UserName -eq $user}
        $currentUser = $currentUsers.UserName

        if ($currentUsers) {
            $Sender = "Группа поддержки рабочих мест Сибур Коннект"
            $Message = "Уважаемый(ая) $fio, на вас числится оборудование c инвентарными номерами $inventoryNumber, в связи с увольнением просим вас подать соответствующее обращение и сдать технику до $endDate"
            $RemoteComputer = $currentUsers.ComputerName
            
            # Invoke-Command -ComputerName $RemoteComputer -ScriptBlock ${function:New-ToastNotification} -ArgumentList $Sender,$Message
        
            try {
                Invoke-Command -ComputerName $RemoteComputer -ScriptBlock {
                    Param($Message)
                    msg * $Message
                    
                } -ArgumentList $Message -ErrorAction Stop 
                $StatusMessage += @{Login=$currentUsers.UserName;UserName=$fio;$ComputerName=$currentUsers.ComputerName;StatusMessage="Отправлено"}
                
            } catch {
                $StatusMessage += @{Login=$currentUsers.UserName;UserName=$fio;$ComputerName=$currentUsers.ComputerName;StatusMessage="Не отправлено. Компьютер выключен."}
                
            }

        } else {
            
            
            $StatusMessage += @{Login=$NULL;UserName=$fio;$ComputerName=$currentUsers.ComputerName;StatusMessage="Авторизованного пользователя не найдено. Сообщение не отправлено"}
            
        }
    }
    return $StatusMessage
}












# Function CreateMsgForUsers($Excel, $ComputersFromAD) {
#     $MaxRows = ($Excel.UsedRange.Rows).count;

#     $LoggedUsersFromComputers = @()
#     $StatusMessage = @()
    
#     foreach ($computer in $ComputersFromAD) {
    
#       if ($computer.Description -eq $NULL) {
#         continue
#       }
#       if ($computer.Description.Split(" ")[0] -eq "Logged") {
#         $ComputerName = $computer.Name
#         $UserName = $computer.Description.Split(" ")[2]
#         $LoggedDate = $computer.Description.Split(" ")[6]
#         $LoggedUsersFromComputers += @{ComputerName=$ComputerName;UserName=$UserName;LoggedDate=$LoggedDate}
#       } else {
#         $ComputerName = $computer.Name
#         $UserName = $computer.Description.Split(" ")[3]
#         $LoggedDate = $computer.Description.Split(" ")[7]
#         $LoggedUsersFromComputers += @{ComputerName=$ComputerName;UserName=$UserName;LoggedDate=$LoggedDate}
#       }
#     }
    
#     for ($count = 2; $count -ile $MaxRows; $count++) {
#         $user = $Excel.Rows.Item($count).Columns.Item(2).Text
#         $fio = $Excel.Rows.Item($count).Columns.Item(1).Text
#         $inventoryNumber = $Excel.Rows.Item($count).Columns.Item(3).Text
#         $inventoryDescription = $Excel.Rows.Item($count).Columns.Item(4).Text
#         $ticketStatus = $Excel.Rows.Item($count).Columns.Item(5).Text
#         $endDate = $Excel.Rows.Item($count).Columns.Item(6).Text
#         $isGivedInventory = $Excel.Rows.Item($count).Columns.Item(7).Text
        
#         $currentDate = Get-Date
    
#         if ($ticketStatus -eq 'да') {
#             continue
#         }

#         if ($isGivedInventory -eq "да") {
#             continue
#         }

#         if ($user -eq '') {
#             $StatusMessage += @{Login=$NULL;UserName=$fio;StatusMessage="Найдено несколько пользователей по ФИО. Не отправлено";MultiLogin=$True}
#             continue
#         }

    
#         $currentDateUsers = $LoggedUsersFromComputers | where {$_.LoggedDate -eq $currentDate.ToString("dd.MM.yyyy")}
#         $currentUsers = $currentDateUsers | where {$_.UserName -eq $user}
#         $currentUser = $currentUsers.UserName

#         if ($currentUsers) {
#             $Sender = "Группа поддержки рабочих мест Сибур Коннект"
#             $Message = "Уважаемый(ая) $fio, на вас числится оборудование c инвентарными номерами $inventoryNumber, в связи с увольнением просим вас подать соответствующее обращение и сдать технику до $endDate"
#             $RemoteComputer = $currentUsers.ComputerName
            
#             # Invoke-Command -ComputerName $RemoteComputer -ScriptBlock ${function:New-ToastNotification} -ArgumentList $Sender,$Message
        
#             try {
#                 Invoke-Command -ComputerName $RemoteComputer -ScriptBlock {
#                     Param($Message)
#                     msg * $Message
                    
#                 } -ArgumentList $Message -ErrorAction Stop 
#                 $StatusMessage += @{Login=$currentUsers.UserName;UserName=$fio;$ComputerName=$currentUsers.ComputerName;StatusMessage="Отправлено"}
                
#             } catch {
#                 $StatusMessage += @{Login=$currentUsers.UserName;UserName=$fio;$ComputerName=$currentUsers.ComputerName;StatusMessage="Не отправлено. Компьютер выключен."}
                
#             }

#         } else {
            
            
#             $StatusMessage += @{Login=$NULL;UserName=$fio;$ComputerName=$currentUsers.ComputerName;StatusMessage="Авторизованного пользователя не найдено. Сообщение не отправлено"}
            
#         }
#     }
#     return $StatusMessage
# }











Add-Type -assembly System.Windows.Forms
$CenterScreen = [System.Windows.Forms.FormStartPosition]::CenterScreen;
$window_form = New-Object System.Windows.Forms.Form
$window_form.StartPosition = $CenterScreen
$window_form.Text ='Message for User'
$window_form.Width = 640
$window_form.Height = 610
$window_form.minimumSize = New-Object System.Drawing.Size(640,610) 
$window_form.maximumSize = New-Object System.Drawing.Size(640,610) 
$window_form.BackColor = '#f5f5f5'
$window_form.AutoSize = $true

$form_status_label1 = New-Object System.Windows.Forms.Label
$form_status_label1 = New-Object System.Windows.Forms.Label
$form_status_label1.Text = "Интервал (секунды):"
$form_status_label1.Location = New-Object System.Drawing.Point(10,10)
$form_status_label1.AutoSize = $true
$form_status_label1.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Regular)
$window_form.Controls.Add($form_status_label1)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(12,40)
$textBox.Size = New-Object System.Drawing.Size(100,30)
$textBox.Height = 40
$window_form.Controls.Add($textBox)

$form_status_label2 = New-Object System.Windows.Forms.Label
$form_status_label2 = New-Object System.Windows.Forms.Label
$form_status_label2.Text = "Количество повторов:"
$form_status_label2.Location = New-Object System.Drawing.Point(10,80)
$form_status_label2.AutoSize = $true
$form_status_label2.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Regular)
$window_form.Controls.Add($form_status_label2)

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(12,110)
$textBox1.Size = New-Object System.Drawing.Size(100,30)
$window_form.Controls.Add($textBox1)

$SendMessage = New-Object System.Windows.Forms.Button
$SendMessage.BackColor = "#3ecede"
$SendMessage.text = "Отправить"
$SendMessage.Location = New-Object System.Drawing.Point(10,140)
$SendMessage.Size = New-Object System.Drawing.Size(180,30)
$SendMessage.AutoSize = $True
$SendMessage.Add_Click({
    
    
    $interation = 0
    $interval = 0;

    try {
        $interation = [int]$textBox1.text
        $interval = [int]$textBox.text
    }
    catch {
        msg * "Не корректный тип данных интервала или количества повторов, пожалуйста введите число"
    }


    $currentPath = ''
    if ($textBox2.text) {
        $currentPath = $textBox2.text
        setLogin -Path $currentPath
        
        $StatusLabel1.Text = "Процесс отправки сообщения..."
        
        
        for ($count = 1; $count -ile $interation; $count++){
            
            $status = CreateMsgForUsers -Path $currentPath
            
            if ($count -eq 1) {
                foreach ($item in $status) {
                                       
                    if ($item.Login -ne $NULL) {
                        $userName = $item.UserName
                        $RequestMessage = $item.StatusMessage
                        $Login = $item.Login
                        $StatusList.Items.Add("$UserName     |      $Login      |      $RequestMessage")
                    } elseif ($item.MultiLogin) {
                        $userName = $item.UserName
                        $RequestMessage = $item.StatusMessage
                        $StatusList.Items.Add("$UserName     |      Полная теска      |      $RequestMessage")
                    } else {
                        $userName = $item.UserName
                        $RequestMessage = $item.StatusMessage
                        $StatusList.Items.Add("$UserName     |      Не найдена авторизация      |      $RequestMessage")
                    }
                }
            }

            $form_status_label4.text = "{0:N1}" -f ((100 / $interation) * $count)
            if ($count -ne $interation) {
                Start-Sleep -Seconds $interval
            }
            if ($count -eq $interation) {
                $StatusLabel1.Text = "Выполнено"
            }
        }
    } else {
        msg * "Файл не найден. Пожалуста введите корректный путь"
        $StatusLabel1.Text = "Ошибка получения файла"
        return $False
    }
    
})
$window_form.Controls.Add($SendMessage)

$form_status_label3 = New-Object System.Windows.Forms.Label
$form_status_label3.Text = "Выполнено %:"
$form_status_label3.Location = New-Object System.Drawing.Point(230,10)
$form_status_label3.AutoSize = $true
$form_status_label3.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Regular)
$window_form.Controls.Add($form_status_label3)

$form_status_label4 = New-Object System.Windows.Forms.Label

$form_status_label4.Location = New-Object System.Drawing.Point(400,10)
$form_status_label4.AutoSize = $true
$form_status_label4.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Regular)
$window_form.Controls.Add($form_status_label4)

$Path = New-Object System.Windows.Forms.Label
$Path.Text = "Путь к Excel файлу:"
$Path.Location = New-Object System.Drawing.Point(230,80)
$Path.AutoSize = $true
$Path.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Regular)
$window_form.Controls.Add($Path)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(232,110)
$textBox2.Size = New-Object System.Drawing.Size(250,30)
$textBox2.Height = 40
$window_form.Controls.Add($textBox2)

$StopMessage = New-Object System.Windows.Forms.Button
$StopMessage.BackColor = "#3ecede"
$StopMessage.text = "Стоп"
$StopMessage.Location = New-Object System.Drawing.Point(10,140)
$StopMessage.Size = New-Object System.Drawing.Size(180,30)
$StopMessage.AutoSize = $True
$window_form.Controls.Add($SendMessage)

$btnFileBrowser = New-Object System.Windows.Forms.Button
$btnFileBrowser.BackColor = '#1a80b6'
$btnFileBrowser.text = 'Обзор'
$btnFileBrowser.Location = New-Object System.Drawing.Point(232,140)
$btnFileBrowser.Size = New-Object System.Drawing.Size(70,30)
$btnFileBrowser.Add_Click({
    Add-Type -AssemblyName System.windows.forms | Out-Null
    $OpenDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
    $OpenDialog.initialDirectory = $initialDirectory
    $OpenDialog.ShowDialog() | Out-Null
    $filePath = $OpenDialog.filename
    $textBox2.text = $filePath  
})

$window_form.Controls.Add($btnFileBrowser)

$StatusList = New-Object System.Windows.Forms.ListBox
$StatusList.Location = New-Object System.Drawing.Point(10,200)
$StatusList.Size = New-Object System.Drawing.Size(575,320)

$window_form.Controls.Add($StatusList)

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Cтатус:"
$StatusLabel.Location = New-Object System.Drawing.Point(10,530)
$StatusLabel.AutoSize = $true
$StatusLabel.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Regular)
$window_form.Controls.Add($StatusLabel)

$StatusLabel1 = New-Object System.Windows.Forms.Label
$StatusLabel1.Text = ""
$StatusLabel1.Location = New-Object System.Drawing.Point(110,530)
$StatusLabel1.AutoSize = $true
$StatusLabel1.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Regular)
$window_form.Controls.Add($StatusLabel1)

$window_form.ShowDialog()