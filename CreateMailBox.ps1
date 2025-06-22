#v1.2 (02.07.2025)
#Developed by Danilovich M.D.


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn




# Проверка запуска с правами администратора
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $currentScript = $MyInvocation.MyCommand.Definition
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$currentScript`"" -Verb RunAs
    exit
}



# Подключение к Exchange с использованием учетных данных текущего пользователя
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://YourExchangeServer-XX/PowerShell/ -Authentication Kerberos
    Import-PSSession $Session -DisableNameChecking -AllowClobber
    Write-Host "Подключение к Exchange успешно установлено." -ForegroundColor Green
} catch {
    Write-Host "Ошибка подключения к Exchange: $($_.Exception.Message)" -ForegroundColor Red
    Start-sleep 10
    exit
}








# Функция для получения пользователей из OU "Main" в Active Directory без почтового ящика
function Get-ADUsersWithoutMailbox {
    try {
        $ou = "OU=ExampleOU,DC=example,DC=com"  #Указать путь к вашей OU
        $users = Get-ADUser -Filter * -SearchBase $ou -Properties DisplayName, Mail, SamAccountName
        $usersWithoutMailbox = $users | Where-Object { -not $_.Mail } | Select-Object DisplayName, SamAccountName | Sort-Object DisplayName
        return $usersWithoutMailbox
    } catch {
        Write-Host "Ошибка при получении пользователей из Active Directory: $_"
        return @()
    }
}



# Функция для обновления содержимого ListBox
function Update-ListBox {
    $listBoxUsers.Items.Clear()
    $usersWithoutMailbox = Get-ADUsersWithoutMailbox
    if ($usersWithoutMailbox.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Нет пользователей без почтового ящика.", "Информация", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        foreach ($user in $usersWithoutMailbox) {
            $listBoxUsers.Items.Add($user.DisplayName)
        }
    }
}



# Функция для получения списка баз данных почтовых ящиков
function Get-MailboxDatabases {
    try {
        $databases = Get-MailboxDatabase | Select-Object -ExpandProperty Name
        return $databases
    } catch {
        Write-Host "Ошибка при получении баз данных почтовых ящиков: $_" -ForegroundColor Red
        return @()
    }
}



# Функция для создания почтового ящика
function Create-Mailbox {
    param (
        [string]$Alias,
        [string]$UserFullName,
        [string]$MailboxDatabase
    )

    try {
        # Поиск пользователя в AD
        $user = Get-ADUser -Filter { DisplayName -eq $UserFullName } -Properties SamAccountName
        if ($user) {
            # Создание почтового ящика
            Enable-Mailbox -Identity $user.SamAccountName -Alias $Alias -Database $MailboxDatabase
            [System.Windows.Forms.MessageBox]::Show("Почтовый ящик успешно создан для пользователя $UserFullName с алиасом $Alias.", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            # Вывод в консоль об успешном добавлении
            Write-Host "Почтовый ящик успешно создан для пользователя $UserFullName" -ForegroundColor Green

            # Обновить ListBox после создания почтового ящика
            #Update-ListBox
        } else {
            [System.Windows.Forms.MessageBox]::Show("Пользователь с именем $UserFullName не найден.", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка при создании почтового ящика: $_", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}














# Создание формы
$form = New-Object System.Windows.Forms.Form
$form.Text = "Настройка почтового ящика"
$form.Size = New-Object System.Drawing.Size(690, 460)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle # Установка фиксированного размера формы
$form.MaximizeBox = $false # Отключение кнопки максимизации

# Устанавливаем стиль и размер шрифта для всех элементов формы
$form.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)


$scriptPath = $PSScriptRoot


# Установка иконки
$iconPath = Join-Path -Path $scriptPath -ChildPath "images\ex.ico" # Укажите путь к вашей иконке
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)

# Загружаем изображение из файла (замените путь на свой)
$imagePath = Join-Path -Path $scriptPath -ChildPath "images\bg.jpg"
$image = [System.Drawing.Image]::FromFile($imagePath)

# Устанавливаем изображение как фон формы
$form.BackgroundImage = $image
$form.BackgroundImageLayout = "Stretch"  # Растягиваем изображение на всю форму



# Создание заголовка
$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Text = "v1.2 (02.07.2025)" 
$labelVersion.Location = New-Object System.Drawing.Point(0, 0)
$labelVersion.Font = New-Object System.Drawing.Font("Arial", 7.5, [System.Drawing.FontStyle]::Bold)  # Увеличение размера шрифта и жирный шрифт
$labelVersion.AutoSize = $true  # Автоматический размер под текст
$labelVersion.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelVersion.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelVersion)



# Создание заголовка
$labelTitle = New-Object System.Windows.Forms.Label
$labelTitle.Text = "CREATE MAILBOX"
$labelTitle.Location = New-Object System.Drawing.Point(180, 40)
$labelTitle.Font = New-Object System.Drawing.Font("Arial", 26, [System.Drawing.FontStyle]::Bold)  # Увеличение размера шрифта и жирный шрифт
$labelTitle.AutoSize = $true  # Автоматический размер под текст
$labelTitle.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelTitle.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelTitle)



# Метка для списка пользователей без почты
$labelUsers = New-Object System.Windows.Forms.Label
$labelUsers.Text = "Пользователи без почтового ящика:"
$labelUsers.Location = New-Object System.Drawing.Point(60, 120)
$labelUsers.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelUsers.AutoSize = $true  # Автоматический размер под текст
$labelUsers.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelUsers)



# ListBox для отображения пользователей без почтового ящика
$listBoxUsers = New-Object System.Windows.Forms.ListBox
$listBoxUsers.Location = New-Object System.Drawing.Point(60, 150)
$listBoxUsers.Size = New-Object System.Drawing.Size(270, 150)
$listBoxUsers.SelectionMode = [System.Windows.Forms.SelectionMode]::One
$form.Controls.Add($listBoxUsers)


# Обработчик события для выбора пользователя из ListBox
$listBoxUsers.Add_SelectedIndexChanged({
    $selectedUser = $listBoxUsers.SelectedItem
    if ($selectedUser) {
        $user = Get-ADUser -Filter { DisplayName -eq $selectedUser } -Properties SamAccountName
        if ($user) {
            $textBoxAlias.Text = $user.SamAccountName
        } else {
            $textBoxAlias.Text = ""
        }
    }
})


# Метка и текстовое поле для алиаса
$labelAlias = New-Object System.Windows.Forms.Label
$labelAlias.Text = "Alias:"
$labelAlias.Location = New-Object System.Drawing.Point(380, 150)
$labelAlias.AutoSize = $true  # Автоматический размер под текст
$labelAlias.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelAlias.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelAlias)

$textBoxAlias = New-Object System.Windows.Forms.TextBox
$textBoxAlias.Location = New-Object System.Drawing.Point(450, 150)
$textBoxAlias.Size = New-Object System.Drawing.Size(170, 20)
$form.Controls.Add($textBoxAlias)



# Метка и комбобокс для выбора базы данных почтовых ящиков
$labelDatabase = New-Object System.Windows.Forms.Label
$labelDatabase.Text = "База данных почтовых ящиков:"
$labelDatabase.Location = New-Object System.Drawing.Point(380, 190)
$labelDatabase.AutoSize = $true  # Автоматический размер под текст
$labelDatabase.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelDatabase.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelDatabase)

$comboBoxDatabase = New-Object System.Windows.Forms.ComboBox
$comboBoxDatabase.Location = New-Object System.Drawing.Point(450, 220)
$comboBoxDatabase.Size = New-Object System.Drawing.Size(170, 20)
$comboBoxDatabase.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$form.Controls.Add($comboBoxDatabase)

# Загрузка данных в ComboBox
$comboBoxDatabase.Items.AddRange((Get-MailboxDatabases))






# Кнопка для создания почтового ящика
$buttonCreateMailBox = New-Object System.Windows.Forms.Button
$buttonCreateMailBox.Text = "СОЗДАТЬ"
$buttonCreateMailBox.Location = New-Object System.Drawing.Point(475, 280)

$buttonCreateMailBox.Width = 120      # Устанавливаем ширину кнопки
$buttonCreateMailBox.Height = 45      # Устанавливаем высоту кнопки
$buttonCreateMailBox.BackColor = [System.Drawing.Color]::Silver  # Устанавливаем цвет фона кнопки
$buttonCreateMailBox.ForeColor = [System.Drawing.Color]::Green      # Устанавливаем цвет текста кнопки
$buttonCreateMailBox.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # Устанавливаем выравнивание текста по правому краю

# Установка курсора при наведении
$buttonCreateMailBox.Cursor = [System.Windows.Forms.Cursors]::Hand

# Установка стиля кнопки на Flat и настройка рамки
$buttonCreateMailBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonCreateMailBox.FlatAppearance.BorderColor = [System.Drawing.Color]::DarkGreen
$buttonCreateMailBox.FlatAppearance.BorderSize = 2

$buttonCreateMailBox.Font = New-Object System.Drawing.Font("Arial", 10.2, [System.Drawing.FontStyle]::Bold)

# Определение относительного пути к иконке
$addmailPath = Join-Path -Path $scriptPath -ChildPath "images\add-mail.png"

# Загрузка и установка иконки для кнопки
$addmail = [System.Drawing.Image]::FromFile($addmailPath)
$buttonCreateMailBox.Image = $addmail

# Устанавливаем выравнивание иконки
$buttonCreateMailBox.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft

# Устанавливаем отступ справа для иконки
$buttonCreateMailBox.Padding = New-Object System.Windows.Forms.Padding(5, 0, 3, 0)

$form.Controls.Add($buttonCreateMailBox)





# Кнопка для применения настроек почтового ящика
$buttonApplySettings = New-Object System.Windows.Forms.Button
$buttonApplySettings.Location = New-Object System.Drawing.Point(60, 320)
$buttonApplySettings.Size = New-Object System.Drawing.Size(120, 45)
$buttonApplySettings.Text = "ПРИМЕНИТЬ НАСТРОЙКИ"
$buttonApplySettings.BackColor = [System.Drawing.Color]::Silver  # Устанавливаем цвет фона кнопки
$buttonApplySettings.ForeColor = [System.Drawing.Color]::Blue    # Устанавливаем цвет текста кнопки
$buttonApplySettings.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # Устанавливаем выравнивание текста по правому краю

$buttonApplySettings.Cursor = [System.Windows.Forms.Cursors]::Hand

# Установка стиля кнопки на Flat и настройка рамки
$buttonApplySettings.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonApplySettings.FlatAppearance.BorderColor = [System.Drawing.Color]::RoyalBlue
$buttonApplySettings.FlatAppearance.BorderSize = 2

$buttonApplySettings.Font = New-Object System.Drawing.Font("Arial", 8.6, [System.Drawing.FontStyle]::Bold)

# Определение относительного пути к иконке
$addsetPath = Join-Path -Path $scriptPath -ChildPath "images\add-set.png"

# Загрузка и установка иконки для кнопки
$addset = [System.Drawing.Image]::FromFile($addsetPath)
$buttonApplySettings.Image = $addset

# Устанавливаем выравнивание иконки
$buttonApplySettings.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft

# Устанавливаем отступ справа для иконки
$buttonApplySettings.Padding = New-Object System.Windows.Forms.Padding(2, 0, 2, 0)

$form.Controls.Add($buttonApplySettings)





# Кнопка для применения разрешений
$buttonSetPermissions = New-Object System.Windows.Forms.Button
$buttonSetPermissions.Location = New-Object System.Drawing.Point(210, 320)
$buttonSetPermissions.Size = New-Object System.Drawing.Size(120, 45)
$buttonSetPermissions.Text = "УСТАНОВИТЬ РАЗРЕШЕНИЯ"

$buttonSetPermissions.BackColor = [System.Drawing.Color]::Silver  # Устанавливаем цвет фона кнопки
$buttonSetPermissions.ForeColor = [System.Drawing.Color]::Blue      # Устанавливаем цвет текста кнопки
$buttonSetPermissions.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # Устанавливаем выравнивание текста по правому краю

$buttonSetPermissions.Cursor = [System.Windows.Forms.Cursors]::Hand

# Установка стиля кнопки на Flat и настройка рамки
$buttonSetPermissions.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonSetPermissions.FlatAppearance.BorderColor = [System.Drawing.Color]::RoyalBlue
$buttonSetPermissions.FlatAppearance.BorderSize = 2

$buttonSetPermissions.Enabled = $false
#$buttonSetPermissions.Enabled = $true

$buttonSetPermissions.Font = New-Object System.Drawing.Font("Arial", 8.6, [System.Drawing.FontStyle]::Bold)

# Определение относительного пути к иконке
$addpermPath = Join-Path -Path $scriptPath -ChildPath "images\add-per.png"

# Загрузка и установка иконки для кнопки
$addperm = [System.Drawing.Image]::FromFile($addpermPath)
$buttonSetPermissions.Image = $addperm

# Устанавливаем выравнивание иконки
$buttonSetPermissions.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft

# Устанавливаем отступ справа для иконки
$buttonSetPermissions.Padding = New-Object System.Windows.Forms.Padding(1, 0, 2, 0)

$form.Controls.Add($buttonSetPermissions)














# Обработчик события для кнопки "Создать"
$buttonCreateMailBox.Add_Click({
    $selectedItem = $listBoxUsers.SelectedItem
    if ($selectedItem) {
        $fullName = $selectedItem
        $alias = $textBoxAlias.Text
        $database = $comboBoxDatabase.SelectedItem
        if ($alias -and $database) {
            Create-Mailbox -Alias $alias -UserFullName $fullName -MailboxDatabase $database
            Update-ListBox
        } else {
            [System.Windows.Forms.MessageBox]::Show("Пожалуйста, заполните все поля.", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Пожалуйста, выберите пользователя из списка.", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    Update-ListBox
})


# Изначальное обновление списка при запуске формы
Update-ListBox





# Обработчик события для кнопки "Применить настройки"
$buttonApplySettings.Add_Click({
    $UserAlias = $textBoxAlias.Text


    $secondsRemaining = 15  # Время ожидания в секундах

Write-Host "Ожидание $secondsRemaining секунд"

while ($secondsRemaining -gt 0) {
    Write-Host -NoNewline "$secondsRemaining "
    Start-Sleep -Seconds 1
    $secondsRemaining--
}

Write-Host "Завершено! Продолжаем выполнение"


    if (Get-Mailbox -Identity $UserAlias -ErrorAction SilentlyContinue) {
        Get-Mailbox -Identity $UserAlias | Set-MailboxRegionalConfiguration -Language ru-ru -TimeZone "Belarus Standard Time" -LocalizeDefaultFolderName:$true
        [System.Windows.Forms.MessageBox]::Show("Настройки успешно применены для пользователя $UserAlias.", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        $mailboxConfig = Get-Mailbox -Identity $UserAlias | Get-MailboxRegionalConfiguration | Format-List | Out-String
        Write-Host "Настройки успешно применены для пользователя $UserAlias." -ForegroundColor Green
        Write-Host "Результат конфигурации почтового ящика для пользователя $UserAlias :"
        Write-Host $mailboxConfig -ForegroundColor Cyan

        # Активировать кнопку $buttonSetPermissions после успешного выполнения действий
        $buttonSetPermissions.Enabled = $true

    } else {
        [System.Windows.Forms.MessageBox]::Show("Пользователь $UserAlias не найден. Пожалуйста, убедитесь, что имя пользователя указано корректно.", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})





# Обработчик события для кнопки "Установить разрешения"
$buttonSetPermissions.Add_Click({


    # Деактивируем кнопку $buttonSetPermissions снова, чтобы она была недоступна до следующего нажатия $buttonApplySettings
    $buttonSetPermissions.Enabled = $false

    $targetUserAlias = $textBoxAlias.Text

    if (Get-Mailbox -Identity $targetUserAlias -ErrorAction SilentlyContinue) {
        try {
            Add-MailboxFolderPermission -Identity "${targetUserAlias}:\Календарь" -User sp_sync -AccessRights Editor -ErrorAction Stop
            [System.Windows.Forms.MessageBox]::Show("Разрешения успешно установлены для пользователя $targetUserAlias", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Write-Host "Настройки успешно применены для пользователя $targetUserAlias." -ForegroundColor Green
            Write-Host "Текущие разрешения для календаря:" 

            $permissions = Get-MailboxFolderPermission -Identity "${targetUserAlias}:\Календарь"
            $permissions | Format-List -Property RunspaceId, Identity, FolderName, User, AccessRights, IsValid, ObjectState | Out-String | Write-Host -ForegroundColor Cyan
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Ошибка при применении настроек: $($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Пользователь $targetUserAlias не найден. Пожалуйста, убедитесь, что имя пользователя введено корректно.", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})






# Загрузка данных в ListBox при запуске формы
Update-ListBox

# Показ формы
$form.ShowDialog()

# Закрытие сессии после использования формы
Remove-PSSession $Session

