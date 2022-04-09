#Set up a working folder.
$MyLocation = Split-Path $MyInvocation.MyCommand.Path
Set-Location -Path $MyLocation

#Today's date
$EToday = Get-Date -Format "yyyy/MM/dd"

#Log output Type
$LTBoth = "B"  # "B" means writing to console and file.
$LTCSL = "C"   # "C" means writing to console only. 
$LTFile = "F"  # "F" means writing to log file only.

# Log File Path
$LPath = ".\logs\SendMailCheckingSchedule_" + $EToday.Replace("/","").Remove(6,2) + ".log"


# Create Log folder
if ($true -ne (Test-Path (".\logs"))) {
    New-Item logs -ItemType Directory
}

# Log Output
Function WLog {
    foreach ($a in $args) {
        $MType = $a[0] #Log output Type
        $OMsg = (Get-Date).ToString() + " " + $a[1] # Log output time, and Log message.
    }

    if ($MType -ne $LTFile) {
        # Write to console.
        Write-Host $OMsg
    }
    if ($MType -ne $LTCSL) {
        # Write to Log file.
        Write-Output $OMsg | Out-File $LPath -Append
    }
}
Function ErrHandling {
    foreach ($msg in $args) {
        #Invoke-Item -Path .\Error.docs

        Add-Type -Assembly System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($msg)
    }
}

#Check the Schedule File
$scheduleFilePath = $MyLocation + "\MySchedule.xlsx"
if (!(Test-Path $scheduleFilePath)) {
    #File is nothing
    $ErrorMsg = "Schedule File is not found."
    WLog($LTBoth, $ErrorMsg)
    #Invoke-Item -Path .\Error.docs
    #Add-Type -Assembly System.Window.Forms
    #[System.Windows.Forms.MessageBox]::Show($ErrorMsg)
    ErrHandling($ErrorMsg)

    exit 999
} else {
    #File exists
    WLog($LTBoth, "Schedule File exists.")
}

#Open the file without displaying.
$ExcelObj = New-Object -ComObject Excel.Application
$ExcelObj.Visible = $false

$myBook = $ExcelObj.Workbooks.Open($scheduleFilePath)
$tSheet = $ExcelObj.Worksheets.Item(1)

$DateLine = 2  #Date
$ValueLine = 4 #出社 or テレワーク
$column = 3    #Check Start Column

while($true) {
    #
    $cellValue = $tSheet.Cells.Item($DateLine, $column).Value()
    Write-Host $cellValue
    #
    if([string]::IsNullOrEmpty($cellValue)) {
        $ErrorMsg = "最終行に到達しました。スケジュールエクセルを確認してください"
        WLog($LTBoth, $ErrorMsg)
        # Error発生時の処理
        #Invoke-Item -Path .\Error.docx
        ErrHandling($ErrorMsg)
        $ExcelObj.Quit()
        pause
        exit 999
    }

    #
    $cellDate = $cellValue.ToString("yyyy/MM/dd")


    if ($cellDate -eq $EToday) {
        #
        $tValue = $tSheet.Cells.Item($ValueLine, $column).Text
        if ($tValue.Contains("テレ")) {
            WLog($LTBoth, "テレワーク！！")
            break
        } else {
            WLog($LTBoth, "出社～")
            $ExcelObj.Quit()
            pause
            exit 999
        }
    }
    #Next Date
    $column++
}
$ExcelObj.Quit()


#Wait 60 seconds
#Start-Sleep -Seconds 60

#
$OutlookProcess = Get-Process | Where-Object {$_.Name -match "OUTLOOK"}
if ($OutlookProcess -eq $null) {
    $ErrorMsg = "Outlook is not running."
    WLog($LTBoth, $ErrorMsg)
    #Invoke-Item -Path .\Error.docx
    ErrHandling($ErrorMsg)
    exit 999
} else {
    WLog($LTBoth, "Outlook is running.")
}


#メールの作成
WLog($LTBoth,"メールを作成します")

$OutlookObj = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")

$NewMail = $OutlookObj.CreateItem(0)
$NewSubject = "定期送信(作成)メール" + $EToday.Remove(0,5)
$NewMail.Subject = $NewSubject
$NewMail.Body = [String]::Join("`r`n",(Get-Content .\MailBody.txt))
$NewMail.Recipients.Add("sample@mail.sample.com")
$NewMail.Close(0)
WLog($LTBoth,"メールの作成が完了しました。")

#
$namespace = $OutlookObj.GetNamespace("MAPI")
#
$folder = $namespace.GetDefaultFolder(16)

foreach ($item in $folder.Items) {
    if($item.Subject -eq $NewSubject) {
        $item.Display(0)
    }
}

$FlgFile = "" + $EToday.Replace("/","-")
Add-Content -Path $FlgFile -Value $null
WLog($LTBoth, "テレワーク実施フラグを出力しました。")

Write-Output $MyLocation