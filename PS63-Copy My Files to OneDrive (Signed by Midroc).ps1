<#  ¯\_(ツ)_/¯
    .NAME
        PS63-Save My Files To OneDrive
    .VERSION
        2.5
    .AUTHOR
        Magnus Persson
    .SYNOPSIS
        Programmet kopierar användarens standard foldrar och filer till OneDrive
    .DESCRIPTION
        Innan du kör vidare med det här programmet bör du se till att du har
        installerat den senaste versionen av OneDrive på datorn.
        Det gör du enklast genom att starta Software Center och installera om OneDrive.

        Det här programmet hjälper sedan till att säkerhetskopiera viktiga filer på din dator.
        Filerna sparas lokalt i OneDrive under en ny folder som heter:
        My Files (Backup made by PowerShell)
        $Log.NewLine()
        Men tänk på att det sedan tar tid innan alla filerna har synkat upp till OneDrive i molnet.
        
        Det är alltid du som har ansvaret för att se till att inga av dina viktiga filer försvinner
        ex i samband med ominstallation av datorn.
    .NOTE
        Programmet måste vara signerat för att gå att köra på klienter.
        Programmet använder klasser, och fungerar därför bara på Windows 10 datorer med PowerShell 5.0
#>


Function Copy-FolderToOneDrive{
    [CmdletBinding()]  
    Param(  
        [string]$From,
        [string]$To
    ) 
     
    try{
        # kolla att From-foldern finns
        if (!(Test-Path $From -PathType Container)){
            $Log.Text('','Hittar inte Foldern','$From')
            $Log.Text('','Ingen kopiering kan genomföras!','Avslutar...')
            Return
        }

        $Log.Text('','Kopiera från',$From)

        # ta fram namnet på nya foldern
        $Folders = $From.Split('\')
        $Folder = $Folders[-1]
        
        # fullständiga sökvögen blir då
        $ToComplete = $To + '\' + $Folder

        # Skapa den nya To-foldern, om den redan råkar finnas så skapas foldern inte om och det kommer inget felmeddelande!
        $Log.Text('','Kopiera till',$ToComplete)
        
        # kopiera foldern, men ta INTE med några security rättigheter
        Copy-Item -Path $From -Destination $To -Recurse -Force
        $Log.Line()
    
    }
    catch{
        $Log.Line()
        $global:Errors++
        $Log.Error($Error)
        $Log.Line()
        Return
    }
}


function Copy-FilesToOneDrive{
    [CmdletBinding()]      
    Param(  
        [string]$From,
        [string]$To,
        [string]$FileFilter
    ) 
    try{    
        
        # om det inte finns någon folder så avsluta
        if (-not (Test-Path $From -PathType Container)){
            $Log.Text('','Hittade inte foldern',$From)
            $Log.Line()
            Return    
        }        

        # kolla om det finns några filer
        $Files = Get-ChildItem -Path $From -Filter $FileFilter

        if($Files -eq $null){
            $Log.Text('','Hittade inte någon fil under',$From)
            $Log.Line()
            Return    
        }
                
        $FilesCount = $Files.Count
        $Log.Text('','Antal filer som hittades',$FilesCount)
        $Log.Line()
        
        # skapa en ny folder i OneDrive
        [Void][System.IO.Directory]::CreateDirectory($To)

        $Log.Text('','Kopiera från',$From)
        $Log.Text('','Kopiera till',$To)
        Copy-Item -Path "$From\$FileFilter" -Destination $To -Force
        $Log.Line()

    }
    catch{
        $Log.Line()
        $global:Errors++
        $Log.Error($Error)
        $Log.Line()
        Return
    }
}



function Copy-PSTFilesToOneDrive{
    [CmdletBinding()]      
    Param(  
        [string]$From,
        [string]$To,
        [string]$FileFilter
    ) 
    try{    
        
        $Log.Text('','Leta .PST filer, men bara på OSDisk',$From)

        # om det inte finns någon folder så avsluta
        if (-not (Test-Path $From -PathType Container)){
            $Log.Text('','Hittade inte denna OSDisk','Avslutar...')
            $Log.Line()
            Return    
        }        

        # leta igenom hela C: efter PST filer
        $Files = Get-ChildItem -Path $From -Filter $FileFilter -Recurse -ErrorAction SilentlyContinue
        
        
        if($Files -eq $null){
            $Log.Text('','Hittade inte någon PST fil','Avslutar...')
            $Log.Line()
            Return    
        }
        
        # skapa en ny folder i OneDrive
        [Void][System.IO.Directory]::CreateDirectory($To)

        $FilesCount = $Files.Count
        $Log.Text('','Antal filer som hittades',$FilesCount)
        $Log.Line()
                
        # kopiera alla filer till samma folder
        ForEach ($File in $Files){
            try{
                $From = $File.FullName
                $FilePath = $File.DirectoryName
                $Log.Text('','Kopiera från',$From)

                # hoppa över alla pst-filer som redan ligger under 'My Files (Backup made by PowerShell)\' mappen
                if($FilePath -eq $To){
                    $Log.Text('','Kopiera inte den här filen!','Den ligger redan i OneDrive...')
                    $Log.Line()
                    Continue
                } 

                # för att det inte ska bli krockar mellan flera PST-filer med samma namn så gör namnet unikt
                $FileLastWriteTime = $File.LastWriteTime
                $FileLastWriteTime = $FileLastWriteTime -replace "/", "-"
                $FileLastWriteTime = $FileLastWriteTime -replace ":", "."
                
                # lägg till ett datum på filnamnet för att undvika krockar när alla filer ska ligga i samma folder
                $FileName = $File.Name
                $FileNameNew = "($FileLastWriteTime) $FileName"

                $Log.Text('','Kopiera till',"$To")
                $Log.Text('','Filen döptes om till',"$FileNameNew")
                Copy-Item -Path "$From" -Destination "$To\$FileNameNew" -Force
                $Log.Line()
            }

            # egen felhanterare i loopen, som bara går vidare till nästa fil
            catch{
                $Log.Line()
                $global:Errors++
                $Log.Error($Error)
                $Log.Line()
            }
        }
    }
    catch{
        $Log.Line()
        $global:Errors++
        $Log.Error($Error)
        $Log.Line()
        Return
    }
}



function Console-ShowGuideOrNot{
    [CmdletBinding()]      
    Param() 
    try{

        # Visa Guiden?
        $Loop = $true
        while ($Loop){
        
            Clear-Host
            Write-Host ""
            Write-Host "  Hej!" -ForegroundColor Yellow
            Write-Host ""
            Write-Host '  För att det här programmet ska fungera så måste senaste versionen av OneDrive finnas installerad på din dator.' -ForegroundColor Yellow
            Write-Host ""
            Write-Host '  Vill du avsluta programmet, och istället öppna guiden "OneDrive - Börja Använda Nya OneDrive" ?' -ForegroundColor Yellow
            Write-Host ""
            # visa guide?
            $answer = Read-Host '  Välj [J/N]'
                
            if ($answer -eq 'j' -or $answer -eq 'y'){
                Clear-Host
                Write-Output ""
                Write-Host '  Öppnar guiden...' -ForegroundColor Yellow
                sleep 2
                start 'https://servicedesk.midroc.se/hc/sv/articles/115001787345'
                $Loop = $false
                Exit
            }

            # kolla om vi bara ska avsluta hela programmet!
            if ($answer -eq 'n'){
                $Loop = $false
                Return
            }
        }
    }
    catch{
        $Log.Line()
        $global:Errors++
        $Log.Error($Error)
        $Log.Line()
        Return
    }
}


function Console-RunTheProgramOrNot{
    [CmdletBinding()]      
    Param() 
    try{

        # skriv till Console-fönstret
        $Loop = $true
        while ($Loop){
            Clear-Host
            Write-Host ""
            Write-Host '  Det här programmet hjälper till att säkerhetskopiera viktiga filer på din dator.' -ForegroundColor Yellow
            Write-Host ""
            Write-Host '  Filerna sparas lokalt i OneDrive under en ny folder som heter:' -ForegroundColor Yellow
            Write-Host '  My Files (Backup made by PowerShell)' -ForegroundColor Yellow
            Write-Host ""
            Write-Host '  Men tänk på att det sedan tar tid innan alla filerna har synkat upp till OneDrive i molnet.' -ForegroundColor Yellow
            Write-Host ""
            Write-Host '  Det är alltid du som har ansvaret för att se till att inga av dina viktiga filer försvinner' -ForegroundColor Yellow
            Write-Host '  ex i samband med ominstallation av datorn.' -ForegroundColor Yellow
            Write-Host ""
            Write-Host '  Har du några frågor så kontakta gärna servicedesk (tel. 010-470 70 01)' -ForegroundColor Yellow
            Write-Host ""
            Write-Host '  Vill du starta programmet och kopiera över dina filer till din OneDrive ?' -ForegroundColor Yellow
            Write-Host ""

            # fortsätt med programmet?
            $answer = Read-Host '  Välj [J/N]'

            # kolla om vi bara ska avsluta hela programmet!
            if ($answer -eq 'n'){
                Clear-Host
                Write-Output ""
                Write-Host "  Programmet avslutas..." -ForegroundColor Yellow
                sleep 2
                Exit
            }

            # Gå vidare?
            if ($answer -eq 'j' -or $answer -eq 'y'){
                $Loop = $False
                Return
            }
        }
    }
    catch{
        $Log.Line()
        $global:Errors++
        $Log.Error($Error)
        $Log.Line()
        Return
    }
}


# ------------------------
# PROGRAMMET BÖRJAR HÄR!
# ------------------------

try{    
    Clear-Host
    
    # förbered för att spara log-filen under %Temp%, skapa en ny folder
    $TempPowerShellFolder = $env:temp + '\PowerShell'
    [Void][System.IO.Directory]::CreateDirectory($TempPowerShellFolder)
    
    # skapa ett nytt Logg-object, skicka alltid med sökvägen till ps1-filen, så att klassen kan räkna ut vad loggfilen ska heta
    Import-Module $($PSScriptRoot + '\Module\Midroc Log.psm1') -Force
    $Log = New-MidrocLog($PSCommandPath)

    # justera inställningar, loggfilen ska sparas under %Temp% ex C:\Users\mssmape\AppData\Local\Temp
    $Log.WidthText1 = 2
    $Log.WidthText2 = 36
    $Log.LineWidth = 155
    $log.FolderPath = $TempPowerShellFolder
    
    # default loggning 'SilentlyContinue' och felhantering 'Continue'
    $VerbosePreference = 'SilentlyContinue'
    $ErrorActionPreference = 'Continue'

    # sätt standard variabler, OBS inga $global:Warnings används i programmet!
    $global:Errors = 0
        
    # spara användarens namn
    $UserName = $env:UserName
    $ComputerName = $env:ComputerName


    # ska guiden visas?
    Console-ShowGuideOrNot
    
    # starta programmet?
    Console-RunTheProgramOrNot


    # ---------------------------------------------------------
    # innan programmet startar så kolla att OneDrive finns
    # ---------------------------------------------------------
    $OneDriveExist = $false

    # sökvägen till användarens OneDrive (OBS Långt '–' i namnet "C:\Users\mssmape\OneDrive – onmidroc"
    $OneDrive = 'C:\Users\' + $UserName + '\OneDrive – onmidroc'
    
    # Kolla om folden finns
    if (Test-Path $OneDrive -PathType Container){
        $OneDriveExist = $true
    }
    else{
        # om sökvägen inte fanns så testa igen men... (OBS Kort '-' i namnet "C:\Users\mssmape\OneDrive - onmidroc"
        $OneDrive = 'C:\Users\' + $UserName + '\OneDrive - onmidroc'

        # Kolla om den här folden finns!
        if (Test-Path $OneDrive -PathType Container){
            $OneDriveExist = $true
        }
    }

    # TEST!!!
    ##$OneDriveExist = $false
        
    # om OneDrive foldern saknas så avsluta programmet!
    if (!$OneDriveExist){
        Clear-Host
        Write-Host ""
        Write-Host "  Varning - Senaste versionen av OneDrive verkar inte vara installerad!" -ForegroundColor Yellow
        Write-Host "  Programmet kommer att avslutas." -ForegroundColor Yellow
        Write-Host ""
        Write-Host '  Har du några frågor så kontakta gärna servicedesk (tel. 010-470 70 01)' -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Tryck Enter..." -ForegroundColor Yellow
        $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp") > $null
        Exit
    }
       
    
    # ---------------------------------------------------------
    # allt verkar ok så programmet kan nu starta!
    # ---------------------------------------------------------
    Clear-Host
    Write-Host ""        
    Write-Host "  Ok, då fortsätter vi!" -ForegroundColor Yellow
    Write-Host "  Tänk på att det tar en stund att skanna av din dator..." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  När programmet sedan är klart så öppnas en loggfil automatiskt, så att kan se vad som kopierades." -ForegroundColor Yellow
    sleep 5
                    
    # skriv till loggfilen
    $Log.Text('','Hej!')
    $Log.NewLine()
    $Log.Text('','För att det här programmet ska fungera så måste senaste versionen av OneDrive finnas installerad på din dator.')    
    $Log.NewLine()
    $Log.Text('','Programmet hjälper till att säkerhetskopiera viktiga filer på din dator.')
    $Log.NewLine()
    $Log.Text('','Filerna sparas lokalt i OneDrive under en ny folder som heter:')
    $Log.Text('','My Files (Backup made by PowerShell)')
    $Log.NewLine()
    $Log.Text('','Det är alltid du som har ansvaret för att se till att inga av dina viktiga filer försvinner')
    $Log.Text('','ex i samband med ominstallation av datorn.')
    $Log.NewLine()
    $Log.Text('','Har du några frågor så kontakta gärna servicedesk (tel. 010-470 70 01).')
    
    $Log.TextHeader('','KOPIERA STANDARD FOLDRAR')
    
    # Skapa en ny folder i OneDrive ex 'C:\Users\mssmape\OneDrive – onmidroc\Copy Of Standard Folders'
    $OneDriveCopyOfStandardFolders = $OneDrive + '\My Files (Backup made by PowerShell)'
    [Void][System.IO.Directory]::CreateDirectory($OneDriveCopyOfStandardFolders)
     

    # Kopiera alla Standardfoldrar
    $From = [Environment]::GetFolderPath("Desktop")
    Copy-FolderToOneDrive -From $From -To $OneDriveCopyOfStandardFolders

    $From = [Environment]::GetFolderPath("MyDocuments")
    Copy-FolderToOneDrive -From $From -To $OneDriveCopyOfStandardFolders

    $From = [Environment]::GetFolderPath("MyMusic")
    Copy-FolderToOneDrive -From $From -To $OneDriveCopyOfStandardFolders

    $From = [Environment]::GetFolderPath("MyVideos")
    Copy-FolderToOneDrive -From $From -To $OneDriveCopyOfStandardFolders
        
    $From = [Environment]::GetFolderPath("MyPictures")
    Copy-FolderToOneDrive -From $From -To $OneDriveCopyOfStandardFolders

    $From = [Environment]::GetFolderPath("Favorites")
    Copy-FolderToOneDrive -From $From -To $OneDriveCopyOfStandardFolders
    
    # kopiera Exchange Autocomplete filer
    $Log.TextHeader('','KOPIERA OUTLOOK AUTOCOMPLETE')
    $From = $env:localappdata + '\Microsoft\Outlook\RoamCache'
    $To = $OneDriveCopyOfStandardFolders + '\Outlook AutoComplete'
    Copy-FilesToOneDrive -From $From -To $To -FileFilter 'Stream_Autocomplete_*'


    # kopiera Exchange Signatur filer
    $Log.TextHeader('','KOPIERA OUTLOOK SIGNATURER')
    $From = $env:appdata + '\Microsoft\Signatures'
    $To = $OneDriveCopyOfStandardFolders + '\Outlook Signatures'
    Copy-FilesToOneDrive -From $From -To $To -FileFilter '*.*'

    # kopiera Chrome Bookmarks
    $Log.TextHeader('','KOPIERA CHROME BOOKMARKS')
    $From = "$env:localappdata\Google\Chrome\User Data\Default"
    $To = $OneDriveCopyOfStandardFolders + '\Chrome Bookmarks'
    Copy-FilesToOneDrive -From $From -To $To -FileFilter 'Bookmarks'


    # hitta och kopiera alla Outlook .PST filer
    $Log.TextHeader('','KOPIERA OUTLOOK PST FILER')
    # OBS! '\' måste finnas efter C: annars fungerar inte sökningen när man sedan startar programmet via genväg!
    $FromFolder = "C:\"
    $ToFolder = $OneDriveCopyOfStandardFolders + '\Outlook PST'
    $Filter = "*.PST"
    Copy-PSTFilesToOneDrive -From $FromFolder -To $ToFolder -FileFilter $Filter

    # om man vill felsöka och pausa consolen innan programmet stänger ner så ta med de här 2 raderna!
    ### Write-Host "Press any key to continue..."    
    ### $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp") > $null     
    
}
catch
{
    # skriv ut fel och gå vidare ner till sammanfattningen
    $Log.Line()
    $global:Errors++
    $Log.Error($Error)
    $Log.Line()    
}

# -------------
# SUMMERING MM
# -------------
$Log.TextHeader('','SAMMANFATTNING')
$Log.Text('','Errors',$global:Errors)
$Log.Line()
$Log.Text('','Användarnamn',$UserName)
$Log.Text('','Datornamn',$ComputerName)
$Log.Line()


# sökvägen till loggfilen under %Temp%
$LogFileFrom = $log.FolderPath + '\' + $log.FileName

# kopiera logfilen till server under ex \MSSMAPE\LT-R9005WHN
try{
    
    $ServerComputerFolder = '\\midroc.net\PowerShell\Midroc Log\' + $ComputerName
    [Void][System.IO.Directory]::CreateDirectory($ServerComputerFolder)
    
    $ServerUserFolder = $ServerComputerFolder + '\' + $UserName
    [Void][System.IO.Directory]::CreateDirectory($ServerUserFolder)

    $LogFileTo = $ServerUserFolder + '\' + $log.FileName
    Copy-Item $LogFileFrom $LogFileTo -Force -ErrorAction SilentlyContinue
}
catch{
    $Log.Error($Error)
    $Log.Line()  
}


# flytta logfilen till Onedrive
if ($OneDriveExist){
    try{
        $OneDriveLogFolder = $OneDriveCopyOfStandardFolders + '\Log'
        [Void][System.IO.Directory]::CreateDirectory($OneDriveLogFolder)

        $LogFileTo = $OneDriveLogFolder + '\' + $log.FileName
        Move-Item $LogFileFrom $LogFileTo -Force
        Invoke-Item $LogFileTo
    }
    catch{
        $Log.Error($Error)
        $Log.Line()  
        Invoke-Item $LogFileTo
    }
}
else
{
    # öppna istället upp loggfilen direkt i %Temp% med Notepad
    Invoke-Item $LogFileFrom
}


# SIG # Begin signature block
# MIIH3QYJKoZIhvcNAQcCoIIHzjCCB8oCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUYinORSduTA1W/+n63sDXmfB2
# 1pqgggXUMIIF0DCCBLigAwIBAgIKdBR98AAAAACA4jANBgkqhkiG9w0BAQUFADBF
# MRMwEQYKCZImiZPyLGQBGRYDbmV0MRYwFAYKCZImiZPyLGQBGRYGbWlkcm9jMRYw
# FAYDVQQDEw1NaWRyb2MgQ0Egdi4yMB4XDTE3MDYxMjEwMzAyOFoXDTE4MDYxMjEw
# MzAyOFowgYgxEzARBgoJkiaJk/IsZAEZFgNuZXQxFjAUBgoJkiaJk/IsZAEZFgZt
# aWRyb2MxFjAUBgNVBAsTDUNsb3VkRXhjbHVkZWQxGTAXBgNVBAsTEFNlcnZpY2Ug
# QWNjb3VudHMxDTALBgNVBAsTBERBSVQxFzAVBgNVBAMTDk1HU0FEUG93ZVNoZWxs
# MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQDvPpoTWY9K3MZyLTccGpMRq40r
# FKv8cDn4HPHAtOOyyU9ngbI+BHCLGRrVb2I2yorwdzo4+/0jX91NN+bOBBC/oLC9
# NmO1o5Z3O07nL0+pagVp6+zPFED80Bbgm2GNORs82Kq857tQv8dImuFDjr/cFFhK
# O8mf6yD2HOT0ibIi1QIDAQABo4IDADCCAvwwJQYJKwYBBAGCNxQCBBgeFgBDAG8A
# ZABlAFMAaQBnAG4AaQBuAGcwEwYDVR0lBAwwCgYIKwYBBQUHAwMwCwYDVR0PBAQD
# AgeAMB0GA1UdDgQWBBRCHeYtpjGFVkSHtOt+LnQIwivIBDAfBgNVHSMEGDAWgBSe
# c2Ram0/AwQOKp0QWVNYY5ihJUDCCAREGA1UdHwSCAQgwggEEMIIBAKCB/aCB+oaB
# uWxkYXA6Ly8vQ049TWlkcm9jJTIwQ0ElMjB2LjIsQ049TUdTMTBTQTEzLENOPUNE
# UCxDTj1QdWJsaWMlMjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25m
# aWd1cmF0aW9uLERDPW1pZHJvYyxEQz1uZXQ/Y2VydGlmaWNhdGVSZXZvY2F0aW9u
# TGlzdD9iYXNlP29iamVjdENsYXNzPWNSTERpc3RyaWJ1dGlvblBvaW50hjxodHRw
# Oi8vbWdzMTBzYTEzLm1pZHJvYy5uZXQvQ2VydEVucm9sbC9NaWRyb2MlMjBDQSUy
# MHYuMi5jcmwwggEjBggrBgEFBQcBAQSCARUwggERMIGvBggrBgEFBQcwAoaBomxk
# YXA6Ly8vQ049TWlkcm9jJTIwQ0ElMjB2LjIsQ049QUlBLENOPVB1YmxpYyUyMEtl
# eSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9bWlk
# cm9jLERDPW5ldD9jQUNlcnRpZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2VydGlm
# aWNhdGlvbkF1dGhvcml0eTBdBggrBgEFBQcwAoZRaHR0cDovL21nczEwc2ExMy5t
# aWRyb2MubmV0L0NlcnRFbnJvbGwvTUdTMTBTQTEzLm1pZHJvYy5uZXRfTWlkcm9j
# JTIwQ0ElMjB2LjIuY3J0MDUGA1UdEQQuMCygKgYKKwYBBAGCNxQCA6AcDBpNR1NB
# RFBvd2VyU2hlbGxAbWlkcm9jLm5ldDANBgkqhkiG9w0BAQUFAAOCAQEAtoN/ls8c
# g2OZeqreMQZ7LCjpwe2fpcrNaMQxpwheUxOQsavSLW1xKmeJ2pqVQfVHFaNeiF2u
# Lpc8cCJx8/evSZdctugFrJKuX5ieLZ9qqUTrBdzh0JfEfnxK+73NrI+XmlJbrkPb
# kGs3zMUILvflUX5nRXA+es8DJnHtkRItUlfCW4xBTHxBXuAzKh48fT4TfZ2ula8H
# kvZETanpPyb3OJdF7AvAmisGklz+M1jQToR0XgO78UC51B8lLqSoFXjVqVHMh0ZX
# ovRIIQomPy/7V34Kwfy303NftUl8D0AmPZPojiku7jD4riOt6XNjSDY1BvkFgC61
# 4w8OvnXjJSDUAjGCAXMwggFvAgEBMFMwRTETMBEGCgmSJomT8ixkARkWA25ldDEW
# MBQGCgmSJomT8ixkARkWBm1pZHJvYzEWMBQGA1UEAxMNTWlkcm9jIENBIHYuMgIK
# dBR98AAAAACA4jAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKA
# ADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYK
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQULuAQ/yf2r/N7cUc7RvM1WPYcyj0w
# DQYJKoZIhvcNAQEBBQAEgYBQxC5R0Bh37q8Po+jJyvaoc/JgUtC4yh5FnsiTHaUu
# c5u19eKd2obuyz8w6E8TBfRz8obGLGzJdJyXvC0GL/gjGBaOS7VBRWPkms3BwvEU
# j3ueKSpXAgSBfU+xUmj1rrWsjZnyibM8J8rUtQgK8I55LuFfPAneY1bKeGj4y5+I
# AQ==
# SIG # End signature block
