#---SETUP---
clear-host
$sw = [Diagnostics.Stopwatch]::StartNew()
$ErrorActionPreference = "silentlycontinue"

$searchFor1 = "sødeste"
$searchFor2 = "brunette"
#$searchFor1 = '\d\d\d\d\d\d[-]\d\d\d\d'
$searchStrings = @($searchFor1,$searchFor2)
$global:folder = $null
$global:outputFile = $null

Main

function Main {
    createOutput
    $physicalDisks = Read-Host "Skal alle fysiske disk scannes (J/N)"
    if ($physicalDisks -eq "j") {
        findPhysicalDisks 
        findingFiles ($global:folder)       
        }
    elseif ($physicalDisks -eq "n") {
    $global:folder = Read-Host "Skriv drev og startmappe f.eks. C:\users\" 
    findingFiles ($global:folder)
    }
    else {
    Main
    }
}

function findingFiles ($folder) {
    $wordFiles = Get-ChildItem -Path $folder -Recurse | ? Name -like "*.do[c,t]*"
    $excelFiles = Get-ChildItem -Path $folder -Recurse | ? Name -like "*.xls*"
        foreach ($wordFile in $wordFiles) {
            $wordFile = $wordFile.FullName
            runWord ($wordFile)
        }
           
        foreach ($excelFile in $excelFiles) {
            $excelFile = $excelFile.FullName
            runExcel ($excelFile)
        }
}

function runWord ($wordFile) {
    foreach ($searchString in $searchStrings) {
        $word = New-Object -ComObject Word.Application
        $wordFile | foreach-object {
        $file = $wordFile
            if ($file -match '.docx') {
                if ($word.Documents.Open($file).Content.Find.Execute($searchString)) {                
                    write-host WARNING: $wordFile contains $searchString
                    $value = 'WARNING: '+$wordFile+' contains '+$searchString
                    Add-Content -Path $global:outputFile -Value $value
                }
            $word.Application.ActiveDocument.Close()
            } 
            else {
                if ((Get-Content $file | %{$_ -match $searchString }) -contains $true) {
                    write-host WARNING: $wordFile contains $searchString
                    $value = 'WARNING: '+$wordFile+' contains '+$searchString
                    Add-Content -Path $global:outputFile -Value $value
                }
            }
    $word.Application.quit(0)
        }
    }
}

function runExcel ($excelFile) {
  $excel = New-Object -ComObject Excel.Application
  $excelFile | foreach-object {
  $file = $excelFile
  if ($file -match '.xls') {
    if ($excel.Documents.Open($file).Content.Find.Execute($searchFor)) {
      write-host WARNING: $excelFile contains $searchFor
    }
    $excel.Application.ActiveDocument.Close()
  } else {
    if ((Get-Content $file | %{$_ -match $searchFor }) -contains $true) {
        write-host WARNING: $excelFile contains $searchFor
        write-host "global output" $global:outputFile
        #Add-Content $global:outputFile WARNING: $excelFile contains $searchFor
    }
  }
}
$excel.Application.quit(0)
    
}

$sw.Stop() 
$sw.Elapsed

#Create a outputfile and folder
function createOutput {
    $date = get-date -Format "yyyyMMdd_HHmmss"
    $name = "log"+$date+".txt"
    $dirpath = $HOME+'\PSScriptLogFiles\'
    if (Test-Path $dirpath) {     
        $output = new-item -path $dirpath -type file -name $name
        $global:outputFile = $dirpath+$name
        
        }
    else {
        $dirpath = new-item -path $HOME -type Directory -Name "PSScriptLogFiles"
        $output = new-item -path $dirpath -type file -name $name
        $global:outputFile = $dirpath+$name
        }
}
#Find physical disks on client function findPhysicalDisks {    $workfolder = '\'
    [array]$disk = Get-WmiObject -Class Win32_logicaldisk -Filter "DriveType =3"

    for ($i = 0; $i -lt $disk.length; $i++) {
        if ($disk[$i].DeviceId) {
            $global:folder = $disk[$i].DeviceId -replace ' ',''
            $global:folder = $global:folder+$workfolder
            }
    }
}