$csvFiles =  @()   #https://www.jonathanmedd.net/2014/01/adding-and-removing-items-from-a-powershell-array.html
$NameExtension = "MachineReadable"
$FolderName = "für Import"
$myPath = $PSScriptRoot #https://stackoverflow.com/questions/5466329/whats-the-best-way-to-determine-the-location-of-the-current-powershell-script/5466355#5466355

Write-Host "**** starting script ****" #https://stackoverflow.com/questions/707646/echo-equivalent-in-powershell-for-script-testing/707666#707666 

$Folder = $myPath + "\" + $FolderName
if (Test-Path $Folder){
    Remove-Item ($Folder+ "\*.*" ) -Force -ErrorAction SilentlyContinue -ErrorVariable ProcessError
}
if($ProcessError){
        Write-Host ("Files in  '" + $Folder+  "' cannot be deleted, aborting script.")
        exit
}

$numberFiles = ( Get-ChildItem $myPath   -Filter *.csv | Measure-Object ).Count #https://stackoverflow.com/questions/14714284/count-items-in-a-folder-with-powershell 
If ($numberFiles -eq 0) { 
    "There is not any csv in this folder, script is aborted" 
    exit
}

"The following csv files are found:"
Get-ChildItem $myPath   -Filter *.csv | 
Foreach-Object{
    $csvFiles = $csvFiles + $_.BaseName 
    "* " + $_.Name
} #https://stackoverflow.com/questions/18847145/loop-through-files-in-a-directory-using-powershell

function AmILastItemInLoop{
    If ($i -eq $numberFiles-1){ 
        "There is not any other csv tp process, ending script now." 
        Exit #https://stackoverflow.com/questions/2022326/terminating-a-script-in-powershell
    }
    else {
    }
 }

 function CreateMarkdownInNote{
    ("`"" + "The full description of the transaction as given in the column '" + $DescriptionColumnTitle + "':  `n" +  "> " + $originalDescription + "`"")
 }

 function checkLock {
    Param(
        [parameter(Mandatory=$true)]
        $filename
    )
    $file = gi (Resolve-Path $filename) -Force
    if ($file -is [IO.FileInfo]) {
        trap {
            return $true
            continue
        }
        $stream = New-Object system.IO.StreamReader $file
        if ($stream) {$stream.Close()}
    }
    return $false
}


Write-Host "**** starting processing ****"
for($i=0; $i -lt $numberFiles;$i++){
    $foundHeader = $false
    $csvFile_lineContent = @() 
    $DescriptionColumnTitle = ""
    $ProcessError = ""
    $stopProcessing = $false 
      
    Write-Host "**processing:" $csvFiles[$i]
    
    $newFile_FullName = $csvFiles[$i] + "-" + $NameExtension + ".csv"
    $newFile_FullName = $myPath + "\" + $FolderName + "\" + $newFile_FullName
    [void](New-Item $newFile_FullName -ItemType File -Force -ErrorAction SilentlyContinue -ErrorVariable ProcessError) #https://stackoverflow.com/questions/46586382/hide-powershell-output
    if($ProcessError){
        Write-Host ("Not able to write '" + $newFile_FullName +  "'")
        $ProcessError ="" 
        if (-Not (AmILastItemInLoop)){
            continue
        }
        else {
            break
        }            
    }
    $csvFileFullName = $myPath + "\" + $csvFiles[$i] + ".csv"      
    $csvFile = get-content ($csvFileFullName) -Force -ErrorAction SilentlyContinue -ErrorVariable ProcessError
    if($ProcessError){
        Write-Host ("Not able to read '" + $csvFileFullName +  "'")
        $ProcessError ="" 
        if (-Not (AmILastItemInLoop)){
            continue
        }
        else {
            break
        }
    }
    $newline = "FirstLine"  
    $csvFile | Where-Object { $stopProcessing -Eq $False } | ForEach-Object { #https://stackoverflow.com/questions/10277994/how-to-exit-from-foreach-object-in-powershell
        $NonEmptyFields = @() 
        $Fields = @()

        $Fields = $_.Split(';')   #https://stackoverflow.com/questions/53764858/powershell-read-text-file-line-by-line-and-split-on
        for($j=0; $j -lt $Fields.count;$j++){
            if ($fields[$j] -ne ""){
                $NonEmptyFields = $NonEmptyFields + $fields[$j]
            }
        }


        switch ($NonEmptyFields.Count){
            2 {
                continue
            }
            0 {
                AmILastItemInLoop
                $stopProcessing = $true
                continue 
            }
            default {
                $originalDescription = $Fields[1]
                if ($foundHeader){
                    if ($originalDescription[0] -eq '"'){ #https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_quoting_rules?view=powershell-7
                        $originalDescription = $originalDescription.substring(1) 
                    }
                    if ($originalDescription[$originalDescription.Length-1] -eq '"'){
                        $originalDescription = $originalDescription.substring("", $originalDescription.Length-1)  
                    }

                    if ($originalDescription -like '*XXXX*'){  #https://morgantechspace.com/2016/08/powershell-check-if-string-contains-word.html 
                        $digitLength = 4
                        if ($originalDescription.Substring($originalDescription.indexof(' XXXX')+ ' XXXX'.Length,$digitLength) -match "^\d+$"){ #https://stackoverflow.com/questions/51171410/check-if-string-contains-numeric-value-in-powershell                          
                            $tempText = CreateMarkdownInNote
                            $Fields = $Fields + $tempText
                            $newDescription = $originalDescription.Substring($originalDescription.indexof(' XXXX') + ' XXXX'.Length + $digitLength)
                            $Fields[1] = $newDescription.Trim()
                        }
                    }
                    elseif($originalDescription -like '*SENDER REFERENZ:*'){
                        $tempText = CreateMarkdownInNote
                        $Fields = $Fields + $tempText
                        $newDescription = $originalDescription.Substring($originalDescription.indexof('SENDER REFERENZ:')+ 'SENDER REFERENZ:'.Length)
                        $Fields[1] = $newDescription.Trim()
                    }
                    elseif($originalDescription -like '*COMMUNICATIONS:*'){
                        $tempText = CreateMarkdownInNote
                        $Fields = $Fields + $tempText
                        $newDescription = $originalDescription.Substring($originalDescription.indexof('COMMUNICATIONS:')+ 'COMMUNICATIONS:'.Length)
                        $Fields[1] = $newDescription.Trim()
                    }
                    elseif($originalDescription -like '*MITTEILUNGEN:*'){
                        $newDescription = $originalDescription.Substring($originalDescription.indexof('MITTEILUNGEN:')+ 'MITTEILUNGEN:'.Length)
                        $newDescription = $newDescription.Trim()
                        if($newDescription.Length -gt 0) {
                            $tempText = CreateMarkdownInNote
                            $Fields = $Fields + $tempText
                            $Fields[1] = $newDescription.Trim()
                            if($newDescription -like '*REFERENZEN:*'){
                                $newDescription = $newDescription.Substring("", $newDescription.indexof('REFERENZEN:'))
                                $Fields[1] = $newDescription.Trim()
                            }
                        }                                
                    }
                    elseif($originalDescription -like 'ESR *'){
                        $firstSpace = $false

                        $tempText = CreateMarkdownInNote
                        $Fields = $Fields + $tempText
                        for($j=0; $j -lt $originalDescription.length;$j++){ 
                            if ($originalDescription.Substring($j,1) -eq " " -and $firstSpace){
                                $newDescription = $originalDescription.Substring($j)
                                $Fields[1] = $newDescription.Trim()
                                break 
                            }
                            elseif($originalDescription.Substring($j,1) -eq " "){
                                $firstSpace = $true
                            }     
                        }
                    }   
                    elseif($originalDescription -like 'E-FINANCE *'){
                        $firstSpace = $false

                        $tempText = CreateMarkdownInNote
                        $Fields = $Fields + $tempText
                        for($j=0; $j -lt $originalDescription.length;$j++){ 
                            if ($originalDescription.Substring($j,1) -eq " " -and $firstSpace){
                                $newDescription = $originalDescription.Substring($j)
                                $Fields[1] = $newDescription.Trim()
                                break 
                            }
                            elseif($originalDescription.Substring($j,1) -eq " "){
                                $firstSpace = $true
                            }     
                        }
                    }
                    elseif($originalDescription -like 'KAUF/DIENSTLEISTUNG VOM *'){
                        $tempText = CreateMarkdownInNote
                        $Fields = $Fields + $tempText
                        $newDescription = $originalDescription.Substring($originalDescription.indexof('KARTEN NR. ')+ 'KARTEN NR. '.Length+8)
                        $Fields[1] = $newDescription.Trim()
                    }    
                    elseif($originalDescription -like 'ACHAT/SERVICE *'){
                        $tempText = CreateMarkdownInNote
                        $Fields = $Fields + $tempText
                        $newDescription = $originalDescription.Substring($originalDescription.indexof('CARTE N')+ 'CARTE N'.Length+10)
                        $Fields[1] = $newDescription.Trim()
                    }   
                    elseif($originalDescription -like 'BARGELDBEZUG VOM *'){
                        $tempText = CreateMarkdownInNote
                        $Fields = $Fields + $tempText
                        $newDescription = $originalDescription.Substring($originalDescription.indexof('KARTEN NR. ')+ 'KARTEN NR. '.Length+8)
                        $Fields[1] = $newDescription.Trim()
                    }
                    elseif($originalDescription -like "RETRAIT D'ESPECES*"){
                        $tempText = CreateMarkdownInNote
                        $Fields = $Fields + $tempText
                        $newDescription = $originalDescription.Substring($originalDescription.indexof('CARTE N')+ 'CARTE N'.Length+10)
                        $Fields[1] = $newDescription.Trim()
                    } 
                    elseif($originalDescription -like 'KAUF/ONLINE SHOPPING VOM *'){
                        $tempText = CreateMarkdownInNote
                        $Fields = $Fields + $tempText
                        $newDescription = $originalDescription.Substring($originalDescription.indexof('KARTEN NR. ')+ 'KARTEN NR. '.Length+8)
                        $Fields[1] = $newDescription.Trim()
                    } 
                    elseif($originalDescription -like 'ACHAT/SHOPPING EN LIGNE *'){
                        $tempText = CreateMarkdownInNote
                        $Fields = $Fields + $tempText
                        $newDescription = $originalDescription.Substring($originalDescription.indexof('CARTE N')+ 'CARTE N'.Length+10)
                        $Fields[1] = $newDescription.Trim()
                    } 
                    elseif($originalDescription -like '*YELLOWPAY / ONLINE-SHOPPING*'){
                        $tempText = CreateMarkdownInNote
                        $Fields = $Fields + $tempText
                        $newDescription = $originalDescription.Substring($originalDescription.indexof('YELLOWPAY / ONLINE-SHOPPING')+ 'YELLOWPAY / ONLINE-SHOPPING'.Length)
                        $Fields[1] = $newDescription.Trim()
                    } 
                    elseif($originalDescription -like 'AUFTRAG DEBIT DIRECT *'){
                        $tempText = CreateMarkdownInNote
                        $Fields = $Fields + $tempText
                        $newDescription = $originalDescription.Substring($originalDescription.indexof('KUNDENNUMMER ')+ 'KUNDENNUMMER '.Length+6)
                        $Fields[1] = $newDescription.Trim()
                    }   

                    
                }
                else{
                    $foundHeader = $true
                    $DescriptionColumnTitle = $originalDescription
                }
                
                if ($newline -eq "FirstLine") {  #happens only in first row. Add additinal comlumn in header.
                    $Fields = $Fields + "firely-iii NOTE"
                }

                $newline = ""
                for($j=0; $j -lt $Fields.count;$j++){
                    $newline = ($newline + ";" + $Fields[$j])
                }
                $newline = $newline.Substring(1) #as first charater is ";"

                #https://stackoverflow.com/questions/49120179/run-a-powershell-script-that-monitors-a-file-thats-lock-but-once-unlocked-runs
                While ($True) {                                                        
                    if ((checkLock $newFile_FullName) -eq $true) {
                        Write-Host "file locked"
                        continue
                    }
                    else {
                        [void]( add-content ($newFile_FullName ) $newline -Force -ErrorAction SilentlyContinue -ErrorVariable ProcessError) 
                        if($ProcessError){
                            Write-Host ("Not able to append: `n "+ $newline + "`nto`n " + $csvFiles[$i] + "-" + $NameExtension + ".csv" ) #https://devblogs.microsoft.com/scripting/powertip-new-lines-with-powershell/
                            $ProcessError ="" 
                            if (-Not (AmILastItemInLoop))  {
                                continue
                            }
                            else{
                                break
                            }
                        } 
                        break  
                    }
                    
                    start-sleep -seconds 0.25
                }


                 
            }
        }
    }

    #https://stackoverflow.com/questions/22349139/utf-8-output-from-powershell
    #write-output "hello" | out-file $newFile_FullName -encoding utf8
    #https://stackoverflow.com/questions/5596982/using-powershell-to-write-a-file-in-utf-8-without-the-bom
    $MyFile = Get-Content $newFile_FullName
    $MyFile | Out-File -Encoding "UTF8" $newFile_FullName
}
