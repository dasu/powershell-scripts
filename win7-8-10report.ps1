#makes a general report on win10/win7/win8 usage for a specific region, for all offices, or for all regions.
#outputs a % chart to console, copy to excel and make it fancy if you want.
#filters out machines from 'banned' OUs, disabled computers in AD, and any PCs that have not been logged in to by the cutoff (14 days)
###
#variables, read the comments :|
$allregions = $true #true: checks all computers in AD (slowish, slight innaccuracies) , false: checks only the region you set below (faster and more accurate)
$CombineToRegion = $true #if above is true, and this is true, it'll combine each office per region and give a region %...Possibly more accurate than the other ways.
#Both set to false would give you regional info only, which requires filling out the variables below.
$region = "NorthEast" #Region OU as it appears in AD
$offices = "NJ","NY"#,"  #fill in your individiual offices in your region, (the first two letters of every computer name basically)...
                                               #I can probably code this differently to avoid having to enter these, but whatever i've hit max laziness on this script.
$banned = "Domain Controllers","Protected"#,"morelocations"    #ban list for combine to region, you probably don't have to touch this.  Uses OU name.
$reportandtrackinglocation = "C:\SCRIPTS\Win10Report\" #Location for csv report, and .json file used to track % change since last run.  If you change this, don't forget the trailing \
$cutoff = (get-date).AddDays(-14) #cutoffdate to include computers which have logged in in the past 14 days.  
###
if (-not (Test-Path $reportandtrackinglocation)){
    New-Item -Path $reportandtrackinglocation -ItemType Directory -Force|Out-Null
}
if($CombineToRegion -and -not $allregions){
    write-host -ForegroundColor Cyan "`$allregions needs to be `$true if you want to use `$CombineToRegion..." 
    return
}
write-host "`nYes I'm working, please wait like 2 minutes or somethin`'...`n"
$comp = @()
if($allregions){
    $all = get-adcomputer -filter {OperatingSystem -like "Windows * Enterprise" -and lastlogontimestamp -ge $cutoff -and enabled -eq $true} -Properties Name, operatingSystem
    $all|foreach{
        $loc = $_
        $loc|Add-Member -MemberType NoteProperty -Name Office -Value $loc.Name.substring(0,2) -Force
        if($allregions -and $CombineToRegion){
            $loc|Add-Member -MemberType NoteProperty -Name Region -Value (($loc.DistinguishedName -split "OU=")[-1] -split ",")[0] -Force
        }
        $comp += $loc
    }
    $group = ($comp|group-object {$_.name.substring(0,2)})|Where-Object{$_.Count -ge 10}
}
else{
    $ittestou = get-adcomputer -filter {OperatingSystem -like "Windows * Enterprise" -and lastlogontimestamp -ge $cutoff -and enabled -eq $true} -SearchBase "OU=Computers,OU=ittest,DC=company,DC=ad" -Properties name, operatingSystem
    $neregion = get-adcomputer -filter {OperatingSystem -like "Windows * Enterprise" -and lastlogontimestamp -ge $cutoff -and enabled -eq $true} -SearchBase "OU=$($region),DC=company,DC=ad" -Properties name, operatingSystem
    $all = $ittestou+$neregion
    $all|foreach{
        $loc = $_
        $offices|foreach{
            if($loc.name.StartsWith($_)){
                $loc|Add-Member -MemberType NoteProperty -Name Office -Value $loc.Name.substring(0,2) -Force
                $comp += $loc
            }
        }
    }
    $group = $comp|Group-Object Office
}
if($allregions -and $CombineToRegion){
    $win10percentages = new-object System.Collections.ArrayList
    $group = $comp|where-object{($_.region -notin $banned)}|where-object{($_.distinguishedname -notmatch "OU=Servers")}|Group-Object Region
    $group = $group|where-object{$_.name -ne "LaCrosse"}
    if(Test-Path "$($reportandtrackinglocation)allregions10percent.json"){
        $old = gc -raw "$($reportandtrackinglocation)allregions10percent.json"|ConvertFrom-Json
    }
    if(Test-Path "$($reportandtrackinglocation)allregionstotalpercent.json"){
        $oldtotal = gc -raw "$($reportandtrackinglocation)allregionstotalpercent.json"|ConvertFrom-Json
    }
    foreach($reg in $group){
        $percentage = new-object psobject
        $percentage|Add-Member -MemberType NoteProperty -Name Region -Value $reg.Name
        $percentage|Add-Member -MemberType NoteProperty -Name Win10 -Value ($reg.group|where-object{$_.OperatingSystem -eq "Windows 10 Enterprise"}|measure-object).count 
        $percentage|Add-Member -MemberType NoteProperty -Name Win7 -value ($reg.group|where-object {$_.OperatingSystem -eq "Windows 7 Enterprise"}|measure-object).count
        $percentage|add-member -MemberType NoteProperty -Name Win8 -value ($reg.group|where-object {$_.OperatingSystem -notmatch "Windows (7|10) Enterprise"}|measure-object).count
        $percentage|Add-Member -Membertype NoteProperty -Name Total -Value $reg.group.count 
        $percentage|Add-Member -MemberType NoteProperty -Name "% Win10" -Value ("{0:P2}" -f ($percentage.Win10 / $percentage.Total))
        if(Test-Path "$($reportandtrackinglocation)allregions10percent.json"){
            #$old[$old.Region.IndexOf($reg.name)]
            $percentage|Add-Member -MemberType NoteProperty -Name "Change since last run" -Value ([string]([decimal]($percentage.'% Win10' -replace " %","") - [decimal]($old[$old.Region.IndexOf($reg.name)].'% Win10' -replace " %",""))+" %")
        }
        [void]$win10percentages.Add($percentage)
    }
    $win10percentages|Sort-Object -Descending -Property {[decimal]($_.'% win10'.replace(" %",""))}|ft
    $win10percentages|foreach{select-object -InputObject $_ -Property Region, '% Win10'}|ConvertTo-Json|out-file -FilePath "$($reportandtrackinglocation)allregions10percent.json"
    write-host "`n`nGRAND TOTAL:"
    $grandtotal = new-object psobject
    $grandtotal|Add-Member -MemberType NoteProperty -Name Region -Value "TOTAL"
    $grandtotal|Add-Member -MemberType NoteProperty -Name Win10 -Value (($win10percentages|Measure-Object -Property Win10 -sum).Sum)
    $grandtotal|Add-Member -MemberType NoteProperty -Name Win7 -Value (($win10percentages|Measure-Object -Property Win7 -sum).Sum)
    $grandtotal|Add-Member -MemberType NoteProperty -Name Win8 -value (($win10percentages|Measure-Object -Property Win8 -sum).Sum)
    $grandtotal|Add-Member -Membertype NoteProperty -Name Total -Value (($win10percentages|Measure-Object -Property Total -sum).Sum)
    $grandtotal|Add-Member -Membertype NoteProperty -Name "% Win10" -Value ("{0:P2}" -f ((($win10percentages|Measure-Object -Property Win10 -sum).Sum)/(($win10percentages|Measure-Object -Property Total -sum).Sum)))
    if(Test-Path "$($reportandtrackinglocation)allregionstotalpercent.json"){
        $grandtotal|Add-Member -MemberType NoteProperty -Name "Change since last run" -Value ([string]([decimal]($grandtotal.'% Win10' -replace " %","") - [decimal]($oldtotal[$oldtotal.Region.IndexOf("TOTAL")].'% Win10' -replace " %",""))+" %")
    }
    $grandtotal|ft
    $comp|Select-Object Office, Region, Name, OperatingSystem, Enabled, DistinguishedName|Export-Csv -NoTypeInformation "$($reportandtrackinglocation)Win10percentage.csv"
    $grandtotal|foreach{select-object -InputObject $_ -Property Region, '% Win10'}|ConvertTo-Json|out-file -FilePath "$($reportandtrackinglocation)allregionstotalpercent.json"
}
else{
    $win10percentages = new-object System.Collections.ArrayList
    if($allregions -eq $false){
        if(Test-Path "$($reportandtrackinglocation)localoffices10percent.json"){
            $old = gc -raw "$($reportandtrackinglocation)localoffices10percent.json"|ConvertFrom-Json
        }
        if(test-path "$($reportandtrackinglocation)localofficestotalpercent.json"){
            $oldtotal = gc -raw "$($reportandtrackinglocation)localofficestotalpercent.json"|ConvertFrom-Json
        }
    }
    else{
        if(Test-Path "$($reportandtrackinglocation)alloffices10percent.json"){
            $old = gc -raw "$($reportandtrackinglocation)alloffices10percent.json"|ConvertFrom-Json
        }
        if(Test-Path "$($reportandtrackinglocation)allofficestotalpercent.json"){
            $oldtotal = gc -raw "$($reportandtrackinglocation)allofficestotalpercent.json"|ConvertFrom-Json
        }
    }
    foreach($off in $group){
        $percentage = new-object psobject
        $percentage|Add-Member -MemberType NoteProperty -Name Office -Value $off.Name.ToUpper()
        $percentage|Add-Member -MemberType NoteProperty -Name Win10 -Value ($off.group|where-object{$_.OperatingSystem -eq "Windows 10 Enterprise"}|measure-object).count  
        $percentage|Add-Member -MemberType NoteProperty -Name Win7 -value ($off.group|where-object {$_.OperatingSystem -eq "Windows 7 Enterprise"}|measure-object).count
        $percentage|add-member -MemberType NoteProperty -Name Win8 -value ($off.group|where-object {$_.OperatingSystem -notmatch "Windows (7|10) Enterprise"}|measure-object).count
        $percentage|Add-Member -Membertype NoteProperty -Name Total -Value $off.group.count
        $percentage|Add-Member -MemberType NoteProperty -Name "% Win10" -Value ("{0:P2}" -f ($percentage.Win10 / $percentage.Total))
        if($allregions -eq $false){
            if(Test-Path "$($reportandtrackinglocation)localoffices10percent.json"){
                $percentage|Add-Member -MemberType NoteProperty -Name "Change since last run" -Value ([string]([decimal]($percentage.'% Win10' -replace " %","") - [decimal]($old[$old.Office.IndexOf($off.name.ToUpper())].'% Win10' -replace " %",""))+" %")
            }
        }
        else{
            if(Test-Path "$($reportandtrackinglocation)alloffices10percent.json"){
                $percentage|Add-Member -MemberType NoteProperty -Name "Change since last run" -Value ([string]([decimal]($percentage.'% Win10' -replace " %","") - [decimal]($old[$old.Office.IndexOf($off.name.ToUpper())].'% Win10' -replace " %",""))+" %")
            }
        }
        [void]$win10percentages.Add($percentage)
    }
    $win10percentages|Sort-Object -Descending -Property {[decimal]($_.'% win10'.replace(" %",""))}|ft
    write-host "`n`nGRAND TOTAL:"
    $grandtotal = new-object psobject
    $grandtotal|Add-Member -MemberType NoteProperty -Name Office -Value "TOTAL" 
    $grandtotal|Add-Member -MemberType NoteProperty -Name Win10 -Value (($win10percentages|Measure-Object -Property Win10 -sum).Sum)
    $grandtotal|Add-Member -MemberType NoteProperty -Name Win7 -Value (($win10percentages|Measure-Object -Property Win7 -sum).Sum)
    $grandtotal|Add-Member -MemberType NoteProperty -Name Win8 -value (($win10percentages|Measure-Object -Property Win8 -sum).Sum)
    $grandtotal|Add-Member -Membertype NoteProperty -Name Total -Value (($win10percentages|Measure-Object -Property Total -sum).Sum)
    $grandtotal|Add-Member -Membertype NoteProperty -Name "% Win10" -Value ("{0:P2}" -f ((($win10percentages|Measure-Object -Property Win10 -sum).Sum)/(($win10percentages|Measure-Object -Property Total -sum).Sum)))
    if($allregions -eq $false){
        if(Test-Path "$($reportandtrackinglocation)localofficestotalpercent.json"){
            $grandtotal|Add-Member -MemberType NoteProperty -Name "Change since last run" -Value ([string]([decimal]($grandtotal.'% Win10' -replace " %","") - [decimal]($oldtotal[$oldtotal.Office.IndexOf("TOTAL")].'% Win10' -replace " %",""))+" %")
        }
    }
    else{
        if(Test-Path "$($reportandtrackinglocation)allofficestotalpercent.json"){
            $grandtotal|Add-Member -MemberType NoteProperty -Name "Change since last run" -Value ([string]([decimal]($grandtotal.'% Win10' -replace " %","") - [decimal]($oldtotal[$oldtotal.Office.IndexOf("TOTAL")].'% Win10' -replace " %",""))+" %")
        }
    }
    $grandtotal|ft
    $comp|Select-Object Office, Name, OperatingSystem, Enabled, DistinguishedName|Export-Csv -NoTypeInformation "$($reportandtrackinglocation)Win10percentage.csv"
    if($allregions -eq $false){
        $win10percentages|foreach{select-object -InputObject $_ -Property Office, '% Win10'}|ConvertTo-Json|out-file -FilePath "$($reportandtrackinglocation)localoffices10percent.json"
        $grandtotal|foreach{select-object -InputObject $_ -Property Office, '% Win10'}|ConvertTo-Json|out-file -FilePath "$($reportandtrackinglocation)localofficestotalpercent.json"

    }
    else{
        $win10percentages|foreach{select-object -InputObject $_ -Property Office, '% Win10'}|ConvertTo-Json|out-file -FilePath "$($reportandtrackinglocation)alloffices10percent.json"
        $grandtotal|foreach{select-object -InputObject $_ -Property Office, '% Win10'}|ConvertTo-Json|out-file -FilePath "$($reportandtrackinglocation)allofficestotalpercent.json"
    }
}
write-host "You can copy paste the above into excel and make it a prettier table if you want.  A full CSV report was also saved to $($reportandtrackinglocation)win10percentage.csv 👌"


Send-MailMessage -BodyAsHtml -to "recipients" -bcc "me@email.com"  -from "Windows10Report@email.com" -Subject "Windows 10 Report $(get-date -Format 'yyyy-MM-dd')" -SmtpServer email.server -body $((($win10percentages|Sort-Object -Descending -Property {[decimal]($_.'% win10'.replace(" %",""))}) + $grandtotal|ConvertTo-Html -head "<style>td{text-align: center; vertical-align: middle;} table{border-collapse: collapse;} table, th, td{border: 1px solid black; padding:4px} th {font-weight: bold; text-align: center; vertical-align: middle;background-color:#cfd4d8;} tr:last-child{font-weight: bold;}</style>" -PostContent "<br>The report is now filtering out computers that have not logged in the past 14 days and disabled in AD.<br>PS: See attached for All offices report")|out-string) -Attachments "$($reportandtrackinglocation)Alloffices.csv"
