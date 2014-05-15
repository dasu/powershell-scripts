# project data gathering v0.2 script
# INSTRUCTIONS:  change variables to whatever you want them to be.  
# create a text file called pn.txt, enter the project numbers into that text file, one for each new line.
# run script -- output: excel with information/locations to appropriate documentation for multiple project folders.

##Variables##

$shares =  "\\server\share" #enter project share here
$pns = gc "U:\pn.txt" #where to save the text file with project numbers in it.
$folders = "\Documentation\folder\Contractor Estimates", "\Documentation\folder\Applications for Payment", "\Documentation\folder\Budget"  #enter directories to search
$csv = "U:\report.csv" #where to save the CSV file 

##CODE, DON'T EDIT (unless you know what you're doing)##
$re = @()
$i = 0
$1 = Get-Date
foreach($share in $shares){
    foreach($pn in $pns){
        $m = $share + "\" + $pn
        foreach($folder in $folders){
            $centralf = $m + $folder
            if(Test-Path $centralf){
                $x = gci -recurse $centralf
                foreach($file in $x){
                    if($file){ 
                        $re += $file|Select-Object @{Name="Project Name";Expression={ (gc $m\**.****.**.txt)[1] -replace "Project Name: "}}, Name, DirectoryName, CreationTime, Extension
                    }
                    
                    else{  
                    $i ++
                    write-host "Empty folder #$i" #hurr, pretty much unnessecary but a good way to see if the script is even doing anything.
                    }
                }
            }
        }
    }
}
$re|Export-Csv -NoTypeInformation $csv

$2 = Get-Date
$tt = ($2 - $1).TotalMinutes


#Emailing capabilities, uncomment and fill out the appropriate information.
<#if (Test-Path $csv) 
	{
		$msg = new-object Net.Mail.MailMessage
		$att = new-object Net.Mail.Attachment($csv)
		$email = new-object Net.Mail.SmtpClient("smtp-server")
		$msg.From = ""
		$msg.To.Add("email@address.com")
		$msg.Subject = “files report”
		$msg.Body = "It took $tt minutes to run this report.”
		$msg.Attachments.Add($att)
        $email.Send($msg)
        $att.Dispose()
	} 
#>
