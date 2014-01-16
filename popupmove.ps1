$LOCATION = "\\server\share\path\to\folder\"
$SCRIPTSTUFF = "C:\SCRIPTS\MOTD"
$status = 'null'
if ((get-date -uformat %a) -eq 'sat') 
    {
    remove-item $LOCATION\JPEGs\*.*
    $month = (get-date).adddays(2).tostring('MM MMM')
    $day = (get-date).adddays(2).tostring('dd')
    copy-item $LOCATION\2014\$month\$day\* -destination $LOCATION\JPEGs\
        if(compare-object -ReferenceObject @(gci "$LOCATION\2014\$month\$day\") -DifferenceObject @(gci "$LOCATION\JPEGs\") -IncludeEqual -Property Name, length)
    {
    $status = 'GOOD'
    compare-object -ReferenceObject @(gci "$LOCATION\2014\$month\$day\") -IncludeEqual -DifferenceObject @(gci "$LOCATION\JPEGs\") -Property Name, length, lastwritetime|Format-List > $SCRIPTSTUFF\difftest.txt
    }
        else
    {
    $status = 'PROBLEM'   
    compare-object -ReferenceObject @(gci "$LOCATION\2014\$month\$day\") -IncludeEqual -DifferenceObject @(gci "$LOCATION\JPEGs\") -Property Name, length, lastwritetime|Format-List > $SCRIPTSTUFF\difftest.txt
    }
    }
else 
{
    "The script was runned outside of saturday for some reason.  No changes have been made." > $SCRIPTSTUFF\difftest.txt
    $status = "???"
     }

C:\SCRIPTS\MOTD\blat\blat.exe $SCRIPTSTUFF\difftest.txt -to "email@address.com,email2@gmail.com" -f from@me.com -s "Popup move status: $status" -server smtp.server
