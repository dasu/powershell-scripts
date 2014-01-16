#unlocks password protected word documents so they can be edited freely, just in case one needs to be edited. 
#Created 6-19-2013 
#mostly created this to save time and learn more about powershell.
if (!(Test-Path -path C:\TEMP\UNLOCK\)) {New-Item C:\TEMP\UNLOCK -Type Directory}
Remove-Item -recurse C:\TEMP\UNLOCK\*

function Select-FileDialog
{
	param([string]$Title,[string]$Directory,[string]$Filter="All Files (*.*)|*.*")
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objForm = New-Object System.Windows.Forms.OpenFileDialog
	$objForm.InitialDirectory = $Directory
	$objForm.Filter = $Filter
	$objForm.Title = $Title
	$Show = $objForm.ShowDialog()
	If ($Show -eq "OK")
	{
		Return $objForm.FileName
	}
	Else
	{
		Write-Error "Operation cancelled by user."
	}
}

$FILE = Select-FileDialog -Title "Select a locked word document" -Filter "Word Documents|*.docx"
$DEST = Copy-Item $FILE -destination C:\TEMP\ -PassThru
$FNAME = $DEST.Name
$NZIP = $FNAME -replace ".docx",".zip"
$ZIP = Rename-Item $DEST -NewName $NZIP -PassThru
$shell_app=new-object -com shell.application
$zip_file = $shell_app.namespace(($ZIP).FullName)
$destination = $shell_app.namespace("C:\TEMP\UNLOCK")
$destination.Copyhere($zip_file.items())
Remove-Item $ZIP.FullName
set-content $ZIP.FullName ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))

$xml = New-Object xml
$xml.Load("C:\TEMP\UNLOCK\word\settings.xml")
$XML.settings.documentProtection|foreach {$_.Enforcement = '0'}
$xml.save("C:\TEMP\UNLOCK\word\settings.xml")

$zip_file.CopyHere("C:\TEMP\UNLOCK\Customxml")
sleep 1
$zip_file.CopyHere("C:\TEMP\UNLOCK\docProps")
sleep 1
$zip_file.CopyHere("C:\TEMP\UNLOCK\word")
sleep 2
$zip_file.CopyHere("C:\TEMP\UNLOCK\_rels")
sleep 1
$zip_file.CopyHere("C:\TEMP\UNLOCK\[Content_Types].XML")
sleep 1
$UDOCX = $ZIP.Name -replace ".zip","_UNLOCKED.docx"
$UFILE = Rename-Item $ZIP.FullName -newname $UDOCX -PassThru
Move-Item $UFILE.FullName -Destination (split-path $FILE)
Remove-Item -recurse C:\TEMP\UNLOCK\*

echo $FILE
echo $DEST
echo $FNAME
echo $NZIP
echo $ZIP.FullName
