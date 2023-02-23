#convert word documents into html and save imagenes with pandoc 

Write-Host "Do you want to convert this word documents Y o N?"
$fileBaseNames = (Get-ChildItem . -Filter *.docx).BaseName
Write-host $fileBaseNames
$confirmacion = Read-Host " "
$confirmacion.ToUpper() | Out-Null
if ($confirmacion -eq 'Y') {
try {
$fileBaseNames|
Foreach-Object {
	$filename = $_ -replace ('\s','')  
	md -Force ./$filename/ | Out-Null
	Move-Item -Path "$_.docx" -Destination "$filename.docx"
	pandoc --from docx --to html --extract-media=./$filename/ $filename'.docx' -o ./$filename/$filename'.txt'
	Write-Host "$filename.docx successfully convert" -ForegroundColor Green
}
Read-Host "Press any key exit"
}
Catch{
	 Throw "errormessage: $($_.Exception.Message)" 
}
}
