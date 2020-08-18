# Powershell displays Field Name="Title" Value="New to Unilever" Type="Text" as Field Name="Title" Type="Text" Value="New to Unilever", i.e. re-arranges Value and Text, so be cognizant

$excel = New-Object -Com Excel.Application
$wb = $excel.Workbooks.Open("C:\Users\hydro\Desktop\New Folder\langs.xlsx")
$sh = $wb.Sheets.Item(1)
$languages = 7
$translations = 105
$englishColumn = $sh.UsedRange.Columns(1).value2

for ($j = 2; $j -le $languages; $j++)
{	
	$file = Get-Content .\Settings-EN.xml -Raw
	for ($i = 2; $i -le $translations ; $i++)
	{
		$currentLanguage = $sh.UsedRange.Columns($j).value2
		$strMsg = "`r`n"
		#$strMsg
		#$englishColumn[$i, 1], $currentLanguage[$i,1]
		$file = ($file -replace $englishColumn[$i, 1], $currentLanguage[$i,1])
	}
	$currentLanguage[1,1]
	$fileName = $currentLanguage[1,1]
	Set-Content $file -Encoding UTF8 -Path .\$fileName.xml
}