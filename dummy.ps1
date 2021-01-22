$Document = "C:\Users\sepp\Downloads\Neuer Ordner\Dummy.docx"



$FindText = "<Nr>"



$ReplaceText = Read-Host "Laufende Nr"


$Datum = Get-Date



$ReplaceAll = 2
$FindContinue = 1
$MatchCase = $false
$MatchWholeWord = $True
$MatchWildcards = $false
$MatchSoundsLike = $false
$MatchAllWordForms = $false
$Forward = $True
$Wrap = $FindContinue
$Format = $false

$Word = New-Object -Comobject Word.Application
$Word.Visible = $True

$OpenDoc = $Word.Documents.Open($Document)

$Selection = $Word.Selection







$Selection.Find.Execute(
$FindText,
$MatchCase,
$MatchWholeWord,
$MatchWildcards,
$MatchSoundsLike,
$MatchAllWordForms,
$Forward,
$Wrap,
$Format,
$ReplaceText,
$ReplaceAll
) | Out-Null


if ($Selection.Find.Found) { 
	Write-Host("Gefunden")
} else { 
	Write-Host("Nicht gefunden")
} 




$FindText = "<Tag>"



$ReplaceText = $Datum.Day 


$Selection.Find.Execute(
$FindText,
$MatchCase,
$MatchWholeWord,
$MatchWildcards,
$MatchSoundsLike,
$MatchAllWordForms,
$Forward,
$Wrap,
$Format,
$ReplaceText,
$ReplaceAll
) | Out-Null


if ($Selection.Find.Found) { 
	Write-Host("Gefunden")
} else { 
	Write-Host("Nicht gefunden")
} 






$OpenDoc.saveas("C:\Users\sepp\Downloads\Neuer Ordner\Dummy1.docx")
$OpenDoc.Close()
$Word.Quit()