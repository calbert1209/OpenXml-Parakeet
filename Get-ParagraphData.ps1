$sampleDirPath = '/samples'
$outputFilePath = './output/data.csv';
$project = './OpenXmlLab'

$files = Get-ChildItem -Filter *.docx -Path $sampleDirPath

foreach ($file in $files) {
    dotnet run --project $project $file.FullName | Out-File -Append -FilePath $outputFilePath
    Write-Host $file.Fullname
}
