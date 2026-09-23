# PnP PowerShell Bulk SharePoint / OneDrive Uploader for Migrated NotebookLM (.docx)
# Reads sharepoint_spmt_manifest.csv and uploads each user's converted .docx sources to SharePoint Online / OneDrive.
param(
    [Parameter(Mandatory=$false)][string]$ManifestCsv = "./sharepoint_spmt_manifest.csv",
    [Parameter(Mandatory=$false)][string]$ClientId = $env:PNP_CLIENT_ID
)

Import-Module PnP.PowerShell -ErrorAction Stop
$rows = Import-Csv -Path $ManifestCsv

foreach ($row in $rows) {
    Write-Host "Uploading Notebook folder $($row.SourcePath) -> $($row.TargetWebUrl)/$($row.TargetDocumentLibrary)/$($row.TargetSubFolder)" -ForegroundColor Cyan
    Connect-PnPOnline -Url $row.TargetWebUrl -Interactive -ClientId $ClientId
    $localFiles = Get-ChildItem -Path $row.SourcePath -Recurse -File
    foreach ($file in $localFiles) {
        $relDir = $file.DirectoryName.Substring($row.SourcePath.Length).TrimStart('\', '/')
        $targetFolder = "$($row.TargetDocumentLibrary)/$($row.TargetSubFolder)/$relDir".TrimEnd('/')
        Add-PnPFile -Path $file.FullName -Folder $targetFolder -Values @{Title=$file.BaseName} | Out-Null
    }
}
Write-Host "All SharePoint / OneDrive NotebookLM folders uploaded successfully!" -ForegroundColor Green
