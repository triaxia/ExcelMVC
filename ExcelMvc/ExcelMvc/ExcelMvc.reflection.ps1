param ([string]$TargetPath, [string]$TargetName)
Get-ChildItem -Name -Path "$TargetPath" -Include "*.dll" -Exclude System*,Microsoft*,ExcelMvc* | Add-Content (Join-Path $TargetPath  $TargetName)