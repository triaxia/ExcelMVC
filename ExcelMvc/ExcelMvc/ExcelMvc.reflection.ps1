param ([string]$TargetPath)
Get-ChildItem -Name -Path "$TargetPath" -Include "*.dll" -Exclude System*,Microsoft*,ExcelMvc* | Add-Content (Join-Path $TargetPath  "ExcelMvc.reflection.txt")