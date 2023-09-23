#Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted
$package = Join-Path $env:USERPROFILE ".nuget" | Join-Path -ChildPath "packages" | Join-Path -ChildPath "excelmvc.net"
Remove-Item -Force -Path $package -Recurse