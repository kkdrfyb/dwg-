$exe='D:\office_toolbox\office_toolbox_flutter\ODAFileConverter\ODAFileConverter.exe'
$input='D:\oda_short\01.dwg'
$out='D:\oda_short\out'
$target=Join-Path $out '01.dxf'
if(Test-Path -LiteralPath $target){Remove-Item -LiteralPath $target -Force}
$p=Start-Process -FilePath $exe -ArgumentList @($input,$out,'DWG','DXF2010','-quiet') -PassThru
Write-Output ('PID='+$p.Id)
$done=Wait-Process -Id $p.Id -Timeout 45 -ErrorAction SilentlyContinue
if($null -eq $done){Write-Output 'RUNNING_AFTER_45'} else {Write-Output ('EXIT='+$p.ExitCode)}
if(Test-Path -LiteralPath $target){Get-Item -LiteralPath $target | Select-Object FullName,Length,LastWriteTime} else {Write-Output 'NO_DXF'}
if(-not $p.HasExited){Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue; Write-Output 'KILLED'}
