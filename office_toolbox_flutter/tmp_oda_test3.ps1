$exe='D:\office_toolbox\office_toolbox_flutter\ODAFileConverter\ODAFileConverter.exe'
$inDir=(Get-Item -LiteralPath 'F:\software_test\暖通书籍资料\精品图纸（一） - 副本\[江苏]苏宁电气广场暖通图纸\苏宁图纸（欢迎关注暖通公共账号：CHINA_HVAC，有更多资料下载）\新建文件夹（欢迎关注暖通公共账号：CHINA_HVAC，有更多资料下载）\暖通（欢迎关注暖通公共账号：CHINA_HVAC，有更多资料下载）').FullName
$out=Join-Path $inDir 'output'
if(-not (Test-Path -LiteralPath $out)){New-Item -ItemType Directory -Path $out | Out-Null}
$target=Join-Path $out '01  目录.dxf'
if(Test-Path -LiteralPath $target){Remove-Item -LiteralPath $target -Force}
$p=Start-Process -FilePath $exe -ArgumentList @($inDir,$out,'ACAD2010','DXF','0','0','01  目录.dwg') -PassThru
Write-Output ('PID='+$p.Id)
$done=Wait-Process -Id $p.Id -Timeout 120 -ErrorAction SilentlyContinue
if($null -eq $done){Write-Output 'RUNNING_AFTER_120'} else {Write-Output ('EXIT='+$p.ExitCode)}
if(Test-Path -LiteralPath $target){Get-Item -LiteralPath $target | Select-Object FullName,Length,LastWriteTime} else {Write-Output 'NO_DXF'}
if(-not $p.HasExited){Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue; Write-Output 'KILLED'}
