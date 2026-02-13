='D:\office_toolbox\office_toolbox_flutter\ODAFileConverter\ODAFileConverter.exe'
='F:\software_test\暖通书籍资料\精品图纸（一） - 副本\[江苏]苏宁电气广场暖通图纸\苏宁图纸（欢迎关注暖通公共账号：CHINA_HVAC，有更多资料下载）\新建文件夹（欢迎关注暖通公共账号：CHINA_HVAC，有更多资料下载）\暖通（欢迎关注暖通公共账号：CHINA_HVAC，有更多资料下载）'
=Join-Path  'output'
='01  目录.dwg'
=Join-Path  '01  目录.dxf'
if(Test-Path -LiteralPath ){Remove-Item -LiteralPath  -Force}
=Start-Process -FilePath  -ArgumentList @(,,'ACAD2010','DXF','0','0',) -WindowStyle Minimized -PassThru
Write-Output ('PID='+.Id)
=[System.Diagnostics.Stopwatch]::StartNew()
=False
while(.Elapsed.TotalSeconds -lt 180){
  if(Test-Path -LiteralPath ){
    =Get-Item -LiteralPath 
    if(.Length -gt 0){=True; break}
  }
  Start-Sleep -Milliseconds 500
}
if(){Write-Output 'DXF_READY'} else {Write-Output 'NO_DXF_180S'}
if(-not .HasExited){Stop-Process -Id .Id -Force -ErrorAction SilentlyContinue; Write-Output 'KILLED'} else {Write-Output ('EXIT='+.ExitCode)}
if(Test-Path -LiteralPath ){Get-Item -LiteralPath  | Select-Object FullName,Length,LastWriteTime} else {Write-Output 'NO_DXF_END'}
