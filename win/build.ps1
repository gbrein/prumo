# Monta o Prumo no Windows: dist\PrumoPanel.dll (o painel) e dist\Prumo.ppam (VBA + ribbon, o mesmo para Mac e Windows).
# Precisa de PowerPoint e Python 3. O csc.exe já vem com o Windows (.NET Framework 4).
# Para importar o VBA, liga "Confiar no acesso ao modelo de objeto do projeto do VBA" só durante a montagem e depois devolve.
#   powershell -ExecutionPolicy Bypass -File win\build.ps1 [-SkipTests]
param([switch]$SkipTests)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$dist = Join-Path $root 'dist'
New-Item -ItemType Directory -Force $dist | Out-Null

python "$root\build\gen_panel.py"
if ($LASTEXITCODE) { throw 'gen_panel.py falhou' }

& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" -nologo -target:library -optimize "-out:$dist\PrumoPanel.dll" `
    -r:System.Windows.Forms.dll -r:System.Drawing.dll -r:Microsoft.CSharp.dll -r:System.Core.dll `
    "-resource:$PSScriptRoot\panel.tsv,panel.tsv" "$PSScriptRoot\PrumoPanel.cs"
if ($LASTEXITCODE) { throw 'a compilação do painel falhou' }
Write-Host "painel: $dist\PrumoPanel.dll"

# O PowerPoint grava as opções de segurança ao sair: a montagem abre e fecha a própria instância.
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Feche o PowerPoint antes de montar.' }
$sec = 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security'
$old = (Get-ItemProperty $sec -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty $sec -Name AccessVBOM -Value 1 -Type DWord
$pptm = "$dist\Prumo.pptm"
try {
    $app = New-Object -ComObject PowerPoint.Application
    $pres = $app.Presentations.Add(0)
    foreach ($m in 'modLayout', 'modText', 'modSlides', 'modRibbon', 'modPanel', 'modUi', 'modTests') {
        $null = $pres.VBProject.VBComponents.Import("$root\src\$m.bas")
    }
    if (Test-Path $pptm) { Remove-Item $pptm }
    $pres.SaveAs($pptm, 25)  # ppSaveAsOpenXMLPresentationMacroEnabled: o build.py converte em .ppam, igual no Mac
    if (-not $SkipTests) {
        $fails = [System.__ComObject].InvokeMember('Run', 'InvokeMethod', $null, $app, @('Prumo.pptm!Prumo_SelfTestRun'))  # Run tem ParamArray: so por IDispatch
        if ($fails) { throw "autoteste falhou: $fails" }
        Write-Host 'autoteste: ok'
    }
}
finally {
    if ($app) { $app.Quit() }
    $pres = $null; $app = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()  # solta o PowerPoint para ele fechar
    Wait-Process POWERPNT -Timeout 60 -ErrorAction SilentlyContinue  # só então devolve a opção, senão ele regrava ao sair
    if ($null -eq $old) { Remove-ItemProperty $sec -Name AccessVBOM } else { Set-ItemProperty $sec -Name AccessVBOM -Value $old }
}

python "$root\build\build.py" $pptm
if ($LASTEXITCODE) { throw 'build.py falhou' }
