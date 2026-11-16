# Arquivo sem BOM e mensagens sem acento: assim roda igual por irm (download) e por -File (lido como ANSI).
# Instala, atualiza ou remove o Prumo no PowerPoint para Windows (sem admin: tudo em HKCU e %APPDATA%\Prumo).
#   & ([scriptblock]::Create((irm https://raw.githubusercontent.com/gbrein/prumo/main/win/install.ps1)))
#   (no fim, -Uninstall desinstala; -From <pasta> instala dos arquivos Prumo.ppam e PrumoPanel.dll já baixados)
# Ou com o arquivo baixado:
#   powershell -ExecutionPolicy Bypass -File install.ps1 -From <pasta>
#   powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall
param([string]$From, [switch]$Uninstall)
$ErrorActionPreference = 'Stop'
$url = 'https://github.com/gbrein/prumo/releases/latest/download'
$dest = Join-Path $env:APPDATA 'Prumo'
$ppamKey = 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\AddIns\Prumo'
$addinKey = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\Prumo.Panel'
# As duas classes COM do painel: ProgId, classe .NET e Guid (os de win/PrumoPanel.cs).
$classes = @(
    @('Prumo.Panel', 'Prumo.Connect', '{6B0B3E4C-6C1E-4C0B-9A57-0E5C2A1D7F01}'),
    @('Prumo.PanelControl', 'Prumo.PanelControl', '{6B0B3E4C-6C1E-4C0B-9A57-0E5C2A1D7F02}'))
$clsidRoots = 'HKCU:\Software\Classes\CLSID', 'HKCU:\Software\Classes\Wow6432Node\CLSID'  # Office 64 e 32 bits

if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    throw 'Feche o PowerPoint antes (os arquivos do Prumo ficam presos enquanto ele esta aberto).'
}

$old = @($ppamKey, $addinKey)
foreach ($c in $classes) { $old += "HKCU:\Software\Classes\$($c[0])"; foreach ($r in $clsidRoots) { $old += "$r\$($c[2])" } }
foreach ($k in $old) {
    if (Test-Path $k) { Remove-Item $k -Recurse -Force }
}
if ($Uninstall) {
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    Write-Host 'Prumo removido.'
    return
}

New-Item -ItemType Directory -Force $dest | Out-Null
foreach ($f in 'Prumo.ppam', 'PrumoPanel.dll') {
    $to = Join-Path $dest $f
    if ($From) { Copy-Item (Join-Path $From $f) $to -Force }
    else { Invoke-WebRequest "$url/$f" -OutFile $to -UseBasicParsing }
    Unblock-File $to  # baixado da internet: o Office bloquearia as macros e o .NET, a DLL
}
$dll = Join-Path $dest 'PrumoPanel.dll'

# Classes COM do painel, como o regasm /codebase faria, mas só para este usuário.
foreach ($c in $classes) {
    $progId, $class, $clsid = $c
    New-Item "HKCU:\Software\Classes\$progId\CLSID" -Force -Value $clsid | Out-Null
    foreach ($r in $clsidRoots) {
        $k = "$r\$clsid"
        New-Item $k -Force -Value $class | Out-Null
        New-Item "$k\ProgId" -Force -Value $progId | Out-Null
        New-Item "$k\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" -Force | Out-Null
        foreach ($s in "$k\InprocServer32", "$k\InprocServer32\1.0.0.0") {
            New-Item $s -Force -Value 'mscoree.dll' | Out-Null
            if ($s -like '*32') { New-ItemProperty $s -Name ThreadingModel -Value Both | Out-Null }
            New-ItemProperty $s -Name Class -Value $class | Out-Null
            New-ItemProperty $s -Name Assembly -Value 'PrumoPanel, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null' | Out-Null
            New-ItemProperty $s -Name RuntimeVersion -Value 'v4.0.30319' | Out-Null
            New-ItemProperty $s -Name CodeBase -Value ([Uri]$dll).AbsoluteUri | Out-Null
        }
    }
}

# O painel (add-in COM, carrega com o PowerPoint) e o suplemento (.ppam, carrega sozinho ao abrir).
New-Item $addinKey -Force | Out-Null
New-ItemProperty $addinKey -Name FriendlyName -Value 'Prumo Painel' | Out-Null
New-ItemProperty $addinKey -Name Description -Value 'Painel lateral do Prumo' | Out-Null
New-ItemProperty $addinKey -Name LoadBehavior -Value 3 -PropertyType DWord | Out-Null
New-Item $ppamKey -Force | Out-Null
New-ItemProperty $ppamKey -Name Path -Value (Join-Path $dest 'Prumo.ppam') | Out-Null
New-ItemProperty $ppamKey -Name AutoLoad -Value 1 -PropertyType DWord | Out-Null

Write-Host "Prumo instalado em $dest."
Write-Host 'Abra o PowerPoint: a aba Prumo aparece no ribbon e o painel abre a direita. Se perguntar sobre macros, clique em Habilitar.'
