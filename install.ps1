<#
  install.ps1 -- copy this repo's skills into your personal Claude Code skills
  folder, for people who prefer not to use the plugin marketplace.

  Usage (from the repo folder, in Windows PowerShell):
      .\install.ps1                      # every plugin
      .\install.ps1 -Plugin excel-dev    # one plugin
      .\install.ps1 -Dest C:\temp\skills # somewhere else (for testing)

  A skill folder that already exists at the destination is first copied to
  <Dest>\.backups\<skill>-<yyyy-MM-dd-HHmmss> and then replaced, so nothing of
  yours is lost. Commands go to the sibling 'commands' folder the same way.
#>
param(
    [string]$Plugin = "*",
    [string]$Dest = (Join-Path $HOME ".claude\skills")
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$stamp = Get-Date -Format "yyyy-MM-dd-HHmmss"
$backups = Join-Path $Dest ".backups"
$cmdDest = Join-Path (Split-Path -Parent $Dest) "commands"
New-Item -ItemType Directory -Force -Path $Dest | Out-Null

$count = 0
foreach ($p in Get-ChildItem (Join-Path $root "plugins") -Directory | Where-Object Name -like $Plugin) {
    $skillsDir = Join-Path $p.FullName "skills"
    if (Test-Path $skillsDir) {
        foreach ($s in Get-ChildItem $skillsDir -Directory) {
            $target = Join-Path $Dest $s.Name
            if (Test-Path $target) {
                New-Item -ItemType Directory -Force -Path $backups | Out-Null
                Copy-Item $target (Join-Path $backups "$($s.Name)-$stamp") -Recurse
                Remove-Item $target -Recurse -Force
                Write-Host "  backed up existing $($s.Name)"
            }
            Copy-Item $s.FullName $target -Recurse
            Write-Host "installed skill   $($s.Name)  ($($p.Name))"
            $count++
        }
    }
    $cmds = Join-Path $p.FullName "commands"
    if (Test-Path $cmds) {
        New-Item -ItemType Directory -Force -Path $cmdDest | Out-Null
        foreach ($c in Get-ChildItem $cmds -Filter *.md) {
            $target = Join-Path $cmdDest $c.Name
            if (Test-Path $target) {
                New-Item -ItemType Directory -Force -Path $backups | Out-Null
                Copy-Item $target (Join-Path $backups "$($c.BaseName)-command-$stamp.md")
            }
            Copy-Item $c.FullName $target -Force
            Write-Host "installed command /$($c.BaseName)  ($($p.Name))"
            $count++
        }
    }
}
Write-Host "`n$count item(s) installed to $Dest. Restart Claude Code to load them."
