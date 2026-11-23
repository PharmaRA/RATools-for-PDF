param([Parameter(Mandatory = $true)][string]$LiteralPath)
$ErrorActionPreference = 'Stop'
$siteRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$target = [IO.Path]::GetFullPath($LiteralPath)
$allowed = @('dist', '.astro', 'node_modules', '.npm-cache') | ForEach-Object { Join-Path $siteRoot $_ }
$inScope = $false
foreach ($root in $allowed) {
    if ($target.Equals($root, [StringComparison]::OrdinalIgnoreCase) -or
        $target.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        $inScope = $true
    }
}
if (-not $inScope) { throw 'Refusing cleanup outside website generated directories.' }
if (-not (Test-Path -LiteralPath $target)) { exit 0 }
$item = Get-Item -LiteralPath $target -Force
if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Refusing to recycle a reparse point.' }
Add-Type -AssemblyName Microsoft.VisualBasic
if ($item.PSIsContainer) {
    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($target, 'OnlyErrorDialogs', 'SendToRecycleBin', 'ThrowException')
} else {
    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($target, 'OnlyErrorDialogs', 'SendToRecycleBin', 'ThrowException')
}
