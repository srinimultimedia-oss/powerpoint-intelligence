$ErrorActionPreference='Stop'
$repo=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$files=@(Get-ChildItem (Join-Path $repo 'scripts') -Filter '*.ps1'; Get-ChildItem (Join-Path $repo 'easySetup') -Filter '*.ps1')
$results=@(foreach($file in $files){$tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$tokens,[ref]$errors);[pscustomobject]@{File=$file.FullName.Substring($repo.Length+1);Errors=@($errors | ForEach-Object {$_.Message});FunctionCount=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst]},$true)).Count}})
$results | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $repo 'project-export/12-testing/current-validation/powershell-ast.json')
if(@($results | Where-Object {$_.Errors.Count -gt 0}).Count){throw 'PowerShell parse errors found'}
Write-Output ('PASS '+$files.Count+' PowerShell files parsed; none executed')
