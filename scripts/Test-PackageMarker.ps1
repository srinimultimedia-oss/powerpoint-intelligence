$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'Package-Content.ps1')
$marker='PPI-DEMO-RIBBON-20261022-v6'
$payload=[Text.Encoding]::Unicode.GetBytes($marker)
$cases=@(
    @{Name='even byte offset';Bytes=[byte[]](@(1,2)+$payload+@(0));Expected=$true},
    @{Name='odd byte offset';Bytes=[byte[]](@(1)+$payload+@(0));Expected=$true},
    @{Name='marker at end';Bytes=$payload;Expected=$true},
    @{Name='old marker rejected';Bytes=[Text.Encoding]::Unicode.GetBytes('PPI-DEMO-RIBBON-20261022-v5');Expected=$false},
    @{Name='truncated marker rejected';Bytes=[byte[]]$payload[0..($payload.Length-2)];Expected=$false},
    @{Name='unrelated bytes rejected';Bytes=[byte[]]@(0,1,2,3);Expected=$false},
    @{Name='empty payload rejected';Bytes=[byte[]]@();Expected=$false}
)
foreach($case in $cases) {
    if((Test-PpiPayloadMarker -Bytes $case.Bytes -Marker $marker) -ne $case.Expected) {throw "FAIL $($case.Name)"}
    Write-Output "PASS $($case.Name)"
}
if(Test-PpiPayloadMarker -Bytes $payload -Marker '') {throw 'Empty marker accepted.'}
Write-Output 'PASS empty marker rejected'
