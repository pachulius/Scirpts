#diasbling all addons exept java
$jav = @('{761497BB-D6F0-462C-B6EB-D4DAF1D92D43}', '{DBC80044-A445-435B-BC74-9C25C1C588A9}')
(Gci -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Ext\Settings' -Exclude $jav).Name |
ForEach-Object {new-ItemProperty -Path Registry::$_ -Name flags -value 1 -force}

#enabling if disbledjava

$jav | ForEach-Object {
new-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Ext\Settings\$_" -name flags -Value 0 -force}
 