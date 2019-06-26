#Need a csv with one column of IP addresses or Hostnames located on your desktop. 
#The subnet must be pingable from the computer that this script is being run on.
#If not, then all IP addresses with show as being unused
#Authored by Dan Harrell
#Created on 6/6/2019

$Server = Get-Content c:\users\$env:USERNAME\desktop\ipaddresses.csv

Start-Transcript -force -LiteralPath $path
ForEach ($Ping in $Server)
{
    if ( 
    Test-Connection $Ping  -Count 2 -ErrorAction SilentlyContinue){
    Write-Host "$Ping is being used" -f "red"}
    else 
    {
    Write-Host "$Ping IP is currently unused" -f "white"}
}
Stop-Transcript
