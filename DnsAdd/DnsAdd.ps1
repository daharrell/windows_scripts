#WIP
#Updates DNS records from a csv by removing existing entries if they exist and adds the records to all DNS servers defined in
#the variables section.
#Authored by Dan Harrell
#Created on 6/26/2019

#Variables
$DNSServer1 = “Dns.Server1.com”
$DNSServer2 = “Dns.Server2.com”
$DnsServers = $DnsServer1, $DnsServer2
$DNSZone = “ZoneName.com”

#CSV must be located on your desktop with the filename dnsupdate.csv
# Format of the csvfile should look like this:
# Name,Type,Address
# ComputerName,A,IpAddress

$Input = “c:\users\$env:USERNAME\desktop\dnsupdate.csv"

# Read the input file which is formatted as name, type, address with a header row
$records = Import-CSV $Input

# Now we loop through the file to delete and re-create records
# DNSCMD does not have a modify option so we must use /RecordDelete first followed by a /RecordAdd

ForEach ($record in $records) {
    #Capture the DNS record contents as variables from the CSV
    $recordName = $record.name
    $recordType = $record.type
    $recordAddress = $record.address

      ForEach ($Dns in $DnsServers){
      # DNSCMD DELETE command syntax
      # Deletes the existing A record
      $Delete = “dnscmd $dns /RecordDelete $DNSZone $recordName $recordType /f”
   
      # DNSCMD ADD command syntax
      $Add = “dnscmd $DNSServer1 /RecordAdd $DNSZone $recordName $recordType $recordAddress”
    
      # Now we execute the command
      # Deletes the existing A record from the DNS Server(s)
      Write-Host “Running the following command: $Delete”
      Invoke-Expression $Delete

      # Creates a new A record
      Write-Host “Running the following command: $Add”
      Invoke-Expression $Add
      }
}
