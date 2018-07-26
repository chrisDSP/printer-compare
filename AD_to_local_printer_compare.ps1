<#
.SYNOPSIS
Local Printer Information Lookup Utility

Written by Chris Vincent 

.DESCRIPTION
This script is useful for comparing a print server's locally configured network printers to those on an Active Directory print server elsewhere.

This script generates a report that contains all of the locally configured printers, along with the printer name and location as it is noted in the printer's AD object. 

A list of all printers are extracted from the 'local' workstation and also from a print server. 
These are then 'joined' together on the TCP/IP hostname to assist an administrator in identifying which 'local' printers correspond to a printer on an AD print server.

The utility also accepts a one-per-line list of IP addresses to compare against the AD print server. A DNS lookup is performed in order to join on hostname.

.PARAMETER print_server
The remote AD print server. Required.

.PARAMETER output_report
Full path of the report to be output. Required.

.PARAMETER ip_list
Full path to an optional list of IP addresses to compare against, instead of using the locally configured network printers. Optional.

#>

param(
    [parameter(mandatory=$true,Position=1)]
    [ValidateNotNullOrEmpty()]
        [string]$print_server=$(throw "print_server must contain an AD print server hostname."),
    [parameter(mandatory=$true,Position=2)]
    [ValidateNotNullOrEmpty()]    
        [string]$output_report=$(throw "output_report must have the path of the desired output report file."),
    [parameter(Position=3)]
    [string]$ip_list
)


#DEBUG begin

#$ip_list = "C:\printer_ip_list.txt"
#$output_report = "C:\report_output.html"
#$print_server = "print-server.dc.myorg.xyz"

#DEBUG end


function Compare-IPListToADPS {

    #perform a DNS lookup of the IP addresses in the file at $ip_list and use the hostnames from the lookup to filter $printers_collection_print_server

    #one printer IP per line
    $ip_list_parsed = Get-Content $ip_list

    #empty target collection for IP->hostname list
    $host_list =  New-Object System.Collections.ArrayList   

    #dns lookup on IP addresses, add to hostname
    $ip_list_parsed | % {$host_list += [System.Net.Dns]::GetHostByAddress($_.ToString().Trim()) | select -ExpandProperty HostName}

     #get all print server/AD print objects whose hostnames are in the list
    $printers_collection_from_text_file = $host_list | % {$printers_collection_print_server | where -Property PortName -EQ $_}

    $ban_print_output = $printers_collection_from_text_file | select -Property PortName,Location,Name

    $ban_print_output | ConvertTo-Html | Out-File $output_report

}


function Compare-LocalPSToADPS {   

    #all local printers
    $printers_collection_local = Get-Printer | where -Property Type -EQ "Local" 



    #for each local printer in $printers_collection_local, find it's AD object among $printers_collection_print_server and select 
    #  the properties we are interested in reporting. in this case, those are the printer's location and name.
    $printers_collection_local | % {
        $_.Location = $printers_collection_print_server | where -Property PortName -EQ $_.PortName | select -ExpandProperty Location;
        Add-Member -InputObject $_ -MemberType NoteProperty -Name "AD Obj Name" -Value  $($printers_collection_print_server | 
        where -Property PortName -EQ $_.PortName | select -ExpandProperty Name)
        }

    #because both printer collections have a name property—both of which we would like to retain—we'll add an additional property 
    #  to each member of the collection called "AD Obj Name" so we can compare the printer's local name to its name in AD. 

    
    $output = $printers_collection_local | Sort-Object -Property PortName | sort -Unique PortName | select -Property PortName, Location, "AD Obj Name" 

    $output | ConvertTo-Html | Out-File $output_report
}



#all AD printers
$printers_collection_print_server = Get-Printer -ComputerName $print_server

if (!(($ip_list -EQ $null) -OR ($ip_list -EQ ""))) {

    Compare-IPListToADPS

}

else {
    
    Compare-LocalPSToADPS

}
