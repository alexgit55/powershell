#Script for querying computers in Active Directory and returning information about them

#Set the scope for the search in AD
$searchbase = 'DC=ad,DC=rvnlabtech,DC=com'

#Get all computers in the search base and store only their names
$computers=Get-ADComputer -SearchBase $searchbase -Filter * | Select-object -expandproperty Name

foreach ($computer in $computers) {
    #set up the table format and the properties we want to retrieve for each computer
    $output = @{'Computer Name' = $null;
                    'IpAddress'=$null;
                    'OperatingSystem'=$null;
                    'AvailableDriveSpace(GB)'=$null;
                    'Memory(GB)'=$null;
                    'UserProfileSize(MB)'=$null;
                    'StoppedServices'=$null
                }
  
    #Get the userprofilessize by adding the size of every file under C:\users
    $userProfilesSize=(Get-ChildItem -Path "\\$computer\c$\Users\" -Recurse -File | Measure-Object -Property Length -Sum).Sum

    #since we're making a number of wmi calls, open up a cim-session on the computer so we don't need a new one each call
    #also using a hash table so we can use splatting when calling the cim sessions
    $getCimInstParams=@{CimSession=New-CimSession -ComputerName $computer}

    #get the hard drive available space by querying the win32_logicaldisk class and the freespace property of the c: drive
    $harddrivespace=(Get-CimInstance @getCimInstParams -ClassName Win32_LogicalDisk -Filter 'DeviceID = "C:"').FreeSpace

    #get the operating system by querying the win32_operating sytem class and pulling the Caption property
    $operatingsystem=(Get-CimInstance @getCimInstParams -ClassName Win32_OperatingSystem).Caption

    #get the total memory size by querying the Win32_PhysicalMemory class and generating a sum of the capacities for each memory stick
    $memorysize=(get-cimInstance @getCimInstParams -ClassName Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum
    
    #close the cim session because we don't need to query it anymore
    Remove-CimSession -CimSession $getCimInstParams.CimSession

    #get the ip address of the device by running the resolve-dns cmdlet and selecting the ipaddress property
    $ipaddress=(resolve-DnsName -name $computer).ipaddress

    #get the stopped services on a computer with the get-service cmdlet and select where the status="stopped", viewing only the displayname of the service
    $stoppedservices=(Get-Service -ComputerName $computer | Where-Object {$PSItem.Status -eq 'Stopped'}).DisplayName

    $output.'Computer Name'=$computer
    $output.'UserProfileSize(MB)'=[int]($userProfilesSize/1MB) #The default format of this is in bytes so we want to convert to MB and remove any decimals
    $output.'AvailableDriveSpace(GB)'=[int]($harddrivespace/1GB) #same as above, but in GB instead
    $output.'OperatingSystem'=$operatingsystem
    $output.'Memory(GB)'=[int]($memorysize/1GB) #similar to userprofilessize and availablehardrivespace, converting to GB
    $output.'IpAddress'=$ipaddress
    $output.'StoppedServices'=$stoppedservices
    [pscustomobject]$output
}