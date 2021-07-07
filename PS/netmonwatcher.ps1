function Get-NetworkStatistics
{
    # Set up properties for export
    $properties = 'Protocol','LocalAddress','LocalAddressName','LocalPort','LocalPortDescription', `
                'RemoteAddress','RemoteAddressName','RemotePort','RemotePortDescription','State', `
                'ProcessName','PID', 'CaptureDateTime'


    # Date of report
    $captureDateTime = (get-date).ToString('yyyy:MM:dd HH:mm:ss')
    
    # Process netstat output
    netstat -ano |Select-String -Pattern '\s+(TCP|UDP)' | ForEach-Object {

        $item = $_.line.split(" ",[System.StringSplitOptions]::RemoveEmptyEntries)

        # Read netstat and create local varibles
        if($item[1] -notmatch '^\[::')
        {
            if (($la = $item[1] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6')
            {
                $localAddress = $la.IPAddressToString
                $localPort = $item[1].split('\]:')[-1]
            }
            else
            {
                $localAddress = $item[1].split(':')[0]
                $localPort = $item[1].split(':')[-1]
            }

            if (($ra = $item[2] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6')
            {
                $remoteAddress = $ra.IPAddressToString
                $remotePort = $item[2].split('\]:')[-1]
            }
            else
            {
                $remoteAddress = $item[2].split(':')[0]
                $remotePort = $item[2].split(':')[-1]
            }

            # Checking port descriptions, if not found is default list then check custom
            $localPortDescription = ($PublicPorts | Where-Object {$_.protocol -eq $item[0] -and $_.port -eq $localPort}).description
            if ( $null -eq $LocalPortDescription) 
                {   
                    $localPortDescription = ($CustomPorts | Where-Object {$_.protocol -eq $item[0] -and $_.port -eq $localPort}).description 
                }
            $remotePortDescription = ($PublicPorts | Where-Object {$_.protocol -eq $item[0] -and $_.port -eq $remotePort}).description
            if ( $null -eq $remotePortDescription) 
            { 
                $remotePortDescription = ($CustomPorts | Where-Object {$_.protocol -eq $item[0] -and $_.port -eq $remotePort}).description 
            }

            # Get a list of local IPs to remove possible incorrect DNS results.
            $mylocalip = (Get-NetIPAddress).IPAddress

            # Create export object, also resolve DNS and Process Names
            New-Object PSObject -Property @{
                PID = $item[-1]
                ProcessName = (Get-Process -Id $item[-1] -ErrorAction SilentlyContinue).Name
                Protocol = $item[0]
                LocalAddress = $localAddress
                LocalAddressName = if ($mylocalip | Where-Object { $_ -contains $localAddress }) { "localhost" } else { ((Resolve-DnsName $localAddress -DnsOnly -QuickTimeout -ErrorAction SilentlyContinue).NameHost)}
                LocalPort = $localPort
                LocalPortDescription = $LocalPortDescription
                RemoteAddress =$remoteAddress
                RemoteAddressName = if ($mylocalip | Where-Object { $_ -contains $remoteAddress }) { "localhost" } else { ((Resolve-DnsName $remoteAddress -DnsOnly -QuickTimeout -ErrorAction SilentlyContinue) | Select-Object -ExpandProperty NameHost -First 1)}
                RemotePort = $remotePort
                RemotePortDescription = $RemotePortDescription
                State = if($item[0] -eq 'tcp') {$item[3]} else {$null}
                CaptureDateTime = $captureDateTime
            } |Select-Object -Property $properties

            # Clean descriptions
            $LocalPortDescription = ""
            $RemotePortDescription = ""
        }
    }
}

# Primary path
$PrimaryPath = "."
# Location of TCP/UDP Port listings
$PublicPorts = Import-Csv -Path "$($PrimaryPath)\Ports\all.csv"
$CustomPorts = Import-CSV -Path "$($PrimaryPath)\Datafiles\custom.csv"
# Location of report
$ReportFile = "$($PrimaryPath)\Results\$($((get-date).ToString('yyyy-MM-dd-HH-mm-ss-')))port-connection-information.csv"


Get-NetworkStatistics | ConvertTo-Csv -NoTypeInformation |Out-File $ReportFile
