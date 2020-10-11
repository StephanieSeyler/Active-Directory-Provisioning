# Stephanie Seyler
# 2020-01-21
# v1.0.0
function new-LocationArray{
    <#
    .SYNOPSIS
        Taking a CSV table of locations buids a multidimensional hashtable with office locations
    .DESCRIPTION
        Takes in a table that was built by importing a CSV file and mapping relevant variables for an office.
        Will output a multidimensional array of the data which is able to be searched against
    .PARAMETER LocationCSV
        object that is build by importing or building a CSV table
        Currently built to use the existing headers Office, street, city, state, zip, country, fax
        The following Attributes relate to their AD Attributes in the following way:
            # office = PhysicalDeliveryOfficeName
            # zip = PostalCode
            # Country = CO
            # fax = FacsimileTelephoneNumber
            # city = L
            # State = ST
            # street = streetaddress
            # PO box = PostofficeBox
    .INPUTS
        CSV table that at minimum has the following headers Office, street, city, state, zip, country, fax
    .OUTPUTS
        Returns a 2 Dimensional array of the office locations in the following format
        $array["OfficeName"]["Streetaddress"] = 1234 Test ST.
    .EXAMPLE
        new-LocationArray -LocationCSV $OfficeCSV
    .Notes
        Author: Stephanie Seyler
        Version: 0.1.0
        Date Created: 2020-01-17cd
        Date Modified:   
    #>
    [CmdletBinding()]
    param($LocationCSV)
    $array = @{}
    foreach ($location in $LocationCSV) {
        # Update the assignments once the final data has been identified
        $array[$location.Office] = @{}
        $array[$location.Office]["PhysicalDeliveryOfficeName"] = $location.office
        $array[$location.Office]["streetaddress"] = $location.street
        $array[$location.Office]["L"] = $location.City
        $array[$location.Office]["ST"] = $location.state
        $array[$location.Office]["PostalCode"] = $location.Zip
        $array[$location.Office]["CO"] = $location.COuntry
        $array[$location.Office]["FacsimileTelephoneNumber"] = $location.Fax
    }
    return $array
}
function set-LocationAtt {
    <#
    .SYNOPSIS
        Sets office and location attributes correctly based on Array
    .DESCRIPTION
        Takes a Array created from new-locationarray and a user object and will verify their office location
        Once office location is verified the location attributes will also be verified 
    .PARAMETER User
        Active Directory user objec that will include at minimum the following Account attributes
        displayName,SAMAccountName, UserprincipalName, PhysicalDeliveryOfficeName, streetaddress,L, ST, PostalCode, CO, FacsimileTelephoneNumber, PostofficeBox
    .PARAMETER Array
        A multidimensional array that was created from the function new-locationArray
        This array will contain the information to update the AD accounts against.
    .PARAMETER DataLog
        Streamwriter object that is used to write log files too. 
        Currently using the following headers '"Date/Time","Action","Entry Type","UPN","Attribute","Value","New Value"'
    .INPUTS
        $User: Active directory user object, $array: Multidimensional hashtable built from New-LocationArray, $Datalog: Log File Streamwriter Object
    .OUTPUTS
        Outputs to a log file of all actions that the program takes and when it changes an attribute in Active Directory0
    .EXAMPLE
        set-LocationAtt -user $ADUser -array $New-LocationArray 
    .Notes
        Author: Stephanie Seyler
        Version: 0.1.0
        Date Created: 2020-01-17
        Date Modified: 
    #>
    [CmdletBinding()]
    param (
        $user,
        $Array,
        $dataLog
    )
    if($array.ContainsKey($user.PhysicalDeliveryOfficeName)){
        if($user.streetaddress -ne $array[$user.PhysicalDeliveryOfficeName]["streetaddress"]){
            # Set-ADUser -Identity $user.samaccountname -replace @{"streetaddress" = "$($array[$user.PhysicalDeliveryOfficeName]["streetaddress"])"}
            $dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""AD Change"",""Change Value"",""$($user.userprincipalname)"",""street address"",""$($user.streetaddress)"",""$($array[$user.PhysicalDeliveryOfficeName]["streetaddress"])""")
        }
        else{$dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Correct"",""No Change"",""$($user.userprincipalname)"",""street address"",""$($user.streetaddress)"",""$($array[$user.PhysicalDeliveryOfficeName]["streetaddress"])""")}
        if($user.L -ne $array[$user.PhysicalDeliveryOfficeName]["L"]){
            # Set-ADUser -Identity $user.samaccountname -replace @{"L" = "$($array[$user.PhysicalDeliveryOfficeName]["streetaddress"])"}
            $dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""AD Change"",""Change Value"",""$($user.userprincipalname)"",""City"",""$($user.L)"",""$($array[$user.PhysicalDeliveryOfficeName]["L"])""")
        }
        else{$dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Correct"",""No Change"",""$($user.userprincipalname)"",""City"",""$($user.L)"",""$($array[$user.PhysicalDeliveryOfficeName]["L"])""")}
        if($user.ST-ne $array[$user.PhysicalDeliveryOfficeName]["ST"]){
            # Set-ADUser -Identity $user.samaccountname -replace @{"ST" = "$($array[$user.PhysicalDeliveryOfficeName]["ST"])"}
            $dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""AD Change"",""Change Value"",""$($user.userprincipalname)"",""State"",""$($user.ST)"",""$($array[$user.PhysicalDeliveryOfficeName]["ST"])""")
        }
        else{$dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Correct"",""No Change"",""$($user.userprincipalname)"",""State"",""$($user.ST)"",""$($array[$user.PhysicalDeliveryOfficeName]["ST"])""")}
        if($user.PostalCode -ne $array[$user.PhysicalDeliveryOfficeName]["PostalCode"]){
            # Set-ADUser -Identity $user.samaccountname -replace @{"PostalCode" = "$($array[$user.PhysicalDeliveryOfficeName]["PostalCode"])"}
            $dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""AD Change"",""Change Value"",""$($user.userprincipalname)"",""Postal Code"",""$($user.PostalCode)"",""$($array[$user.PhysicalDeliveryOfficeName]["PostalCode"])""")
        }
        else{$dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Correct"",""No Change"",""$($user.userprincipalname)"",""Postal Code"",""$($user.PostalCode)"",""$($array[$user.PhysicalDeliveryOfficeName]["PostalCode"])""")}
        if($user.CO -ne $array[$user.PhysicalDeliveryOfficeName]["CO"]){
            # Set-ADUser -Identity $user.samaccountname -replace @{"CO" = "$($array[$user.PhysicalDeliveryOfficeName]["CO"])"}
            $dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""AD Change"",""Change Value"",""$($user.userprincipalname)"",""Country"",""$($user.CO)"",""$($array[$user.PhysicalDeliveryOfficeName]["CO"])""")
        }
        else{$dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Correct"",""No Change"",""$($user.userprincipalname)"",""Country"",""$($user.CO)"",""$($array[$user.PhysicalDeliveryOfficeName]["CO"])""")}
        if($user.FacsimileTelephoneNumber -ne $array[$user.PhysicalDeliveryOfficeName]["FacsimileTelephoneNumber"]){
            # Set-ADUser -Identity $user.samaccountname -replace @{"FacsimileTelephoneNumber" = "$($array[$user.PhysicalDeliveryOfficeName]["FacsimileTelephoneNumber"])"}
            $dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""AD Change"",""Change Value"",""$($user.userprincipalname)"",""Fax Number"",""$($user.FacsimileTelephoneNumber)"",""$($array[$user.PhysicalDeliveryOfficeName]["FacsimileTelephoneNumber"])""")
        }
        else{$dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Correct"",""No Change"",""$($user.userprincipalname)"",""Fax Number"",""$($user.FacsimileTelephoneNumber)"",""$($array[$user.PhysicalDeliveryOfficeName]["FacsimileTelephoneNumber"])""") }   
    }
    else{
        $dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""not in DB"",""No Change"",""$($user.userprincipalname)"","""",""$($user.PhysicalDeliveryOfficeName)"",""""")
    } 
}
function format-PhoneNumber {
    <#
    .SYNOPSIS
        builds a properly formatted phone number from a given string
    .DESCRIPTION
        Takes a string of either the formatted phone number or the data from msRTCSIP-line,
        Then creates a string with the proper formatting based on the country code.
        Current Countries supported:
        US(+1), CA(+1), UK(+44), FR(+33), SE(+46), DE(+49), TR(+90), AU(+61) 
    .PARAMETER phone
        String value of either a standard format or SIP formatted phone number
    .INPUTS
        Standard format or msrtcSIP-line formated string value
    .OUTPUTS
        Formatted phone number for use in "telephoneNumber" field
    .EXAMPLE
        format-phoneNumber -phone "+1 (555) 1234-5678"
    .EXAMPLE
        format-phoneNumber -phone "TEL:+155512345678"
    .EXAMPLE
        format-phoneNumber -phone "155512345678"
    .Notes
        Author: Stephanie Seyler
        Version: 0.3.0
        Date Created: 2019-12-31
        Date Modified: 

        Full Formats Supported:
        +1 (XXX) XXX-XXXX United States & Canada 
        +44 XXXX XXX XXX United Kingdom
        +33 XXX XXX XXX France
        +46 XX XXX XXX Sweden
        +49 XXX XXXX XXXX Germany
        +90 XXX XXX XXXX Turkey
        +61 X XXXX XXXX Australia
    #>
    param([string]$phone)
    # Remove all characters except Digits and add back "+" to beginning of string
    $line = "+" + ($phone -replace "\D" , "")
    switch -Wildcard ($line.substring(0,3)){
        '+44' { # Determine if it is a 9 or 10 digit phone number before formatting 
            if($line.length -eq 12){$Number = $line.substring(0,3)+" "+$line.Substring(3,3)+" "+$line.Substring(6,3)+" "+$line.Substring(9,3); break}
            if($line.length -eq 13){$Number = $line.substring(0,3)+" "+$line.Substring(3,4)+" "+$line.Substring(7,3)+" "+$line.Substring(10,3); break}}
        '+33' {$number = $line.substring(0,3)+" "+$line.Substring(3,3)+" "+$line.Substring(6,3)+" "+$line.Substring(9,3); break}
        '+46' {
            if($line.length -eq 11){$number = $line.substring(0,3)+" "+$line.Substring(3,2)+" "+$line.Substring(5,3)+" "+$line.Substring(8,3); break}
            if($line.length -eq 12){$number = $line.substring(0,3)+" "+$line.Substring(3,2)+" "+$line.Substring(5,3)+" "+$line.Substring(8,4); break}}
        '+49' {$number = $line.substring(0,3)+" "+$line.Substring(3,3)+" "+$line.Substring(6,4)+" "+$line.Substring(10,4); break}
        '+90' {$number = $line.substring(0,3)+" "+$line.Substring(3,3)+" "+$line.Substring(6,3)+" "+$line.Substring(9,4); break}
        '+61' {$number = $line.substring(0,3)+" "+$line.Substring(3,1)+" "+$line.Substring(4,4)+" "+$line.Substring(8,4); break}
        # USA & CA Should be last before Default incase of any country codes that are between +10 and +19  
        '+1?' {$number = $line.substring(0,2)+" ("+$line.Substring(2,3)+") "+$line.Substring(5,3)+"-"+$line.Substring(8,4); break}
        default {$number = $null; break}
    }
    return $number
}

function Remove-Manager{
    <#
    .SYNOPSIS
        Note the AD Manager and Remove it
    .DESCRIPTION
        Adds the Manager name to the info field in Active Directory, Removes Manager from ADObject
    .PARAMETER User
        Active Directory user object taken from the get-aduser cmdlet
    .INPUTS
        User object from Get-ADUser cmdlet
    .OUTPUTS
        No outputs only changes to Active Directory
    .EXAMPLE
        Remove-Manager -user $user
    .Notes
        Author: Stephanie Seyler   
        Version: 1.0.0
        Date Created: 2020-03-10
        Date Modified:
    #>
    param ([object] $user)
    $Header = "Generated at [{0:yyyy/MM/dd} {0:HH:mm:ss}]" -f (Get-Date) 
    if($null -eq $user.manager){$manager = " Manager: No manager, "}
    else{$manager = " Manager: " + $user.manager}
    if ($null -eq $user.info){$info = " Info: No info field"}
    else{$info = " Info: " + $user.info}
    $note = $header + $manager + $info
    $shortenedNote = $note.Substring(0, [Math]::Min($note.Length, 1023))
    Set-ADUser -Identity $user.Samaccountname -Replace @{info="$shortenedNote"}
    Set-ADUser -Identity $user.Samaccountname -clear manager
}
function Remove-AdGroups {
  <#
    .SYNOPSIS
        Note the AD groups and remove them all 
    .DESCRIPTION
        Adds the list of group names to the info field in Active Directory, Removes all AD groups excluding Domain Users
    .PARAMETER User
        Active Directory user object taken from the get-aduser cmdlet
    .INPUTS
        User object from Get-ADUser cmdlet
    .OUTPUTS
        No outputs only changes to Active Directory
    .EXAMPLE
        Remove-adGroups -user $user
    .Notes
        Author: Stephanie Seyler   
        Version: 1.0.0
        Date Created: 2020-03-10
        Date Modified:
    #>
    param ([object] $user)
    #Find AD Group Names and Convert to string list with Line Breaks
    $ADGroups = Get-ADPrincipalGroupMembership $user.Samaccountname 
    $GroupName = $null
    foreach($group in $ADGroups){
        $GroupName += $group.name +";`r`n"
    }
    #Validate that Note Variable will not go over AD Schema limit of 1023 characters for info
    $Note = $GroupName.Substring(0, [Math]::Min($GroupName.Length, 1023))
    Set-ADUser -Identity $user.Samaccountname -Replace @{info="$Note"}

    #Remove all AD Groups from object excluding "Domain Users"
    $Comparison = Compare-Object -ReferenceObject $ADgroups.name -DifferenceObject "Domain Users" 
    $RemovalGroups = $comparison.inputobject
    foreach($group in $RemovalGroups){
        remove-adgroupmember -Identity $group -members $user -confirm:$false
    }
}
function open-FileBrowser{
    <#
    .SYNOPSIS
        opens a file browser for user to easily select files with a GUI
    .DESCRIPTION
        opens the file browser to defined initial directory and allows selection of file.
        will return the path of the file that was selected by user
    .PARAMETER directory
        Location of the initial directory that will be displayed by the file browser
    .INPUTS
        location of initial directory is optional and will default to C:\
    .OUTPUTS
        path of the file that was selected by user
    .EXAMPLE
        open-fileBrowser -directory $initialDirectory
        open-fileBrowser
    .Link

    .Functionality
    
    .Notes
        Author: Stephanie Seyler
        Version: 1.0.0
        Date Created: 2019-12-12
        Date Modified: 
    #>
    param([string]$directory = "C:\")
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{InitialDirectory = $directory}
    # -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
    $null = $FileBrowser.ShowDialog()
    $FileSelected = $FileBrowser.filename
    return $FileSelected
}
function exit-ScriptCleanly {
    <#
    .SYNOPSIS
        Used to exit the script cleanly and provide cleanup and email error handling
    .DESCRIPTION
        Exits the script cleanly, sends and email if errors were encountered, closes log file and exits script
    .PARAMETER DataLog
        Log file that is being used in script
    .PARAMETER Body
        Body of email to send out
    .PARAMETER Subject
        Subject of email to send out
    .PARAMETER From
        From Address for email to use
    .PARAMETER To
        to Value for where the email will be sent to 
    .INPUTS
        no inputs are mandatory and can be used to change from the defaults
    .OUTPUTS
        no returns on script
    .EXAMPLE
        exit-ScriptCleanly
    .EXAMPLE
        exit-ScriptCleanly -to $to -from $from -subject $subject -body $body -datalog $log
    .Notes
        Author: Stephanie Seyler
        Version: 1.0.0
        Date Created: 2020-04-08
        Date Modified:
        
    #>
    param(
        [parameter(mandatory=$false,valueFromPipeline=$true)] [object] $DataLog = $dataLog,    
        [parameter(mandatory=$false,valueFromPipeline=$true)] [string] $Body = "Please review logs",
        [parameter(mandatory=$false,valueFromPipeline=$true)] [string] $Subject =  "A Scheduled Script has encountered an Error",
        [parameter(mandatory=$false,valueFromPipeline=$true)] [string] $From = "Scripting-Error-Reporter@res-group.com",
        [parameter(mandatory=$false,valueFromPipeline=$true)] $To = "Stephanie.seyler@res-group.com"
    )
    if($errorCount -gt 0){
        $anonCreds = New-Object System.Management.Automation.PSCredential("anonymous",(ConvertTo-SecureString -String "anonymous" -AsPlainText -Force))
        Send-MailMessage -From $from -To $to -Subject $subject -Credential $anonCreds -SmtpServer 'autodiscover.res-group.com' -Body $body
        $dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Success Log"",""Sent Email for error log""")
    }
    $dataLog.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""END"",""End Logging""")
    $dataLog.Close()
    exit
}