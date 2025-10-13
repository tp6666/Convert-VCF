# Complete Powershell code to process a VCF (contact) data file and convert its data for output.
# Examples (end of this file) show how to output to console, GridView or a CVS file, as required.
# Contact "cards" in the VCF file have one of several Format versions 1.x onward.
# The code below is known to work for Versions 2.x and 3.x as per the test data file supplied with this project.
# The code should also work on earlier 1.x and some later 4.x versions, but I haven't tested this.  

function New-Card {
    # this function creates a new 'card' data structure that holds data for a single contact
    # it is an ordered list of data field names and string values (initially empty)
    # changing the order of the fields here (in the definition) is safe
    # it will merely change the order of the data columns in the output
    # Do NOT change the data field names as they are referenced in code below to store data.

    ## Phone Types are Cell, Work, Home, Main. Address Types are Home, Work.
    return [PSCustomObject][ordered]@{'FullName'='';'Categories'='';'Organization'='';'WORKAddress'='';'HOMEAddress'='';'Address'='';`
'CELLPhone'='';'WORKPhone'='';'HOMEPhone'='';'Phone1'='';'Phone2'='';'Phone3'='';'WORKEmail'='';'HOMEEmail'='';'Email'='';'Note'=''}
}


function Convert-VCF{
    # main function that processes each line of the VCF data file
    # ASSERT: data for each contact is typically found between a BEGIN:VCARD and END:VCARD lines
    Param (
        # a mandatory parameter for the VCF (contact) data file to convert
        # this must be the full path and file name e.g. 'C:\temp\My Contacts.vcf'
        [Parameter(Mandatory = $true)] [string] $vcfDataFile
    )

    # check VCF data file exists
    if ( -not ( Test-Path -Path $vcfDataFile -PathType Leaf ) ) {
        Write-Host -ForegroundColor Red "Aborting - $vcfDataFile does not exist!"
        exit
    }
    # read in the VCF data file
    $content = Get-Content $vcfDataFile
    # initialise card count
    $cardCount = 0;
    # process each line of the VCF data file
    foreach ($line in $content) {
        if ($line -match "^BEGIN:VCARD.*" -or $line -match "^PRODID.*") {
            # the content of this line indicates that the lines that follow are for a new contact
            # create a new 'card' data structure for the new contact
            $currentCard = New-Card
            # increment card count
            $cardCount++;
            # initialise count of unknown phone types
            $telephones = 0;
            # initialise count of unknown email types
            $emails = 0;
            # skip to the next line of the VCF data file
            continue;
        }
        if ($line -match "^END:VCARD.*") {
            # the content of this line indicates that data for this contact is ended
            # so we output the 'card' data (to console)
            $currentCard
            # skip to the next line of the VCF data file
            continue;
        }

        if ($line -match "^PHOTO;.*") {
            # ignore photo data and skip to the next line of the VCF data file
            continue;
        }

        if ($line -match "FN:.*") {
            # process line with full name data
            $tokens = $line -split ":"
            $currentCard."FullName" = $tokens[1] -replace "\\,", ","
            # skip to the next line of the VCF data file
            continue;
        }

        if ($line -match "ORG:.*") {
            # process line with organistaion data
            $tokens = $line -split ":"
            $currentCard."Organization" = (($tokens[1] -split ';') -join "`n") -replace "\\,", ","
            # skip to the next line of the VCF data file
            continue;
        }

        if ($line -match "^ADR.*:.*" -or $line -match "^item..ADR.*:.*") {
            # process line with address data
            $tokens = $line -split ":"
            if ($tokens[0].Contains('TYPE')){
                # check if it is a type we know 
                $types = $tokens[0] -split '='
                if ('CELL,WORK,HOME'.Contains($types[1])){
                $currentCard."$($types[1])Address" = (($tokens[1] -split ';').Where({-not [string]::IsNullOrWhitespace($_)}) -join "`n") -replace "\\,", "," -replace "\\n", "`n"
                continue;
                }
            }
            # unkown address type
            $currentCard."Address" = (($tokens[1] -split ';').Where({-not [string]::IsNullOrWhitespace($_)}) -join "`n") -replace "\\,", "," -replace "\\n", "`n"
            # skip to the next line of the VCF data file
            continue;
        }

        if ($line -match "^TEL") {
            # process line with phone data, can be multiple
            $tokens = $line -split ":"
            if ($tokens[0].Contains('TYPE')){
                # check if it is a type we know 
                $types = $tokens[0] -split '='
                if ('CELL,WORK,HOME'.Contains($types[1])){
                $currentCard."$($types[1])Phone" = (($tokens[1] -split ';') -join "`n") -replace "\\,", ","
                continue;
                }
            }
            # unkown phone type
            $telephones++;
            $currentCard."Phone$telephones" = (($tokens[1] -split ';') -join "`n") -replace "\\,", ","
            # skip to the next line of the VCF data file
            continue;
        }

        if ($line -match "^EMAIL") {
            # process line with email data, can be multiple
            $type = ''
            $tokens = $line -split ":"
            if ($tokens[0].Contains('TYPE')){
                # check if it is a type we know 
                if ($tokens[0].Contains('WORK')){
                    $type = 'WORK'
                }
                elseif ($tokens[0].Contains('HOME')) {
                    $type = 'HOME' 
                }
                $currentCard."$($type)Email" = $tokens[1]
                continue;                
            }
        }

        if ($line -match "NOTE:") {
            $telephones++;
            $tokens = $line -split ":"
            $currentCard."Note" = (($tokens[1] -split ';') -join "`n") -replace "\\,", ","
            # skip to the next line of the VCF data file
            continue;
        }

        if ($line -match "CATEGORIES.*") {
            $tokens = $line -split ":"
            $currentCard."Categories" = $tokens[1]
            # skip to the next line of the VCF data file
            continue;
        }  
    }

}

# example - convert test data file (in same folder as script) and output to console
# Convert-VCF ".\Convert-VCF-TestData.vcf"

# example - convert test data file (in same folder as script) and output in GridView
# requires PowerSHell 5.x or 7.x - Out-GridView is not included in PowerShell 6.x
Convert-VCF ".\Convert-VCF-TestData.vcf" | Out-GridView

# example - convert test data file (in same folder as script) and output to
# a CSV file called PowerShell-VCF-ConvertFrom-TestData.csv (in same folder as script)
# Convert-VCF ".\Convert-VCF-TestData.vcf" | Export-Csv .\PowerShell-VCF-ConvertFrom-TestData.csv
