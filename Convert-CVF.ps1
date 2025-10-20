<#
Powershell code to process a VCF (contact) data file and convert its data for output.
Examples (end of this file) show how to output to console, GridView or a CVS file, as required.
VCF data comes in many formats - this code only caters for some formats / versions, NOT all.
VCF format is documented at https://www.rfc-editor.org/rfc/rfc6350
IMPORTANTLY: this code only extracts data for fields specified in the New-Card data structure.
All other contact fields, are ignored. 

Adding data extraction for other data fields, requires writting code to 
 - add the field to the data structure.
 - identify relevant lines for the field and its data (can be multi-line) in the VCF file.
 - extracting the fields data and writting it to the data structure.
 #>

function New-Card {
    <#
    this function creates a new 'card' data structure that holds fields / data for a single contact.
    the structure is an ordered list of field names and their data (string values - initially empty).
    You can safely change the order of the fields here (in the definition),
    it will merely change the order of the fields / data columns in the output.
    Do NOT change the field names as they are referenced in code.
    Phone types supported are Cell, Work and Home + 3 unknown. Address and Email types supported are Home, Work + 1 unknown.
    #>
    return [PSCustomObject][ordered]@{
        'FullName' = ''; 'Categories' = ''; 'Organisation' = ''; `
        'CELLPhone' = ''; 'WORKPhone' = ''; 'HOMEPhone' = ''; 'Phone1' = ''; 'Phone2' = ''; 'Phone3' = ''; `
        'WORKEmail' = ''; 'HOMEEmail' = ''; 'Email' = ''; 'Note' = ''; `
        'WORKAddress' = ''; 'HOMEAddress' = ''; 'Address' = ''
    }
}

function Convert-HEXData {
    <#
    this function take VCF hex encoded data and converts it to a text string
    hex encoded data are a series of =xx where xx is the hex value of a character
        e.g. =61=2D=7A=20=41=2D=5A decodes to a-z A-Z
        in hex encoded data ; is used to separate lines of data
    #>
    Param ( [Parameter(Mandatory = $true)] [string] $hexData )
    $stringOut = ''
    #first replace any ; with =0A (line feed)
    $hexData = $hexData.Replace(';','=0A')
    # split out the hex values to process
    $hexCodes = $hexData -split '='
    foreach ( $hexCode in $hexCodes ) {
        # hex code must be 2 chars
        # skip empty strings and CR/LF (CR is ASCII 13) which can happen when encoded data is multi-line
        if ($hexCode.Length -eq 2 -and [byte][char]$hexCode[0] -ne 13) {
            if ($hexCode -eq '0A') {
                # convert line feed (0A) in data to new line in text output i.e. replace LF with CR/LF
                $stringOut += "`n"
            }
            else {
                # add character of hex code to text output
                $stringOut += [char]([convert]::ToInt16($hexCode, 16))
            }
        }
    }
    return $stringOut
}

function Convert-VCF {
    <#
    main function that processes each line of the VCF data file
    ASSERT: data for each contact is found between a BEGIN:VCARD and an END:VCARD line.
    ASSERT: basic pattern is the first : is used to separate field information from field data on a VCF line.
        within field information and field data ; is used to separate types of information / data.
        e.g. PHOTO;ENCODING=BASE64;JPEG:/9j/4AAQSkZJRgABAQAA....
        field info is PHOTO;ENCODING=BASE64;JPEG i.e. filed name is PHOTO ; encoding is base 64 ; type of data is JPEG
        field data is /9j/4AAQSkZJRgABAQAA....
    ASSERT: field data can be multi-line and it can be formated as text or encoded (many types of encoding exist)
        text data lines use \ to escape special characters e.g. \, for comma \: for : and \n for new line etc.
            beware \; is a ; in the text, but ; (no slash) is used for a new line
        hex encoded data lines are a series of =xx (xx is the hex value of character) see Convert-HEXData
            in hex encoded data ; is used for a new line
        base64 encoded data lines are used for various types of data (not supported in this code)
            e.g. photo (jpeg,bmp,png etc.), file attachments etc.
    the general rules for multi-line data are
        ASSERT - when processing VCF lines for any field, all following lines starting with space or = are more data
        if the next line starts with a space, then it is more data for the field being processed
            so ignore the space and add the rest of the line to data being processed
        if the next line starts with =, then it is more hex encoded data for the field being processed
    #>
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
    # initialise variables for processing VCF data
    $fieldName = ''
    $fieldFormat = ''
    # many fields are unlikely to be multi-line data, but safer to assume they can be
    $validMultilineFieldNames = 'FullName,Categories,Organisation,CELLPhone,WORKPhone,HOMEPhone, `
    Phone1,Phone2,Phone3,WORKEmail,HOMEEmail,Email,Note,WORKAddress,HOMEAddress,Address'

    # process each line of the VCF data file
    foreach ( $line in $content ) {

        # check if this line is more field data
        if( $line[0] -eq ' ' -or $line[0] -eq '=' ){
            # ASSERT this is more field data
            if ( $fieldFormat -eq 'TEXT' -and $validMultilineFieldNames.Contains($fieldName) -and $line[0] -eq ' ' ) {
                # ASSERT this is more text data for a supported field
                # drop leading space and add data to textData
                $textData += $line.TrimStart()
                # skip to the next line of the VCF data file
                continue;
            }
            elseif( $fieldFormat -eq 'HEX' -and $validMultilineFieldNames.Contains($fieldName) -and $line[0] -eq '=' ) {
                # ASSERT this is more encoded data for a supported field
                # add data to encodedData
                $encodedData += $line
                # skip to the next line of the VCF data file
                continue;
            }
            else {
                # multi-line data for a field or format we have NOT coded for, skip to the next line of the VCF data file
                continue;
            }
        }

        # check if we were processing field data
        if( $fieldName.length -gt 0 -and $fieldFormat.length -gt 0 ){
            # ASSERT previous VCF line was field data, but this line is NOT more data
            # Need to process the 'accumulated' field data we have and add it to the current card
            if($fieldFormat -eq 'TEXT'){
                # text data lines use \ to escape special characters e.g. \, for comma \: for : and \n for new line etc.
                #   beware \; is a ; in the text, but ; (no preceding \) is a new line
                $textData = $textData -replace '\\,', ','
                $textData = $textData -replace '\\:', ':'
                $textData = $textData -replace '\\n', "`n"
                # replace ; (no preceding \) with new line 
                $textData = $textData -replace '(?<!\\);', "`n"
                # replace \; with ;
                $textData = $textData -replace '\\;', ';'
                # assign data to field
                $currentCard.($fieldName) = $textData
            }
            elseif ($fieldFormat -eq 'HEX') {
                # convert hex data to string and assign to field
                $currentCard.($fieldName) = Convert-HEXData($encodedData)
            }
            # ASSERT field data has been added to current card
            # reset field and format indicators and fall through to process current VCF line 
            $fieldName = ''
            $fieldFormat = ''
        }

        <#
        # ASSERT current VCF line is NOT multi-line data for a supported field / format
        #>
        if( $line -match "^BEGIN:VCARD.*" -or $line -match "^PRODID.*" ) {
            # the content of this line indicates that the lines that follow this, are for a new contact
            # create a new 'card' data structure for the new contact
            $currentCard = New-Card
            # reset status information for card
            $unknownPhoneTypes = 0
            $fieldName = ''
            $fieldFormat = ''
            # skip to the next line of the VCF data file
            continue;
        }
        if( $line -match "^END:VCARD.*" ) {
            # the content of this line indicates that data for this contact is ended
            # so we output the 'card' data (to console)
            $currentCard
            # skip to the next line of the VCF data file
            continue;
        }

        if( $line -match "^PHOTO;.*" ) {
            # ignore photo data and skip to the next line of the VCF data file
            $fieldName = 'Photo'
            $fieldFormat = 'BASE64'
            # skip to the next line of the VCF data file
            continue;
        }

        if( $line -match "^FN:.*" ) {
            # split line to get field info and data, only split on first :
            $fieldInfoAndData = $line -split ":", 2
            # set field name
            $fieldName = 'FullName'
            # set field format and data
            if ( $fieldInfoAndData[0].Contains('ENCODING=QUOTED-PRINTABLE')) {
                $fieldFormat = 'HEX'
                $encodedData = $fieldInfoAndData[1]
            }
            else {
                $fieldFormat = 'TEXT'
                $textData = $fieldInfoAndData[1]
            }
            # skip to the next line of the VCF data file
            continue;
        }

        if( $line -match "ORG:.*" ) {
            # split line to get field info and data, only split on first :
            $fieldInfoAndData = $line -split ":", 2
            # set field name
            $fieldName = 'Organisation'
            # set field format and data
            if ( $fieldInfoAndData[0].Contains('ENCODING=QUOTED-PRINTABLE')) {
                $fieldFormat = 'HEX'
                $encodedData = $fieldInfoAndData[1]
            }
            else {
                $fieldFormat = 'TEXT'
                $textData = $fieldInfoAndData[1]
            }
            # skip to the next line of the VCF data file
            continue;
        }

        if( $line -match "^ADR.*:.*" -or $line -match "^item..ADR.*:.*" ) {
            # split line to get field info and data, only split on first :
            $fieldInfoAndData = $line -split ":", 2
            # set field name
            if ( $fieldInfoAndData[0].Contains('WORK') ) {
                $fieldName = 'WORKAddress'
            }
            elseif ( $fieldInfoAndData[0].Contains('HOME') ) {
                $fieldName = 'HOMEAddress'
            }
            else {
                $fieldName = 'Address'
            }
            # set field format and data
            if ( $fieldInfoAndData[0].Contains('ENCODING=QUOTED-PRINTABLE') ) {
                $fieldFormat = 'HEX'
                $encodedData = $fieldInfoAndData[1]
            }
            else {
                $fieldFormat = 'TEXT'
                $textData = $fieldInfoAndData[1]
            }
            # skip to the next line of the VCF data file
            continue;
        }

        if ($line -match '^TEL' -or $line -match 'item.\.TEL') {
            # split line to get field info and data, only split on first :
            $fieldInfoAndData = $line -split ":", 2
            # set field name
            if ( $fieldInfoAndData[0].Contains('CELL') ) {
                $fieldName = 'CELLPhone'
            }
            elseif ( $fieldInfoAndData[0].Contains('WORK') ) {
                $fieldName = 'WORKPhone'
            }
            elseif ( $fieldInfoAndData[0].Contains('HOME') ) {
                $fieldName = 'HOMEPhone'
            }
            else {
                ++$unknownPhoneTypes
                $fieldName = "Phone$($unknownPhoneTypes)"
            }
            # set field format and data
            if ( $fieldInfoAndData[0].Contains('ENCODING=QUOTED-PRINTABLE')) {
                $fieldFormat = 'HEX'
                $encodedData = $fieldInfoAndData[1]
            }
            else {
                $fieldFormat = 'TEXT'
                $textData = $fieldInfoAndData[1]
            }
            # skip to the next line of the VCF data file
            continue;
        }

        if ($line -match "^EMAIL") {
            # split line to get field info and data, only split on first :
            $fieldInfoAndData = $line -split ":", 2
            # set field name
            if ( $fieldInfoAndData[0].Contains('WORK') ) {
                $fieldName = 'WORKEmail'
            }
            elseif ( $fieldInfoAndData[0].Contains('HOME') ) {
                $fieldName = 'HOMEEmail'
            }
            else {
                $fieldName = 'Email'
            }
            # set field format and data
            if ( $fieldInfoAndData[0].Contains('ENCODING=QUOTED-PRINTABLE')) {
                $fieldFormat = 'HEX'
                $encodedData = $fieldInfoAndData[1]
            }
            else {
                $fieldFormat = 'TEXT'
                $textData = $fieldInfoAndData[1]
            }
            # skip to the next line of the VCF data file
            continue;
        }

        if ($line -match "^NOTE") {
            # split line to get field info and data, only split on first :
            $fieldInfoAndData = $line -split ":", 2
            # set field name
            $fieldName = 'Note'
            # set field format and data
            if ( $fieldInfoAndData[0].Contains('ENCODING=QUOTED-PRINTABLE')) {
                $fieldFormat = 'HEX'
                $encodedData = $fieldInfoAndData[1]
            }
            else {
                $fieldFormat = 'TEXT'
                $textData = $fieldInfoAndData[1]
            }
            # skip to the next line of the VCF data file
            continue;
        }

        if ($line -match "^CATEGORIES.*") {
            # split line to get field info and data, only split on first :
            $fieldInfoAndData = $line -split ":", 2
            # set field name
            $fieldName = 'Categories'
            # set field format and data
            if ( $fieldInfoAndData[0].Contains('ENCODING=QUOTED-PRINTABLE')) {
                $fieldFormat = 'HEX'
                $encodedData = $fieldInfoAndData[1]
            }
            else {
                $fieldFormat = 'TEXT'
                $textData = $fieldInfoAndData[1]
            }
            # skip to the next line of the VCF data file
            continue;
        }  
    }

}

# example - convert test data file (in same folder as script) and output to console
Convert-VCF '.\Convert-VCF-TestData.vcf'

# example - convert test data file (in same folder as script) and output in GridView
# requires PowerSHell 5.x or 7.x or later - Out-GridView is not included in PowerShell 6.x
Convert-VCF '.\Convert-VCF-TestData.vcf' | Out-GridView

# example - convert test data file (in same folder as script) and output to
# a CSV file called PowerShell-VCF-ConvertFrom-TestData.csv (in same folder as script)
Convert-VCF '.\Convert-VCF-TestData.vcf' | Export-Csv '.\Convert-VCF-TestData.csv'
