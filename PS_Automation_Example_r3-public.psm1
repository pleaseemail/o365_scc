#
$SCC__userVariablesLocationPath = "$env:OneDrive\O365_SCC_script\"
$SCC__userVariablesFile = "scc_script_variables.ps1"
$SCC__variablePath = $SCC__userVariablesLocationPath+$SCC__userVariablesFile
#
# Example:   $tmp = SCC-New-Search -CaseName $SCC__case -CustodianSources $SCC__dataSources -SearchQuery $SCC__search -Verbose
# Example:   SCC-New-Search -CaseName 'eDiscovery Case Name' -CustodianSources 'site1','mailbox1','od4b1' -SearchQuery '[EXO] sent>01/01/2017','[SPO] author:contoso','[OD4B] contoso'
#

<# TODO:

Categories may include:
    General, Case, Search, Hold, Export

- (SEARCH) SCC-Get-SearchStatus 
-- for parameters of Continuous update for non-complete. [CD] after 60 or 120 second update?
-- add logic to find errors and report that errors exist. 

- (SEARCH) SCC-Get-SearchQuery - should it only return 1 searchquery?  Or should it parse out all searchqueries for a particular $SourceType?
- (HOLD) No support for Holds yet.
- (SEARCH) New-ComplianceSearch ; how to handle -HoldNames parameter, if at all? https://technet.microsoft.com/en-us/library/mt210905(v=exchg.160).aspx
- (SEARCH) GetFolderSearchParameters.ps1 integration (M. Hagen's script).  Would have to add in prompt user for EXO source + perhaps optional boolean flag parameter to kick off the script logic?

- (General) Add checkSCC to all functions to ensure they're usable by themselves.

- (General) Check what you're returning out of functions... do not format the objects, let the "main" function do the formatting as needed.
- (General) Add in "proper" error handling, warnings, and exceptions (if needed?)
- (General) Finish this script documentation

- (EXPORT) Need "ReportsOnly" flag so one may download the csv, and selectively export items after CSV markup
- (EXPORT) update SCC-Get-ExportStatus for "_ReportsOnly" -- along with start export

SOLUTIONS NEEDING MICROSOFT INPUT:
- (EXPORT) Export data to On-site File Storage
- (EXPORT) Export data to Tenant's own Azure environment
- (EXPORT) is the SAS token always a specific length (101 char)? e.g.: "?sv=2014-02-14&sr=c&si=eDiscoveryBlobPolicy9%7C0&sig=[48chars]"


JJ NOTES
Have a SearchName, but don't have the CaseName:
    Get-ComplianceCase -Identity (Get-ComplianceSearch -Identity $SearchName).caseID  (can Select-Object Name on previous to get human casename)

Have the CaseName, but no SearchName:
    Get-CompliancSearch -Case $CaseName

Have a SearchName or a CaseName, but no Hold Names or want to view Holds:
    Get-CaseHoldPolicy -Case $ (caseID from searchname or identity from CaseName)

2017-09-27 : Updated Get-ComplianceSearchAction to include -report parameter. Per MS engineering via MS Support: 
    "there is a -report parameter on Powershell when executing *-ComplianceSearchAction -report which provide updates."
    -- CORRECTION:  removed '-report' as it failed in my environment as unrecognized parameter.


$allCaseHolds = get-caseholdpolicy 
# lists all holds in a case
# $allCaseHolds now has one object per case hold

$oneHold = $allCaseHolds[1]
# get that specific hold

$oneHoldDetails = Get-HoldCompliancePolicy -Identity $oneHold.Identity
# gets custodians in that hold and locations and whether hold is Enabled (true/false) and Mode = Enforce; DistributionStatus = Success ; (SharePointLocation; ExcahngeLocation; PublicFolderLocation;)

$oneHold_Jobs = Get-HoldComplianceRule -Policy $oneHoldDetails.Identity
# returns the specific Names of the location holds:  HoldName_Exchange and HoldName_SharePoint
# includes the ContentMatchQuery; dates; HoldContent duration, Disabled (True/False) etc.





#> #end_of_TODO



Write-Host "Welcome to MS O365 Security & Compliance Center | eDiscovery Script"

Write-Host -foregroundcolor white -backgroundcolor red "`r`nPlease ensure you've defined variables in a separate file or that you've defined in this PS1."
Write-Host -nonewline "Separate Variable File configured as: "
    Write-Host -foregroundcolor black -backgroundcolor yellow "$($SCC__variablePath)"

# variables that may be used in this script at a global level.
$SCC__SearchEmailQuery = $null
$SCC__SearchOD4BQuery = $null
$SCC__SearchSPOQuery = $null


#This could become SCC-Start-Workflow
# - Prompt for user's case name, sources, queries, etc.
# -   (maybe this is what Clint meant for Do While not null for sources.)?

if(Test-Path -Path $SCC__variablePath) {
    # import variables into scope of this script.
    . $SCC__variablePath
} else {

    <# Example Variables for this script
    - If you opt not to specify variables in a separate file/script, then some empty variables will be defined for your use.


    The rest of this comment was from previous versions... I haven't read it to determine validity and applicability to this current version of script.
    But it sure is a lot of words.

    Below are examples of variables that may be defined.
    Some of these variables are kept here for old revision of functions whose variable names are hard coded.

    $custodianSources : You may provide a list of mailboxes, SPO sites, or OD4B sites.
                        If you intend to use this with CREATE_SEARCH function, then you may provide the list in any order and the function will identify the type of source and handle accordingly.
                        If intended for CREATE_SEARCH, for SPO and OD4B sites, you may provide with or without the trailing slash (/) as the function will place the trailing slash in (as needed for SCC).
                        An example definition is as follow:
                            $exampleCustodianSources = 'user1@domain.com', 'user2@domain.com', 'https://domain-my.sharepoint.com/personal/user1_domain_com/', 'https://domain-my.sharepoint.com/personal/user_domain_com', 'https://domain.sharepoint.com/teams/site/subsite/', 'https://domain.sharepoint.com/teams/site/subsite'

    $searchQuery      : You may provide a list of search queries that are specific to be applied to a Mailbox, SPO, OD4B or All Types.
                        Pre-pend your search query with "[EXO] ", "[SPO] ", "[OD4B] ", or "[ALL] " to indicate where that query is applicable.
                        This functionality should not be used for complex queries or for multiple criteria queries where an inadvertant overlap may occur.
                        It is up to the user to validate proper query syntax and parameters.
                        An example definition is as follow:
                            $exampleSearchQuery = "[EXO] sent>06/01/2017 AND firstname","[SPO] created>06/01/2017","[OD4B] fileextension:docx"
                            $exampleSearchQuery = "[ALL] secretProjectName"

                        An example definition that may result in unexpected results:
                            $encouragedNoUseQuery = "[EXO] sent>06/01/2017 AND firstname","[ALL] secretProjectName"
                        The above is encouraged not to be used because the results may be unexpected (is EXO query an AND or an OR with ALL query?).

                        TODO:  Could vet out how this works (logical OR or AND.. or something else?).  But could also parse out SearchQuery for counts of types of sources and if an ALL, the other type counts better be Zero, otherwise fail with a message.
    #>
    $SCC__CaseName = "Test_PS_Cmdlet_Case"

    #variables for search names as pre-determined by this script.
    #because we don't handle for repeat names of a search, adding dates as unique identifier

    $SCC__tempDate = Get-Date -Format yyyymmdd_hhmm
    $SCC__holdName = $SCC__CaseName+'-Hold-'+$SCC__tempDate
    $SCC__searchName_OD4B = $SCC__CaseName + '-OD4B-' + $SCC__tempDate
    $SCC__searchName_EXO = $SCC__CaseName + '-EXO-' + $SCC__tempDate
    $SCC__searchName_SPO = $SCC__CaseName + '-SPO-' + $SCC__tempDate
    
    # determine if we still need below variables.
    $SCC__SearchEmail = $SCC__CaseName+"Test_PS_Search_Email"
    $SCC__SearchOD4B = $SCC__CaseName+"Test_PS_Search_OD4B"
    $SCC__SearchSPO = $SCC__CaseName+"Test_PS_Search_SPO"

    $SCC__ExportEmail = $SCC__SearchEmail+"_Export"
    $SCC__ExportOD4B = $SCC__SearchOD4B+"_Export"
    $SCC__ExportSPO = $SCC__SearchSPO+"_Export"

    $SCC__CustodianEmail = ''
    $SCC__CustodianOneDrive = ''
    $SCC__CustodianSPO = ''
}


function jj-SCC-Start-Workflow {
    # Maybe create a User Prompt?
    # - Do you want to type in words parameters from CLI?  This may take awhile, please grab some coffee.
    # - Do you want to use empty variables?
    # - Do you want to exit and define $userVariablesLocationPath and $userVariablesFile variables on Lines 27 and and 28 of this script?
    # - Do you want to exit gracefully [back away slowly and pretend none of this ever happened]?
    # - Do you want to exit aggressively [please don't forget to type this value followed by the 'enter' key before starting your aggressive, over-the-top rage quit of this task].

    SCC-New-Case -CaseName $SCC__caseName -Verbose

    # add custodian sources
  ### place on hold or not
    # create searches 
  ### create search for all Case Hold content
    # start searches
    # get search status
    # start exports
    # get export status

    SCC-New-Search -CaseName $SCC__caseName -CustodianSources $SCC__sourceData -SearchQuery $SCC__example_searchQuery_Bad2 -Verbose
    return SCC-Get-SearchStatus -CaseName $SCC__caseName
}

function DO_NOT_USE-SCC-New-Hold {
<#
    .Synopsis
      TBD

    .PARAMETER $CaseName
    .PARAMETER $HoldName
    .PARAMETER $dataSources
    .PARAMETER $enableHold
    .Description
      tbd

    .Link
      LASTEDIT: 06/25/2017 20:00 PT
#>
    [CmdletBinding()]
    Param(
        [Parameter(mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$CaseName

        ,[Parameter(mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$HoldName

        ,[Parameter(mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$dataSources

        ,[Parameter(mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [bool]$enableHold
    )

    
    # parse datasources
    $exo_src 
    $spo_src
    $od4b_src


    # create holds per data source type using holdname as the platform
    #must create new-caseholdpolicy 
    #then must create new-caseholdrule and apply to that policy ; otherwise no data will be held

    #by design, this function can only apply holds or add sources to a case
    #  this function will not remove any existing holds.  

    if($enableHold) {
        New-CaseHoldPolicy -Case $CaseName -Name $HoldName -Enabled
        New-CaseHoldRule -Name 
    }



}

function SCC-Get-Variables {Get-Variable | where {$_.Name.startsWith("SCC__")} | ft -autosize Name,Value
<#
.Synopsis
    Return variable names and values that begin with text 'SCC__' available in this session.

.Description
    See synopsis.

.Link
    LASTEDIT: 06/25/2017 19:05 PT
#>} #end_of Get-SCC_Variables

function SCC-New-Case {
<#
.Synopsis
    Create an SCC eDiscovery case with name $CaseName.

.PARAMETER $CaseName
    Name of case to create.

.Description
    Function will check for existing SCC PSSession; if none exist one will be created.
    If $CaseName already exists, no new case will be created and nothing will be done to the existing case.
    If $CaseName does not exist, it will be created and function will return the SCC Case Object

.Link
    LASTEDIT: 06/12/2017 14:20 PT
#>
    [CmdletBinding()]
    Param(
        [Parameter(mandatory=$true, position=0)][string] $CaseName
    )

    if($Case.count -gt 1) {
        # TODO: I don't think this will ever trigger as PS should auto reject non-strings and throw an error... ?
        write-error 'CaseName parameter is not expected. Please define as one string value and try again.'
        return
    } 

    # Connect to SCC; keep it in a variable so not to return irrelevant data out of function
    $return = checkSCC

    # Get-ComplianceCase returns $true if $Case already exists; otherwise $false.
    if(Get-ComplianceCase -Identity $CaseName) {
        write-warning "Case name ($CaseName) already exists in S&CC."
    } else {
        $return = New-ComplianceCase -Name $CaseName
         
        $create_time = Get-Date
        write-verbose "Case name ($CaseName) has been created in S&CC: $($create_time.ToUniversalTime()) UTC" 
        $return
    }
} #end_of SCC-New-Case

function SCC-Get-SearchQueryOld {
<#
.SYNOPSIS
Returns one query from $SearchQuery that is for the given data source ($SourceType).

.DESCRIPTION
If $SearchQuery contains searches that are not applicable to $SourceType, they will be skipped.
Returns $null (e.g., ($obj.count -eq 0) or ($obj -eq $null)) if no $SearchQuery is applicable to $SourceType.
Otherwise, returns the first applicable Search Query applicable to $SourceType.

.PARAMETER SourceType
Optional parameter.
Identify what type of SourceType ([A]ll, [E]XO, [S]PO, or [O]D4B) you'd like to parse from a larger string of values.

.PARAMETER SearchQuery
Optional parameter.  
May contain one or more search queries: One for each source type of data (Exchange Online [EXO], SharePoint Online [SPO], or OneDrive for Business [OD4B]).
Alternatively, may contain one search query to be applied to all data sources ([ALL]).
Each search query should be formatted as follows: "[EXO] ", "[SPO] ", or "[OD4B] ", or "[ALL] "
#>
    [CmdletBinding()]
    Param(
        [Parameter(mandatory=$true, position=0)]
        [string]$SourceType
        
        ,[Parameter(mandatory=$true, position=1)]
        [AllowNull()]
        [string[]]$SearchQuery
    )

    # Note: $SearchQuery we AllowNull(), but PS will a $null value an empty string (''), '' -eq $null = $false ; but ''.Length -eq 0 = $true
    # Background: Easier to handle for the $null searchQuery within this function than handle the exception before calling this function.
    if($searchQuery.Length -eq 0) { return }

    # initialize search query variable
    $parsedQuery = @()

    Write-Verbose "Starting to parse $($SearchQuery.count) Search Queries.  Only parsing for type: $SourceType"

    switch ($SourceType.toLower()[0]) 
        {
            # [ALL] SearchQuery
            a {  
               foreach($search in $SearchQuery) {
                    Write-Verbose "Processing SearchQuery: $($search)"
                    #if an "[ALL] " query already processed, do not process this query
                    if($search.StartsWith("[ALL] ") -and $parsedQuery.count -eq 0) {
                        #query is for all source types
                        Write-Verbose "Found an ALL search query: $($search)"
                        $search_pared = $search.TrimStart("[ALL] ")
                        $parsedQuery += $search_pared
            
                        Write-Warning "An 'ALL' search will take precedence.  Any other defined search queries will not be executed."
                        Write-Verbose "(It is recommended that only one 'ALL' search be applied and not mixed with other types of searches.)"
                    } else {  Write-Verbose "Irrelevant search query [$($parsedQuery.count)]: $($search)"  }
                } #end_of foreach $search

                Write-Verbose "Exiting out of ALL with $($parsedQuery.count) item(s)"
                #return the relevant SearchQuery
                $parsedQuery
              } #end_of 'a' case switch 
            
            # [EXO] SearchQuery
            e { 
                foreach($search in $SearchQuery) {
                    Write-Verbose "Processing SearchQuery: $($search)"
                    if ($search.toString().StartsWith("[EXO] ") -and $parsedQuery.count -eq 0) {
                        #query is for Mailbox(es)
                        Write-Verbose "$search is to be applied to Mailbox(es)"
                        $search_pared = $search.TrimStart("[EXO] ")
                        $parsedQuery += $search_pared
                    } else {  Write-Verbose "Encountered irrelevant search query [$($parsedQuery.count)]:  $($search)"  }
                } #end_of foreach $search

                Write-Verbose "Exiting out of EXO with $($parsedQuery.count) item(s)"
                #return the relevant SearchQuery
                $parsedQuery
            } #end_of 'e' case

            # [SPO] SearchQuery
            s {
                foreach($search in $SearchQuery) {
                    Write-Verbose "Processing SearchQuery: $($search)"                    
                    #if an "[SPO] " query already processed, do not process this query
                    if ($search.toString().StartsWith("[SPO] ") -and $parsedQuery.count -eq 0) {
                        #query is for SPO site(s)
                        Write-Verbose "$search is to be applied to SPO site(s)"
                        $search_pared = $search.TrimStart("[SPO] ")
                        $parsedQuery += $search_pared
                    } else {  Write-Verbose "Encountered irrelevant search query [$($parsedQuery.count)]: $($search)"  }
                } #end_of foreach $search

                Write-Verbose "Exiting out of SPO with $($parsedQuery.count) item(s)"
                #return the relevant SearchQuery
                $parsedQuery

            } #end_of 's' case

            # [OD4B] SearchQuery
            o {
                foreach($search in $SearchQuery) {
                    Write-Verbose "Processing SearchQuery: $($search)"
                    #if an "[OD4B] " query already processed, do not process this query
                    if ($search.toString().StartsWith("[OD4B] ") -and $parsedQuery.count -eq 0) {
                        #query is for OD4B source(s)
                        Write-Verbose "$search is to be applied to OD4B site(s)"
                        $search_pared = $search.TrimStart("[OD4B] ")
                        $parsedQuery += $search_pared
                    } else {  Write-Verbose "Encountered irrelevant search query [$($parsedQuery.count)]: $($search)"  }
                } #end_of foreach $search

                Write-Verbose "Exiting out of OD4B with $($parsedQuery.count) items"
                #return the relevant SearchQuery
                $parsedQuery
            } #end_of 'o' case

            default {  Write-Host "SourceType you seek is not understood. Please enter 'a' (for ALL), 'e' (for EXO), 's' (for SPO), or 'o' (for OD4B) next time." }
    } #end_of switch $SearchQuery
} #end_of SCC-Get-SearchQueryOld

function SCC-Get-SearchQuery {
<#
.SYNOPSIS
Returns an object containing the Source Type and Query given one or more queries in $SearchQuery

.DESCRIPTION
Returns $null (e.g., ($_.count -eq 0) or ($_ -eq $null)) if $SearchQuery is empty.
Object returned has two properties: (1) SourceType and (2) Query
SourceType include: '[EXO]', '[SPO]', '[OD4B]', '[ALL]', or 'unknown'.

.PARAMETER SearchQuery
Optional parameter.  
May contain zero or more search queries related to Exchange Online [EXO], SharePoint Online [SPO], OneDrive for Business [OD4B], all data sources [ALL].
Each search query must be formatted as follows: "[EXO] ", "[SPO] ", "[OD4B] ", or "[ALL] "
#>
    [CmdletBinding()]
    Param(
        [Parameter(mandatory=$true, position=0)]
        [AllowNull()]
        [AllowEmptyString()]
        [string[]]$SearchQuery
    )

    # initialize search query variable
    $parsedQuery = @()
    
    Write-Verbose "Starting to parse $($SearchQuery.count) Search Queries."
    
    foreach ($query in $SearchQuery) {
        Write-Verbose "Processing query: $($query)"
        # Note: $SearchQuery we AllowNull() and EmptyString, but $null value and empty string ('') are treated differently, ['' -eq $null = $false ; but ''.Length -eq 0 = $true]
        # Background: Easier to handle for the $null searchQuery within this function rather than assume a $null or empty string is handled prior to calling this function.
        if($query.Length -eq 0) { Write-Verbose "SearchQuery submitted is empty" }
        
        # this is for ALL
        elseif($query.StartsWith("[ALL] ")) {
            #query is for all source types
            Write-Verbose "Found an ALL query: $($query)"
            
            $t_obj = New-Object PSObject
            $t_obj | Add-Member -MemberType NoteProperty -Name SourceType -Value '[ALL]'
            $t_obj | Add-Member -MemberType NoteProperty -Name Query -Value $query.TrimStart("[ALL] ")

            $parsedQuery += $t_obj
            
            Write-Warning "An 'ALL' search was found. Any source specific query may not be executed."
        } 
        elseif ($query.StartsWith("[EXO] ")) {
            #query is for Mailbox(es)
            Write-Verbose "Found EXO query: $($query)"

            $t_obj = New-Object PSObject
            $t_obj | Add-Member -MemberType NoteProperty -Name SourceType -Value '[EXO]'
            $t_obj | Add-Member -MemberType NoteProperty -Name Query -Value $query.TrimStart("[EXO] ")

            $parsedQuery += $t_obj
        }
        elseif ($query.StartsWith("[OD4B] ")) {
            #query is for OD4B site
            Write-Verbose "Found OD4B query: $($query)"

            $t_obj = New-Object PSObject
            $t_obj | Add-Member -MemberType NoteProperty -Name SourceType -Value '[OD4B]'
            $t_obj | Add-Member -MemberType NoteProperty -Name Query -Value $query.TrimStart("[OD4B] ")

            $parsedQuery += $t_obj
        }
        elseif ($query.StartsWith("[SPO] ")) {
            #query is for SPO site
            Write-Verbose "Found SPO query: $($query)"

            $t_obj = New-Object PSObject
            $t_obj | Add-Member -MemberType NoteProperty -Name SourceType -Value '[SPO]'
            $t_obj | Add-Member -MemberType NoteProperty -Name Query -Value $query.TrimStart("[SPO] ")

            $parsedQuery += $t_obj
        }
        else {
        # query type is unknown
            Write-Verbose "Found unknown query: $($query)"

            $t_obj = New-Object PSObject
            $t_obj | Add-Member -MemberType NoteProperty -Name SourceType -Value 'unknown'
            $t_obj | Add-Member -MemberType NoteProperty -Name Query -Value $query

            $parsedQuery += $t_obj
        }
    } #end_of foreach $query
    $parsedQuery
} #end_of SCC-Get-SearchQuery



function SCC-New-Search {
<#
.SYNOPSIS
Create and start searches for an O365 Security & Compliance Center eDiscovery Case.
Example: CREATE_SEARCH('eDiscovery Case Name',('site1','mailbox1','od4b1'),('[EXO] sent>01/01/2017','[SPO] author:contoso','[OD4B] contoso'))

.DESCRIPTION
Function will check for CaseName existence; if none exist user will be prompted.
Function will check for existing SCC PSSession; if none exist one will be created.

.PARAMETER CaseName
Identity of case within S&CC eDiscovery.

.PARAMETER CustodianSources
A string or array of strings containing one or more mailboxes (inactive or active), SharePoint Online sites, or OneDrive for Business sites.

.PARAMETER SearchQuery
Parameter may contain a max of three strings (array of strings), one for each source of data (Exchange Online [EXO], SharePoint Online [SPO], or OneDrive for Business [OD4B]).
Precede each string with the following characters without double quotes and followed by a space before your query to specify the source of data to apply the query to:
"[EXO]", "[SPO]", or "[OD4B]".

#>

    [CmdletBinding()]
    Param(
        [Parameter(mandatory=$true, position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$CaseName
        
        ,[Parameter(mandatory=$true, position=1)]
        [ValidateNotNullOrEmpty()]
        [string[]]$CustodianSources

        ,[Parameter(mandatory=$false, position=2)]
        [string[]]$SearchQuery
    )

    Write-Verbose 'Checking for $($CaseName).  Will create if does not exist.'
    $case = SCC-New-Case $CaseName

    # obtain object with search and search type; leave unknown sources in case we want to reference later.
    $parsedSources = SCC-Get-CustodianLocations -CustodianSources $CustodianSources

    # initialize data source variables
    $source_OD4B = @()
    $source_EXO = @()
    $source_SPO = @()
    $source_unprocessed = @()

    # load source-specific variables from the master parsedSources
    foreach ($parsedSource in $parsedSources) {
        if($parsedSource.Type -eq 'EXO') {  $source_EXO += $parsedSource.Location  }
        elseif($parsedSource.Type -eq 'SPO') {  $source_SPO += $parsedSource.Location  }
        elseif($parsedSource.Type -eq 'OD4B') {  $source_OD4B += $parsedSource.Location  }
        elseif($parsedSource.Type -eq 'unknown') {  $source_unprocessed += $parsedSource.Location  }
        else { write-warning "unhandled case from SCC-Get-CustodianLocations output into SCC-New-Search." }
    }

    # Get SearchQuery prepared
    $searchQueries = SCC-Get-SearchQuery -SearchQuery $SearchQuery
    $search_ALL = @()
    $search_EXO = @()
    $search_SPO = @()
    $search_OD4B = @()
    $search_unprocessed =@()

    foreach($search in $searchQueries) {
        if($search.SourceType -eq '[ALL]') {  $search_ALL += $search  }
        elseif($search.SourceType -eq '[EXO]') {  $search_EXO += $search  } 
        elseif($search.SourceType -eq '[SPO]') {  $search_SPO += $search  } 
        elseif($search.SourceType -eq '[OD4B]') {  $search_OD4B += $search  } 
        elseif($search.SourceType -eq 'unknown') {  $search_unprocessed += $search  } 
        else {  Write-Warning "Unhandled sourceType from SCC-Get-SearchQuery within SCC-New-Search"  } 

    }

    #check for multiple queries against the same source and warn user if present
    if( ($search_exo.count -ge 2) -or
        ($search_SPO.count -ge 2) -or
        ($search_OD4B.count -ge 2)) {
            Write-Warning "More than one search query exists for the same source type. Unintended search query results may ensue."
    }
    # explicitly check and warn for multiple ALL queries
    if($search_ALL.count -ge 2) {Write-Warning "More than one ALL query exists.  Unintended search query results may ensue."}

    if($search_ALL.count -ne 0) {
        # An ALL search was found, process the ALL on all source types
        Write-Verbose "An [ALL] query was found.  Any specific source queries (EXO, SPO, OD4B) will be ignored."

        #is there an EXO custodian source to search?
        if($source_EXO.count -gt 0) {
            # 'AllowNotFoundExchangeLocationsEnabled' is required if you want to include inactive mailboxes in the search, because inactive mailboxes don't resolve as regular mailboxes.
            # source: https://technet.microsoft.com/en-us/library/mt210905(v=exchg.160).aspx
            if(New-ComplianceSearch -Case $CaseName -Name $SCC__searchName_EXO -exchangelocation $source_EXO -ContentMatchQuery $search_all[0].Query -AllowNotFoundExchangeLocationsEnabled $true <#need to add unindexed option to include all data#> ) {
                Start-Sleep -Seconds 2
                Start-ComplianceSearch -Identity $SCC__searchName_EXO
            }

        } else {
            Write-Warning "No mailbox search created (no mailboxes identified to search)."
        }

        #is there an OD4B custodian source to search?
        if($source_OD4B.count -gt 0) {
            ## CREATE OD4B SEARCH
            if(New-ComplianceSearch -Case $CaseName -Name $SCC__searchName_OD4B -SharePointLocation $source_OD4B -ContentMatchQuery $search_all[0].Query ) {
                Start-Sleep -Seconds 2
                Start-ComplianceSearch -Identity $SCC__searchName_OD4B
            }
        } else {
            Write-Warning "No OD4B search created (OneDrive for Business sites identified to search)."
        }   

        #is there an SPO custodian source to search?
        if($source_SPO.count -gt 0) {
            ## CREATE SPO SEARCH
            # Currently -OneDriveLocation paremeter to New-ComplianceSearch is reserved for internal MS use.
            # https://technet.microsoft.com/en-us/library/mt210905(v=exchg.160).aspx
            if(New-ComplianceSearch -Case $CaseName -Name $SCC__searchName_SPO -SharePointLocation $source_SPO -ContentMatchQuery $search_all[0].Query) {
                Start-Sleep -Seconds 2
                Start-ComplianceSearch -Identity $SCC__searchName_SPO
            }
        } else {
            Write-Warning "No SharePoint Online sites identified to search."
        }
    } else {
        # No ALL search was found, so process the proper Search Query on the proper Source Type
        # Unknown  what a null/empty -ContentMatchQuery will do to a search, so explicitly treating No Content Match Query separate from a Content Match Query existing.

        if($source_EXO.count -gt 0) {
            #mailboxes to be searched; parse if any SearchQuery to be applied

            # 'AllowNotFoundExchangeLocationsEnabled' is required if you want to include inactive mailboxes in the search, because inactive mailboxes don't resolve as regular mailboxes.
            # source: https://technet.microsoft.com/en-us/library/mt210905(v=exchg.160).aspx
            if($search_EXO.count -gt 0) {
                # there exists a search query
                if(New-ComplianceSearch -Case $CaseName -Name $SCC__searchName_EXO -exchangelocation $source_EXO -ContentMatchQuery $search_EXO[0].Query -AllowNotFoundExchangeLocationsEnabled $true <#need to add unindexed option to include all data#> ) {
                    Start-Sleep -Seconds 2
                    Start-ComplianceSearch -Identity $SCC__searchName_EXO
                }
            } else {
                # there does not exist a search query
                if(New-ComplianceSearch -Case $CaseName -Name $SCC__searchName_EXO -exchangelocation $source_EXO -AllowNotFoundExchangeLocationsEnabled $true <#need to add unindexed option to include all data #>) {
                    Start-Sleep -Seconds 2
                    Start-ComplianceSearch -Identity $SCC__searchName_EXO
                }
            }
        } else {
            Write-Host "No mailboxes identified to search."
        }

        if($source_OD4B.count -gt 0) {
            #od4b's to be searched; parse if any SearchQuery to be applied

            ## CREATE OD4B SEARCH
            if($search_OD4B.count -gt 1) {
                # there exists a search query
                if(New-ComplianceSearch -Case $CaseName -Name $SCC__searchName_OD4B -SharePointLocation $source_OD4B -ContentMatchQuery $search_OD4B[0].Query) {
                    Start-Sleep -Seconds 2
                    Start-ComplianceSearch -Identity $SCC__searchName_OD4B
                }
            } else {
                #there does not exist a search query
                if(New-ComplianceSearch -Case $CaseName -Name $SCC__searchName_OD4B -SharePointLocation $source_OD4B) {
                    Start-Sleep -Seconds 2
                    Start-ComplianceSearch -Identity $SCC__searchName_OD4B
                }
            }
        } else {
            Write-Warning "No OneDrive for Business sites identified to search."
        }   

        if($source_SPO.count -gt 0) {
            #SPO's to be searched; parse if any SearchQuery to be applied
            
            ## CREATE SPO SEARCH
            # Currently -OneDriveLocation paremeter to New-ComplianceSearch is reserved for internal MS use.
            # https://technet.microsoft.com/en-us/library/mt210905(v=exchg.160).aspx
            if($search_SPO.count -gt 1) {
                #there exists a search query
                if(New-ComplianceSearch -Case $CaseName -Name $SCC__searchName_SPO -SharePointLocation $source_SPO -ContentMatchQuery $search_SPO[0].Query) {
                    Start-Sleep -seconds 2
                    Start-ComplianceSearch -Identity $SCC__searchName_SPO
                }
            } else {
                #there does not exist a search query
                if(New-ComplianceSearch -Case $CaseName -Name $SCC__searchName_SPO -SharePointLocation $source_SPO) {
                    Start-Sleep -seconds 2
                    Start-ComplianceSearch -Identity $SCC__searchName_SPO
                }
            }
        } else {
            Write-Warning "No SharePoint Online sites identified to search."
        }
    } #end_of else (no ALL search found)

    write-Verbose "Successfully identified $($source_OD4B.count + $source_od4b.count + $source_exo.count) Custodian Sources."
    write-Verbose " - Mailbox source count: $($source_exo.count) [$search_exo]"
    write-Verbose " - OD4B source count: $($source_od4b.count) [$source_od4b]"
    write-verbose " - SPO source count: $($source_spo.count) [$source_spo]"
} #end_of SCC-New-Search

function SCC-Get-CustodianLocations {
<#
.SYNOPSIS
Will parse the CustodianSources input into an object with identified source type locations.
Duplicate sources will be removed.
Returns an object with two properties: 'Type' and 'Location'.


.DESCRIPTION
Type has 4 options: 'EXO', 'OD4B', 'SPO', or 'unknown'.
Location contains the custodian source location (email address, OD4B site url, SPO site URL).

.PARAMETER CustodianSources
A string or array of strings containing one or more mailboxes (inactive or active), SharePoint Online sites, or OneDrive for Business sites.

.PARAMETER noUnknowns
A switch parameter.  By default is $false.  If switch is used (-noUnknowns), will be $true and will not add any 'unknown' Types to the object.

#>

    [CmdletBinding()]
    Param(
        [Parameter(mandatory=$true, position=1)]
        [ValidateNotNullOrEmpty()]
        [string[]]$CustodianSources

        ,[Parameter(mandatory=$false, position=1)]
        [switch]$noUnknowns
    )

    # initialize data source variables
    $source_arr = @()

    Write-Verbose "Parsing $($CustodianSources.count) sources"

    foreach($custodianSource in $CustodianSources) {

        if($custodianSource.Contains("@")) {
            #item is a mailbox
            #check for duplicate
            $dupe = $false
            foreach ($obj in $source_arr) {
                if($obj.Location -eq $custodianSource) {
                    Write-Verbose "Duplicate Mailbox (skipping): $custodianSource"
                    $dupe = $true
                }
            }

            #add source if not present
            if(!$dupe) {
                Write-Verbose "Adding Mailbox:  $custodianSource"

                $t_obj = New-Object PSObject
                $t_obj | Add-Member -MemberType NoteProperty -Name Type -Value EXO
                $t_obj | Add-Member -MemberType NoteProperty -Name Location -Value $custodianSource

                $source_arr += $t_obj
            }

        } elseif ($custodianSource.toString().StartsWith("https://") -and $custodianSource.toString().Contains(".sharepoint.com/personal")) {
            #item is a OD4B site
            if(!$custodianSource.EndsWith("/")) {
                #URL does not contain trailing slash
                $custodianSource = $custodianSource + '/'
            }

            #check for duplicate
            $dupe = $false
            foreach ($obj in $source_arr) {
                if($obj.Location -eq $custodianSource) {
                    Write-Verbose "Duplicate OD4B (skipping): $custodianSource"
                    $dupe = $true
                }
            }

            #add source if not present
            if(!$dupe) {
                Write-Verbose "Adding OD4B: $custodianSource"

                $t_obj = New-Object PSObject
                $t_obj | Add-Member -MemberType NoteProperty -Name Type -Value OD4B
                $t_obj | Add-Member -MemberType NoteProperty -Name Location -Value $custodianSource

                $source_arr += $t_obj
            }
        } elseif ($custodianSource.toString().StartsWith("https://") -and $custodianSource.toString().Contains(".sharepoint.com")) {
            #item is a SPO site
            if(!$custodianSource.EndsWith("/")){
                #URL does not contain trailing slash
                $custodianSource = $custodianSource + '/'
            }

            #check for duplicate
            $dupe = $false
            foreach ($obj in $source_arr) {
                if($obj.Location -eq $custodianSource) {
                    Write-Verbose "Duplicate SPO (skipping): $custodianSource"
                    $dupe = $true
                }
            }

            #add source if not present
            if(!$dupe) {
                Write-Verbose "Adding SPO: $custodianSource"

                $t_obj = New-Object PSObject
                $t_obj | Add-Member -MemberType NoteProperty -Name Type -Value SPO
                $t_obj | Add-Member -MemberType NoteProperty -Name Location -Value $custodianSource

                $source_arr += $t_obj
            }
        } else {
            Write-Warning "Unknown source type: $custodianSource" 
            if(!$noUnknowns) {
                $t_obj = New-Object PSObject
                $t_obj | Add-Member -MemberType NoteProperty -Name Type -Value unknown 
                $t_obj | Add-Member -MemberType NoteProperty -Name Location -Value $custodianSource

                $source_arr += $t_obj
            }
        }
    } #end_of foreach custodian
    $source_arr
} #end_of SCC-Get-CustodianLocations

function SCC-Get-SearchStatus {
<#
    .Synopsis
      Returns the status of one or more searches given ($SearchName) or within a case ($CaseName) as selected by user.
      Caution: A $SearchName and a $CaseName can have the same "name".  Be certain whether the name you have is for a $CaseName or $SearchName.

      If no case or search name is provided, you will be guided through selecting an available case and checking status of one or more searches within the case.

    .PARAMETER $CaseName
    .PARAMETER $SearchName
    .Description
      tbd

    .Link
      LASTEDIT: 06/25/2017 20:00 PT
#>
    [CmdletBinding()]
    Param(
        [Parameter(mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string[]]$CaseName

        ,[Parameter(mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string[]]$SearchName
    )
    Write-Verbose "entered: SCC-Get-SearchStatus"

    # $CaseName nor $SearchName provided
    if(($CaseName -eq $null) -and ($SearchName -eq $null)) {
        $userSelection = Read-Host -Prompt "You did not provide a Case or Search name, would you like to list the eDiscovery Cases available to you? [Y]es/[N]o" 

        # User wants to quit.
        if ($userSelection.toLower().StartsWith("n")) { Write-Host "Please come back with a Case or Search name."; return }

        # User provided unhandled input
        if (!$userSelection.toLower().StartsWith("y")) { Write-Error "Did not understand your entry $($userSelection).`r`nExpected a 'y' or an 'n'."; return }

        # User wants to list all eDiscovery Cases
        # Create PSSession if needed
        checkSCC

        $availableCases = Get-ComplianceCase
            
        if($availableCases.count -eq 0) {  Write-Error "You do not have permissions to any eDiscovery Cases or no eDiscovery Cases exist."; return }

        Write-Verbose "There are $($availableCases.count) cases available to you."

        #create case reference to reference from User Input
        $availableCases | Add-Member -MemberType NoteProperty -Name SCC_CaseID -Value 0
        $i=0; foreach($obj in $availableCases){ $obj.SCC_CaseID = $i++}
        
        #Write this table to Console only (Out-String | Write-Host)
        $availableCases | Sort-Object CreatedDateTime | format-table -AutoSize SCC_CaseID,Name,Status,@{Name="CreatedDate";Expression={$_.CreatedDateTime.toString("yyyy-MM-dd")}} | Out-String | %{Write-Host $_}

        #get user input for which case they'd like to check searches of (check all searches or check individual search)
        $caseID = Read-Host -Prompt "Which single eDiscovery Case would you like to get searches? (SCC_CaseID # or 'e' for exit)"

        if($caseID.toLower() -eq 'e') {  Write-Host "You chose (e)xit.  Goodbye."; return }
        
        $caseSearchList = Get-ComplianceSearch -Case $availableCases[$caseID].Identity
        Write-Verbose "There are $($caseSearchList.count) searches in the case: $($caseSearchList.Name)"

        #create search reference to reference from User Input
        $caseSearchList | Add-Member -MemberType NoteProperty -Name SCC_SearchID -Value 0
        $i=0; foreach($obj in $caseSearchList){ $obj.SCC_SearchID = $i++}
        
        #Write this table to Console only (Out-String | Write-Host)
        $caseSearchList | Sort-Object JobEndTime | Format-Table -AutoSize SCC_SearchID,Name,Status,JobEndTime | Out-String | %{Write-Host $_} 

        #get user input to determine which searches they'd like to check status
        $searchID = Read-Host -Prompt "Would you like to check status for [a]ll searches, [s]ome searches, or [e]xit?"

        if($searchID.toLower() -eq 'e') {  Write-Host "You chose (e)xit.  Goodbye."; return  } 
        elseif($searchID.toLower() -eq 'a') {
            #user wants to view status of ALL searches in the case
            $continuous = Read-Host -Prompt "Would you like to constantly update search results of incomplete items? [Y]es/[N]o"

            if($continuous.toLower().StartsWith("n")) {
                #get status of all searches One time.
                Write-Verbose "Getting all searches one time"
                SCC-Get-SearchStatus -CaseName $availableCases[$caseID].Identity

            } elseif (!$continuous.toLower().StartsWith("y")) { Write-Output "Did not understand your entry $($continuous).`r`nExpected a 'y' or 'n'."; return } 
            else {
                #user wants to continuously update all searches that are in progress
                
                # https://www.sapien.com/blog/2014/11/18/removing-objects-from-arrays-in-powershell/
                [System.Collections.ArrayList]$searchobjArray = @()

                #recall: we have $caseSearchList, create an arraylist based on this for easier removal of objects from array.
                foreach($searchobj in $caseSearchList) {
                    $searchobjArray += $searchobj.SCC_SearchID
                }

                #get while loop for constantly updating search results
                $searches_notStarted = @()
                $searches_completed = @()

                while($searchobjArray.count -ne 0) {
                    foreach($searchobj in $caseSearchList) {
                        # create list of InProgress objects
                        if($searchobj.Status -eq 'NotStarted') {
                            #search has not started, remove from checking list
                            Write-Verbose "$($searchObj.Name) has status $($searchobj.Status). Removing $($searchObj.SCC_SearchID) from searchobjArray."

                            $searchobjArray.Remove($searchObj.SCC_SearchID)

                            #add to searches_notStarted for notifying user
                            $searches_notStarted += $searchObj
                        } elseif ($searchobj.Status -eq 'Completed') {
                            Write-Verbose "$($searchObj.Name) has status $($searchobj.Status). Removing $($searchObj.SCC_SearchID) from searchobjArray."

                            #search has completed, so may remove from constant check
                             $searchobjArray.Remove($searchObj.SCC_SearchID)    
                             
                             #add to searches_completed for notifying user.  
                             $searches_completed += $searchObj              
                        } else {
                            #This search object is likely InProgress
                            Write-Verbose "$($searchObj.Name) has status $($searchobj.Status).`r`nSleeping for 5 seconds."

                            sleep -Seconds 5
                            $search_status = Get-ComplianceSearch -Identity $searchobj.Name
                            
                            $search_complete = $false
                            while(!$search_complete){
                                $search_status = Get-ComplianceSearch -Identity $searchobj.Name
                                if(($search_status.Status -eq "Completed") -and  ($search_status.JobProgress -eq 100)) {
                                    $search_complete = $true
                                    Write-Verbose "$($search_status.Name) has completed."

                                    $searchobjArray.Remove($searchObj.SCC_SearchID)
                                    $searches_completed += $search_status
                                } else {
                                    Write-Host "$($searchObj.Name) is not completed (Progress: $search_status.JobProgress / 100)"
                                    Start-Sleep -s 5
                                }
                            } #end_of_while Search not Completed or Not Started
                        } #end_of_else search is not completed or not started
                    } #end_of_foreach
                } #end_of while $searchobjArray.count -ne 0

                Write-Host "$($availableCases[$caseID].Name) contains $($caseSearchList.count) searches`r`n`t$($searches_completed.count) Completed `r`n`t$($searches_notStarted.count) NotStarted"
            } #end_of user wants to continuously update all searches that are in progress
        }#end_of All search status elseif
        elseif ($searchID.toLower() -ne 's') {  Write-Error "Did not understand your entry $($continuous).`r`nExpected [s]ome, [a]ll, or [e]xit."; return }
        else {
            #User would like to search [s]ome searches
            $searchID = @()
            do {
                $userItem = Read-Host "If done, type 'd'. Otherwise, please enter one SCC_SearchID and hit Enter"
                #TODO input validation...boring.  

                if($userItem.toLower() -ne 'd') { $searchID += $userItem }
            } until ($userItem.toLower() -eq 'd')

            Write-Verbose "User input for SCC_SearchID is: $($searchID)"

            $continuous = Read-Host -Prompt "Would you like to constantly update search results of your selected items? [Y]es/[N]o"

            if($continuous.toLower() -eq 'y') {
                #continuously update results until they are all "Completed"
                #user wants to continuously update all searches that are in progress

                $caseSearchSelected = @()
                foreach($item in $searchID) {
                    Write-Verbose "item = $($item): $($caseSearchList[$item].Name)"
                    $caseSearchSelected += Get-ComplianceSearch -Identity $caseSearchList[$item].Identity
                }

                #create search reference to reference from User Input
                $caseSearchSelected | Add-Member -MemberType NoteProperty -Name SCC_SearchID -Value 0
                $i=0; foreach($obj in $caseSearchSelected){ $obj.SCC_SearchID = $i++}

                $caseSearchSelected | ft -AutoSize SCC_SearchID,Name,Status

                # https://www.sapien.com/blog/2014/11/18/removing-objects-from-arrays-in-powershell/
                [System.Collections.ArrayList]$searchobjArray = @()
                foreach($searchobj in $caseSearchSelected) { $searchobjArray += $searchobj.SCC_SearchID }

                #get while loop for constantly updating search results
                $searches_notStarted = @()
                $searches_completed = @()

                while($searchobjArray.count -ne 0) {
                    foreach($searchobj in $caseSearchSelected) {
                        # create list of InProgress objects
                        if($searchobj.Status -eq 'NotStarted') {
                            #search has not started, remove from checking list
                            Write-Verbose "$($searchObj.Name) has status $($searchobj.Status). Removing $($searchObj.SCC_SearchID) from searchobjArray."

                            $searchobjArray.Remove($searchObj.SCC_SearchID)

                            #add to searches_notStarted for notifying user
                            $searches_notStarted += $searchObj
                        } elseif ($searchobj.Status -eq 'Completed') {
                            Write-Verbose "$($searchObj.Name) has status $($searchobj.Status). Removing $($searchObj.SCC_SearchID) from searchobjArray."

                            #search has completed, so may remove from constant check
                             $searchobjArray.Remove($searchObj.SCC_SearchID)    
                             
                             #add to searches_completed for notifying user.  
                             $searches_completed += $searchObj              
                        } else {
                            #This search object is likely InProgress
                            Write-Verbose "$($searchObj.Name) has status $($searchobj.Status).`r`nSleeping for 5 seconds."

                            sleep -Seconds 5
                            $search_status = Get-ComplianceSearch -Identity $searchobj.Name
                            
                            $search_complete = $false
                            while(!$search_complete){
                                $search_status = Get-ComplianceSearch -Identity $searchobj.Name
                                if(($search_status.Status -eq "Completed") -and  ($search_status.JobProgress -eq 100)) {
                                    $search_complete = $true
                                    Write-Verbose "$($search_status.Name) has completed."

                                    $searchobjArray.Remove($searchObj.SCC_SearchID)
                                    $searches_completed += $search_status
                                } else {
                                    Write-Host "$($searchObj.Name) is not completed (Progress: $search_status.JobProgress / 100)"
                                    Start-Sleep -s 5
                                }
                            } #end_of_while Search not Completed or Not Started
                        } #end_of_else search is not completed or not started
                    } #end_of_foreach
                } #end_of while $searchobjArray.count -ne 0

                Write-Host "You selected $($searchID.count) searches from $($availableCases[$caseID].Name)`r`n`t$($searches_completed.count) Completed `r`n`t$($searches_notStarted.count) NotStarted"

            } elseif($continuous.ToLower() -eq 'n') {
                $searchStatus = @()
                foreach($item in $searchID) {
                    $searchStatus += Get-ComplianceSearch -identity $caseSearchList[$item].Identity
                }
                $searchStatus | ft -autosize Name,Status,jobprogress
            } else {  Write-Error "Did not understand your entry $($continuous)."; return  }
       }
    } #end_of No SearchName or CaseName provided
    
    elseif (($CaseName -eq $null) -and ($SearchName -ne $null)) {
        # $SearchName provided
        checkSCC
        $searchName_arr = @()
        foreach($searchobj in $SearchName) {
            $searchname_arr += Get-ComplianceSearch -Identity $searchobj
        }
        $searchname_arr | ft -AutoSize Name,Status,JobProgress
    } elseif (($CaseName -ne $null) -and ($searchname -eq $null)) {
        # $CaseName provided
        checkSCC

        $caseName_arr = @()
        foreach($caseobj in $casename) {
            $caseName_arr += Get-ComplianceSearch -Case $caseobj
        }
        $caseName_arr | ft -AutoSize Name,Status,JobProgress
    } else { Write-Error "Please provide either a Case name or a Search Name.  Not both." ; return }
} #end_of_SCC-Get-SearchStatus


function SCC-Start-Export {
    [CmdletBinding()]
    Param(
        [Parameter(mandatory=$true, position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$SearchName

        ,[Parameter(mandatory=$true, position=1)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceType

        ,[Parameter(mandatory=$false, position=2)]
        [string]$EXOFormat

        ,[Parameter(mandatory=$false, position=3)]
        [string]$Scope

# TODO -ExchangeArchiveFormat \"IndividualMessage\"


    )
<#
    .Synopsis
      Start the data export given the search name, source type, and scope of data requested.

    .PARAMETER $SearchName
      This parameter specifies the name of the existing compliance search to associate with the export compliance search action.

    .PARAMETER $SourceType
      Valid values are: (1) EXO, (2) SPO, or (3) OD4B

    .PARAMETER $EXOFormat
      (Optional) Applies to EXO source only. Valid values are Unknown | FxStream | Mime | Msg | BodyText.  If no value provided, FxStream is assumed.

    .PARAMETER $Scope
      (Optional) This parameter specifies the items to include when the action is Export.  If no value provided, IndexedItemsOnly is assumed.
      Valid values are: (1) IndexedItemsOnly, (2) UnindexedItemsOnly, or (3) BothIndexedAndUnindexedItems

    .Description
      tbd

    .Link
      LASTEDIT: 06/25/2017 20:00 PT
#>

    checkSCC

    # If no $Scope is provided, set it to IndexedItemsOnly
    if($Scope.length -eq 0) {  $Scope = 'IndexedItemsOnly'  }

    # Determine if Scope provided is valid
    $isValidScope = $false
    $validScope = 'indexeditemsonly','unindexeditemsonly','bothindexedandunindexeditems'
    foreach ($i in $validScope) { if($scope.tolower() -eq $i){$isValidScope = $true} }

    if(!$isValidScope) {  Write-error "Did not understand your Scope, please check spelling" ; return  }

    if($EXOFormat.length -eq 0) { $EXOFormat = 'FxStream' }

    switch ($SourceType.tolower()[0])
        {
            e { checkSCC; write-verbose "you chose EXO"; New-ComplianceSearchAction -SearchName $SearchName -Export -Format $EXOFormat -Scope $Scope -Force }
            o { checkSCC; write-verbose "you chose OD4B"; New-ComplianceSearchAction -SearchName $SearchOD4B -Export -IncludeSharePointDocumentVersions $true -Scope $Scope -Force }
            s { checkSCC; write-verbose "you chose SPO"; New-ComplianceSearchAction -SearchName $SearchSPO -Export -IncludeSharePointDocumentVersions $true -Scope $Scope -Force }
            default {write-error "Exiting. Did not understand SourceType"}
        }
} #end_of SCC-Start-Export

function SCC-Get-ExportStatus {
    [CmdletBinding()]
    Param(
        [Parameter(mandatory=$true, position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$exportName

        ,[Parameter(mandatory=$false, position=1)]
        [ValidateNotNullOrEmpty()]
        [switch]$showToken
    )

    checkSCC

    $export_status = Get-ComplianceSearchAction -identity $exportName -Details
    $sleepTime = 5

    #Assume export is not complete and there will be failed sources until explicitly found to be otherwise.
    $export_complete = $false
    $export_contains_fail = $true
    $export_contains_err = $true

    if($export_status) {
        $loops = 0
        while(!$export_complete) {
            
            #$export_status | fl | Out-String | %{Write-Verbose $_} 

            if( $export_status.Results.Contains("; Progress: 100.00 %") -or $export_status.Results.Contains("; Export status: Completed") -or ($export_status.Results -match "; Completed time: \d") ) {
                # Export is now at 100.00 % completion
                # or Export status = Completed
                # or Export has a Completed Time
                # the three items above are observed to not always be present, so must evaluate each of the three items separately
                $export_complete = $true

                if($showToken) {
                    $export_status = Get-ComplianceSearchAction -identity $exportName -IncludeCredential -Details
                    $token = $export_status.Results.IndexOf("; SAS token:")
                    
                    if($token -gt 0 ) {
                    #get token assumes Scenario occurs directly after SAS token.  If not, will have to acocunt for later.
                        $export_status.results.substring($token+2,$export_status.Results.IndexOf("; Scenario:")-$token-2)
                    }
                } else {
                    if((Read-Host -Prompt "Export completed. Show SAS token? [Y/N]").toLower() -eq "y") {
                        SCC-Get-ExportStatus -exportName $exportName -showToken
                        return
                    }
                }
            } elseif(!$export_status.Results.Contains("; Started sources:")) {
                Write-Host "Export has not started yet.  Will check in $($sleepTime) seconds"
                $loops++
                if($loops/10 -eq 1) {
                    $loops = 0
                    $sleepTime = Read-host -Prompt "It's been $([int]$sleepTime*[int]10) seconds.  How often would you like to update results? (default = 5 seconds or 'e' for exit)"
                    if($sleepTime.length -eq 0) {  Write-Host "0 seconds is unacceptable, 5 seconds will be used."; $sleepTime = 5  }
                    elseif($sleepTime -eq 'e') { Write-Host "Exiting."; return }
                    Write-Verbose "Changed updates to every $($sleepTime) seconds."
                }
                Start-Sleep -s $sleepTime
            } 
            else {
                #Export is not complete; report progress; sleep.
                $export_status.Results.Substring($export_status.Results.IndexOf("; Progress: ")+2,18)
                $export_status.Results.Substring($export_status.Results.IndexOf("; Duration: ")+2,28)
                
                Write-Verbose $export_status.Results
                Write-Verbose "Loop count: $($loops)"
                
                $loops++
                Start-Sleep -s $sleepTime
                if($loops/10 -eq 1) {
                    $loops = 0
                    $sleepTime = Read-host -Prompt "It's been $([int]$sleepTime*[int]10) seconds.  How often would you like to update results? (default = 5 seconds or 'e' for exit)"
                    if($sleepTime.length -eq 0) {  Write-Host "0 seconds is unacceptable, 5 seconds will be used."; $sleepTime = 5  }
                    elseif($sleepTime -eq 'e') { Write-Host "Exiting."; return }
                    Write-Verbose "Changed updates to every $($sleepTime) seconds."
                }
                $export_status = Get-ComplianceSearchAction -identity $exportName -Details
            }
        }

        #now that export is complete, check explicitly for zero failed sources and errors

        #(Select-String).length will return a positive int if a match and a Zero if no match.
        if(($export_status.results | Select-String -Pattern "Failed sources: 0;").length) {
            Write-Verbose "No failed source(s) reported."
            $export_contains_fail = $false
        } 
        
        if (!$export_status.Errors.Length){
            Write-Verbose "Error(s) not found."
            $export_contains_err = $false
        }

        #Finally, inform user whether the completed export has failed sources or no failed sources.
        switch($export_contains_fail){
            $true {if($export_contains_err){ Write-Warning "$exportName completed ****with failed sources**** and ****with error(s)****"} else {Write-Warning "$exportName completed with ****failed sources****"}}
            $false {if($export_contains_err) { Write-Warning "$exportname completed with zero failed sources, but ****error(s) may exist*****" } else {Write-Host "$exportName completed with zero failed sources and zero errors."}}
        }
    }
    Write-Verbose $export_status.Results
} #end_of SCC-Get-ExportStatus


<################################
 # "Helpful" O365 S&CC functions.
 ################################>

function SCC-Get-Functions { [CmdletBinding()] Param()

    $function_list = 'closePS','Close all PSSessions (regardless of where it is connected to)',
                     'resetPS','Close all PSSessions and create a PSSession to S&CC',
                     'checkSCC','Checks for S&CC PSSession, if none exist will create one.',
                     'SCC-Get-Variables','Get available variables beginning with "SCC__".',
                     'SCC-New-Case','Create new S&CC eDiscovery Case',
                     'SCC-New-Search','Create a new search within an eDiscovery Case',
                     'SCC-Get-Functions','Get SCC function names and displays brief purpose (not Microsoft cmdlets)',
                     'SCC-Get-ExportStatus','Get the status of an Export request',
                     'SCC-Start-Export','Start an Export of Search Results',
                     'SCC-Get-SearchStatus','Get status of search(es)',
                     'SCC-Get-CustodianLocations','Get source locations given an array of sources (EXO, OD4B, SPO)',
                     'SCC-Get-SearchQuery','Get search queries and locations given an array of search queries (EXO, OD4B, SPO, or ALL)',
                     'SCC-Get-TargetedSearch','Get details for targeted search (EXO, OD4B, SPO)'
    
    $functions = @()
    $i=0
    while($i -lt ($function_list.count)+1) {
        $t_obj = New-Object PSObject
        $t_obj | Add-Member -MemberType NoteProperty -Name 'Function Name' -Value $function_list[$i]
        $t_obj | Add-Member -MemberType NoteProperty -Name Purpose -Value $function_list[$i+1]

        $functions += $t_obj
        $i+=2
    }
    $functions | Sort-Object 'Function Name' | ft -AutoSize

    Write-Host -ForegroundColor Black -BackgroundColor Green 'Tip: Use "Get-Help <Function Name>" for more details.'

} #end_of SCC-Get-Functions

function closePS {  [CmdletBinding()] Param()
    Get-PSSession | Remove-PSSession
    Write-verbose "All PSSession(s) closed"
} #end_of closePS

function resetPS{  [CmdletBinding()] Param()
    closePS
    checkSCC
} #end_of resetPS

function checkSCC{  [CmdletBinding()] Param()
<#
    .Synopsis
      This function returns available sessions to MS O365 Security & Compliance Center sessions or creates a new session.
    .Description
      Function will list any available EXO CC PSSession; if none exist/are available one will be created.

      No parameters are used.
    .Link
      LASTEDIT: 06/12/2017 13:45 PT
#>

    #Check if PSSession exists to Security & Compliance Center
    $Sessions = (Get-PSSession | Where-Object {$_.ComputerName -like "*ps.compliance*" -and $_.Availability -eq "Available"}) | Select-Object Id,Name,ComputerName,State,Availability

    #If PSSession does not exist, create one.
    If ($Sessions -eq $null) {
        Write-Verbose "No Security & Compliance Center PSSession exists.  Connecting to O365 SCC."
        start-sleep -s 3
        
        $Cred = Get-Credential
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection
	    
        # Import PSSession as global modules so other scripts/module may use the cmdlets.
        # https://sites.dartmouth.edu/aig/2013/04/26/custom-powershell-module-easy-remoting-for-office-365/
        Import-Module (Import-PSSession $Session -AllowClobber) -Global

        #If you do not want to import global, use below instead
        #Import-PSSession $Session -AllowClobber

        #Per https://technet.microsoft.com/en-us/library/mt587092(v=exchg.160).aspx , "How do you know this worked?"
        # Microsoft: What does this cmdlet actually do?
        Install-UnifiedCompliancePrerequisite
    }

    #If PSSession session exists, do nothing.
    Else {
        foreach ($sess in $Sessions) {
            Write-Verbose "$($sess.Name) ( ID $($sess.Id) ) is $($sess.State) to $($sess.ComputerName) and is $($sess.Availability)"
        }
    }
} #end_of checkSCC

function SCC-Get-TargetedSearch { 
        [CmdletBinding()]
        Param(
            [Parameter(mandatory=$true, position=0)]
            [ValidateNotNullOrEmpty()]
            [string]$exportName

            ,[Parameter(mandatory=$false, position=1)]
            [ValidateNotNullOrEmpty()]
            [switch]$showToken
        )

    checkSCC
#########################################################################################################
# original source: https://support.office.com/en-us/article/Use-Content-Search-in-Office-365-for-targeted-collections-e3cbc79c-5e97-43d3-8371-9fbc398cd92e
# last obtained from source: 2018-03-19
#
# Modified by:
# -- Removing prompt for credentials,
#
#
# This PowerShell script will prompt you for:								#
#    * Admin credentials for a user who can run the Get-MailboxFolderStatistics cmdlet in Exchange	#
#      Online and who is an eDiscovery Manager in the Security & Compliance Center.			#
# The script will then:											#
#    * If an email address is supplied: list the folders for the target mailbox.			#
#    * If a SharePoint or OneDrive for Business site is supplied: list the folder paths for the site.	#
#    * In both cases, the script supplies the correct search properties (folderid: or path:)		#
#      appeneded to the folder ID or path ID to use in a Content Search.				#
# Notes:												#
#    * For SharePoint and OneDrive for Business, the paths are searched recursively; this means the 	#
#      the current folder and all sub-folders are searched.						#
#    * For Exchange, only the specified folder will be searched; this means sub-folders in the folder	#
#      will not be searched.  To search sub-folders, you need to use the specify the folder ID for	#
#      each sub-folder that you want to search.								#
#    * For Exchange, only folders in the user's primary mailbox will be returned by the script.		#
#########################################################################################################

# Collect the target email address or SharePoint Url
$addressOrSite = Read-Host "Enter an email address or a URL for a SharePoint or OneDrive for Business site"

# Authenticate with Exchange Online and the Security & Complaince Center (Exchange Online Protection - EOP)
if (!$credentials)
{
    $credentials = Get-Credential
}

if ($addressOrSite.IndexOf("@") -ige 0)
{
    # List the folder Ids for the target mailbox
    $emailAddress = $addressOrSite

    # Authenticate with Exchange Online
    if (!$ExoSession)
    {
        $ExoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell-liveid/ -Credential $credentials -Authentication Basic -AllowRedirection
        Import-PSSession $ExoSession -AllowClobber -DisableNameChecking
    }

    $folderQueries = @()
    $folderStatistics = Get-MailboxFolderStatistics $emailAddress
    foreach ($folderStatistic in $folderStatistics)
    {
        $folderId = $folderStatistic.FolderId;
        $folderPath = $folderStatistic.FolderPath;

        $encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
        $nibbler= $encoding.GetBytes("0123456789ABCDEF");
        $folderIdBytes = [Convert]::FromBase64String($folderId);
        $indexIdBytes = New-Object byte[] 48;
        $indexIdIdx=0;
        $folderIdBytes | select -skip 23 -First 24 | %{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
        $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))";

        $folderStat = New-Object PSObject
        Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderPath -Value $folderPath
        Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderQuery -Value $folderQuery

        $folderQueries += $folderStat
    }
    Write-Host "-----Exchange Folders-----"
    $folderQueries |ft
}
elseif ($addressOrSite.IndexOf("http") -ige 0)
{
    $searchName = "SPFoldersSearch"
    $searchActionName = "SPFoldersSearch_Preview"

    # List the folders for the SharePoint or OneDrive for Business Site
    $siteUrl = $addressOrSite

    # Authenticate with the Security & Complaince Center
    if (!$SccSession)
    {
        $SccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $credentials -Authentication Basic -AllowRedirection
        Import-PSSession $SccSession -AllowClobber -DisableNameChecking
    }

    # Clean-up, if the the script was aborted, the search we created might not have been deleted.  Try to do so now.
    Remove-ComplianceSearch $searchName -Confirm:$false -ErrorAction 'SilentlyContinue'

    # Create a Content Search against the SharePoint Site or OneDrive for Business site and only search for folders; wait for the search to complete
    $complianceSearch = New-ComplianceSearch -Name $searchName -ContentMatchQuery "contenttype:folder" -SharePointLocation $siteUrl
    Start-ComplianceSearch $searchName
    do{
        Write-host "Waiting for search to complete..."
        Start-Sleep -s 5
        $complianceSearch = Get-ComplianceSearch $searchName
    }while ($complianceSearch.Status -ne 'Completed')


    if ($complianceSearch.Items -gt 0)
    {
        # Create a Complinace Search Action and wait for it to complete. The folders will be listed in the .Results parameter
        $complianceSearchAction = New-ComplianceSearchAction -SearchName $searchName -Preview
        do
        {
            Write-host "Waiting for search action to complete..."
            Start-Sleep -s 5
            $complianceSearchAction = Get-ComplianceSearchAction $searchActionName
        }while ($complianceSearchAction.Status -ne 'Completed')

        # Get the results and print out the folders
        $results = $complianceSearchAction.Results
        $matches = Select-String "Data Link:.+[,}]" -Input $results -AllMatches
        foreach ($match in $matches.Matches)
        {
            $rawUrl = $match.Value
            $rawUrl = $rawUrl -replace "Data Link: " -replace "," -replace "}"
            Write-Host "path:""$rawUrl"""
        }
    }
    else
    {
        Write-Host "No folders were found for $siteUrl"
    }

    Remove-ComplianceSearch $searchName -Confirm:$false -ErrorAction 'SilentlyContinue'
}
else
{
    Write-Error "Couldn't recognize $addressOrSite as an email address or a site URL"
}


}

#end_of "Helpful" O365 S&CC functions
#end_of_file PS_Automation_Example_r3-public.psm1