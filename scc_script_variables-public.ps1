# Purpose: Introduce case-specific variables for use with PS_Automation_Example_r3-public.ps1

# It is recommended to adhere to $SCC__ naming convention of any variables you may use for use with the function: SCC-Get-Variables

# EXAMPLE variables include the following:

$SCC__caseName = 'SCC_eDiscovery_Case-Test'
$SCC__sourceData =  'user1@company.com',
                    'user2@company.com',
                    'https://company-my.sharepoint.com/personal/user1/',
                    'https://company-my.sharepoint.com/personal/user2/',
                    'https://company.sharepoint.com/teams/SiteName1/',
                    'https://company.sharepoint.com/teams/SiteName2/Sub-siteName/'

# Two examples of sourceQuery:
#  - "[ALL] " indicates that the search query is to be applied to all sources of data (EXO, SPO, OD4B)
$SCC__sourceQuery = '[ALL] size>1000'

#  - Second example indicates specific search queries are to be applied to only certain sources of data (EXO, SPO, *or* OD4B)
$SCC__sourceQuery = "[EXO] sent>06/01/2017 AND keyword1","[SPO] created>06/01/2017","[OD4B] fileextension:docx"

# Do not use [ALL] with any other queries.
# The use of [ALL] will *always* take precedence over any source-specific query.
# Only one query may be used per type of source data.
#
# The following are two examples that will not produce intended results.
# - This query mixes an [ALL] with another source query, [EXO]:
# - - $badQuery = "[EXO] sent>06/01/2017 AND keyword1","[ALL] size>1000"
#
# - This query has two search queries for the same source type, [EXO]:
# - - $badQuery = "[EXO] sent>06/01/2017 AND keyword1","[EXO] size>1000"

$SCC__searchNameOD4B = $SCC__caseName + '-OD4B'
$SCC__searchNameEXO = $SCC__caseName + '-EXO'
$SCC__searchNameSPO = $SCC__caseName + '-SPO'

$SCC__exportEXO = $SCC__searchNameEXO + "_Export"
$SCC__exportOD4B = $SCC__searchNameOD4B + "_Export"
$SCC__exportSPO = $SCC__searchNameSPO + "_Export"

<# Use Date as a Unique ID (of sorts) for search names

# This temp date may not be necessary in your workflow.
# Temp date is used as a unique ID of sorts
$SCC__tempDate = Get-Date -Format yyyymmdd_hhmmss
$SCC__holdName = $SCC__CaseName+'-Hold-'+$SCC__tempDate

$SCC__searchNameOD4B = $SCC__caseName + '-OD4B-' + $SCC__tempDate
$SCC__searchNameEXO = $SCC__caseName + '-EXO-' + $SCC__tempDate
$SCC__searchNameSPO = $SCC__caseName + '-SPO-' + $SCC__tempDate

#>