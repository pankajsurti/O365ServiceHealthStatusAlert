<#
.SYNOPSIS
Office 365 Service Health Alert
.DESCRIPTION 
Office 365 Service Health is a reporting and alerting solution. Due to its design and functionality, I like to refer it as Solution. It uses Office 365 Service Communications API.

It uses PowerShell, HTML and CSS for conditional formatting. 

We can generate a HTML output report or schedule it to run on periodic intervals with the help of Windows Task Scheduler. 

We can also add small snippet of code at the end to Send Email Alert Notifications. 

Though Office 365 Service Health does not give tenant specific information. This solution can be modified later to incorporate new functionalities that will be rolled out by Microsoft - Like to include user count of an affected tenant.

The Solution uses JSON Config file to load the configuration like Application ID, Client Secret, AAD Instance, TenantDomain and Log file path. 

.OUTPUTS
HTML Output file or we can use Send-MailMessage to send an email.
.EXAMPLE
.\O365ServiceHealthAlert.ps1
This will run and generate an Output in HTML 
.NOTES

*** License ***
The MIT License (MIT)
Copyright (c) 2018 Sirbuland Khan
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>


# Load configuration file
$configObj = Get-Content ".\appsettings.json" | ConvertFrom-Json

# Initialize variables with config data
$clientID = $configObj.ClientId
$clientSecret = $configObj.ClientSecret
$loginURL = $configObj.AADInstance
$tenantDomain = $configObj.TenantDomain
$resource = $configObj.ResourceURI
$logFile = $configObj.LogFilePath
$currentDateTime = Get-Date
$daysToAdd = -2
$oldDateTime = $currentDateTime.AddDays($daysToAdd)

$htmlHeader = "
<!DOCTYPE html>
<html lang=""en"">
<head>
<meta name=""viewport"" content=""width=device-width, intial-scale=1"">
<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />
            <style type=""text/css"">
                body {
                    width: 75%;
                    padding: 40px;
                    font-family: Arial;
                    font-size: 12pt;
                }
                h1{
                    font-size: 26pt;
                    color: #323233;
                    text-transform: uppercase;
                    text-align: center;
                }
                table {
                    border: 1px solid #000000;
                    border-collapse: collapse;
                    table-layout: fixed;
                    font-size: 10pt;
                    margin-bottom: 20px;
                }
                th, td {
                    border: 1px solid #000000;
                    text-align: left;
                    padding: 5px;
                }
                th {
                    width: 150px;
                    background-color: #000000;
                    color: #ffffff;
                    border-bottom: 1px solid #ffffff;
                }
                td {
                    width: 720px;
                }
                th.degraded {
                    background-color: #ff3600;
                    color: #ffffff;
                }
                th.restoring {
                    background-color: #ffee5e;
                    color: #000000;             
                }
                th.restored {
                    background-color: #45c405;
                    color: #ffffff;
                }
                td.incident {
                    background-color: #d1431c;
                    color: #ffffff;
                }
                td.advisory {
                    background-color: #1d92d1;
                    color: #ffffff;
                }
            </style>
            </head>
            <body>
            <p>
                    <h1><span>Office 365 Services Health Status</span></h1>
            </p>
"



# Function to write log file 
function Write-Log {

    [CmdletBinding()]
    
    Param (
    
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO", "WARNING", "ERROR", "CRITICAL")]
    $Level = "INFO",
    
    [Parameter(Mandatory=$True)]
    $Message = $null,
    
    [Parameter(Mandatory=$False)]
    $FilePath = $logFile + "\$((Get-Date -Format "yyyyMM").Replace(':', '')).log" 
    
    )
    
    if (-not(Test-Path $FilePath))
    {
         New-Item $FilePath -Type file

    } else {
    
        $DateTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $InputLine = "$DateTime $Level $Message"
        Add-Content -Path $FilePath -Value $InputLine

    }
    
}

# String manipulation function because the Message Text in response from the API will have TiTle, User Impact, Current Status, Scope of Impact
# and Next Update will all be in a single string. We want to separate the text and show in specific table rows.
function Get-ExtricatedText {
    param (
        [string]$inputMessageText = "" , # This parameter will require case sensitive string
        [string]$startString = "",
        [string]$endString = ""
    )

    if($inputMessageText -ne $null) {
        
    try {

        $startIndex = $inputMessageText.IndexOf($startString, 0)
        $endIndex = $inputMessageText.IndexOf($endString, 0)
        $textLength = $endIndex - $startIndex
        $extricatedText = $inputMessageText.Substring($startIndex, $textLength)

    } catch {

        Write-Log -Level ERROR -Message $_.Exception
    }
    
    }
    return $extricatedText
}


#Function to get access token using client credentials
function Get-AccessToken(){

    $body = @{grant_type="client_credentials";resource=$resource;client_id=$clientID;client_secret=$clientSecret}
    try {

        $oauth = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantDomain/oauth2/token?api-version=1.0 -Body $body
    }
    catch {
        
        Write-Log -Level ERROR -Message $_.Exception

    }
    
    $headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}

    Write-Log -Level INFO -Message "Authentication Successfull"
    return $headerParams

}

#Function to get service health data using rest api
function Get-ServiceHealthStatus (){
            $headers = Get-AccessToken
    [string]$url = "https://manage.office.com/api/v1.0/$tenantDomain/ServiceComms/Messages"
    
    Write-Log -Level INFO -Message "Fetching Service Health Messages from Office 365 Service Communications API"

    try {

        $serviceHealthMessages = Invoke-RestMethod -Method GET -Headers $headers -Uri $url 
    }
    catch {
        
        Write-Log -Level ERROR -Message $_.Exception

    }
    
    Write-Log -Level INFO -Message "Result Fetch Successful!!"
    $healthMessages = ($serviceHealthMessages.value)
    return $healthMessages
}

try {
    
    $degradedServices = Get-ServiceHealthStatus | ?{$_.Status -eq "Service Degradation"}
}
catch {

    Write-Log -Level ERROR -Message $_.Exception
}

if($degradedServices -ne $null) {

$degradationServicesTable = "  
                    <p>
                    <table cellpadding=""0"">
                    <tr>
                    <th class=""degraded"" colspan=""2"" style=""font-size: 20pt;text-align: center;border-bottom: 1px solid #000000;"">Services Degraded</th>
                    </tr>
"

foreach($dgdService in $degradedServices){
    [string]$msg = $dgdService.Messages[-1]
    [datetime]$startTime = $dgdService.StartTime
    [datetime]$updatedTime = $dgdService.LastUpdatedTime

    $currentStatus = Get-ExtricatedText -inputMessageText $msg -startString "Current status"  -endString "Scope of impact"

    $degradationServicesTable += "
                    <tr>
                    <th scope=""row"">Id</th>
                    <td style=""background-color: #000000;color: #ffffff;"">$($dgdService.Id)</td>
                    </tr>
                    <tr>
                    <th scope=""row"">Service</th>
                    <td>$($dgdService.WorkloadDisplayName)</td>
                    </tr>
                    <tr>
                    <th scope=""row"">Feature</th>
                    <td>$($dgdService.Feature)</td>
                    </tr>
                    <tr>
                    <th scope=""row"">Impact Description</th>
                    <td>$($dgdService.ImpactDescription)</td>
                    </tr>
"
if ($dgdService.Classification -eq "Incident")
{
    $degradationServicesTable += "
                    <tr>
                    <th scope=""row"">Classification</th>
                    <td class=""incident"">$($dgdService.Classification)</td>
                    </tr>
                    "
} else {

    $degradationServicesTable += "
                    <tr>
                    <th scope=""row"">Classification</th>
                    <td class=""advisory"">$($dgdService.Classification)</td>
                    </tr>
                    "
}

    $degradationServicesTable += "
                    <tr>
                    <th>Status</th>
                    <td>$($dgdService.Status)</td>
                    </tr>
                    <tr>
                    <th>Start Time</th>
                    <td>$($startTime) UTC</td>
                    </tr>
                    <tr>
                    <th>Last Updated Time</th>
                    <td>$($updatedTime) UTC</td>
                    </tr>
                    "
        if($currentStatus -ne $null)
        {
            $degradationServicesTable += "
                                    <tr>
                                    <th>Current Status</th>
                                    <td>$($currentStatus.Trim("Current status:"))</td>
                                    </tr>
                                    "
        } else {
            
            $degradationServicesTable += "
                                    <tr>
                                    <th>Last Update</th>
                                    <td>There has been no recent update since last update.</td>
                                    </tr>
                                    "
        }
    }

    $degradationServicesTable += " 
                            </table>
                            </p>
                            <br><br>
                            "
}

try {
    
    $restoringServices = Get-ServiceHealthStatus | ?{$_.Status -eq "Restoring Service"}
}
catch {
    
    Write-Log -Level ERROR -Message $_.Exception
}


if ($restoringServices -ne $null) {

$restoringServicesTable = "
                    <p>
                    <table>
                    <tr>
                    <th class=""restoring"" colspan=""2"" style=""font-size: 20pt;text-align: center;border: 1px solid #000000;"">Restoring Services</th>
                    </tr>
"

foreach($rtdService in $restoringServices){
    [string]$msg = $rtdService.Messages[-1]
    [datetime]$startTime = $rtdService.StartTime
    [datetime]$updatedTime = $rtdService.LastUpdatedTime

    $currentStatus = Get-ExtricatedText -inputMessageText $msg -startString "Current status" -endString "Scope of impact"

    $restoringServicesTable += "
                    <tr>
                    <th scope=""row"">Id</th>
                    <td style=""background-color: #000000;color: #ffffff;"">$($rtdService.Id)</td>
                    </tr>
                    <tr>
                    <th scope=""row"">Service</th>
                    <td>$($rtdService.WorkloadDisplayName)</td>
                    </tr>
                    <tr>
                    <th scope=""row"">Feature</th>
                    <td>$($rtdService.Feature)</td>
                    </tr>
                    <tr>
                    <th scope=""row"">Impact Description</th>
                    <td>$($rtdService.ImpactDescription)</td>
                    </tr>
"
if ($rtdService.Classification -eq "Incident")
{
    $restoringServicesTable += "
                    <tr>
                    <th scope=""row"">Classification</th>
                    <td class=""incident"">$($rtdService.Classification)</td>
                    </tr>
                    "
} else {

    $restoringServicesTable += "
                    <tr>
                    <th scope=""row"">Classification</th>
                    <td class=""advisory"">$($rtdService.Classification)</td>
                    </tr>
                    "
}

    $restoringServicesTable += "
                    <tr>
                    <th>Status</th>
                    <td>$($rtdService.Status)</td>
                    </tr>
                    <tr>
                    <th>Start Time</th>
                    <td>$($startTime) UTC</td>
                    </tr>
                    <tr>
                    <th>Last Updated Time</th>
                    <td>$($updatedTime) UTC</td>
                    </tr>
                    "
        if ($currentStatus -ne $null)
        {
            $restoringServicesTable += "
                                    <tr>
                                    <th>Current Status</th>
                                    <td>$($currentStatus.Trim("Current status:"))</td>
                                    </tr>
                                    "
        } else {
            
            $restoringServicesTable += "
                                    <tr>
                                    <th>Last Update</th>
                                    <td>There has been no recent update since last update.</td>
                                    </tr>
                                    "
        }

    }

    $restoringServicesTable += " 
                            </table>
                            </p>
                            <br><br>
                            "
}



try {
    
    $restoredServices  = Get-ServiceHealthStatus | ?{$_.EndTime -ne $null}
}
catch {
    
    Write-Log -Level ERROR -Message $_.Exception
}

if ($restoredServices -ne $null) {


    $restoredServiceAll = foreach($rtdService in $restoredServices){
        $properties = [ordered]@{
            Id = $rtdService.Id
            WorkloadDisplayName = $rtdService.WorkloadDisplayName
            Feature = $rtdService.Feature
            ImpactDescription = $rtdService.ImpactDescription
            Status = $rtdService.Status 
            Classification = $rtdService.Classification
            UpdateDetails = ([string]$update =$rtdService.Messages[-1])
            StartTime = ([datetime]$sTime = $rtdService.StartTime)
            EndTime = ([datetime]$eTime = $rtdService.EndTIme)
            LastUpdatedTime = ([datetime]$lTime = $rtdService.LastUpdatedTime)

        }
        New-Object PSObject -Property $properties
    }
   
  $restoredServiceDetails = $restoredServiceAll | ?{$_.EndTime -gt $oldDateTime -and $_.EndTime -le $currentDateTime}

  if ($restoredServiceDetails -ne $null){
    
  $restoredServicesTable = "
                            <p>
                            <table>
                            <tr>
                            <th class=""restored"" colspan=""2"" style=""font-size: 20pt;text-align: center;border-bottom: 1px solid #000000;"">Restored Services</th>
                            </tr>
                            "

  foreach($rtdService in $restoredServiceDetails){
    [string]$msg = $rtdService.UpdateDetails
    [datetime]$startTime = $rtdService.StartTime
    [datetime]$endTime = $rtdService.EndTime
    [datetime]$updatedTime = $rtdService.LastUpdatedTime

    $currentStatus = Get-ExtricatedText -inputMessageText $msg -startString "Current status" -endString "Scope of impact"
    $finalStatus = Get-ExtricatedText -inputMessageText $msg -startString "Final status" -endString "Scope of impact"

    $restoredServicesTable += "
                            <tr>
                            <th scope=""row"">Id</th>
                            <td style=""background-color: #000000;color: #ffffff;"">$($rtdService.Id)</td>
                            </tr>
                            <tr>
                            <th scope=""row"">Service</th>
                            <td>$($rtdService.WorkloadDisplayName)</td>
                            </tr>
                            <tr>
                            <th scope=""row"">Feature</th>
                            <td>$($rtdService.Feature)</td>
                            </tr>
                            <tr>
                            <th scope=""row"">Impact Description</th>
                            <td>$($rtdService.ImpactDescription)</td>
                            </tr>
"
if ($rtdService.Classification -eq "Incident")
{
    $restoredServicesTable += "
                            <tr>
                            <th scope=""row"">Classification</th>
                            <td class=""incident"">$($rtdService.Classification)</td>
                            </tr>
                            "
} else {

    $restoredServicesTable += "
                            <tr>
                            <th scope=""row"">Classification</th>
                            <td class=""advisory"">$($rtdService.Classification)</td>
                            </tr>
                            "
}

    $restoredServicesTable += "
                            <tr>
                            <th>Status</th>
                            <td>$($rtdService.Status)</td>
                            </tr>
                            <tr>
                            <th>Start Time</th>
                            <td>$($startTime) UTC</td>
                            </tr>
                            <tr>
                            <th>Last Updated Time</th>
                            <td>$($updatedTime) UTC</td>
                            </tr>
                            "
     if($currentStatus -ne $null)
     {
        $restoredServicesTable += "
                                <tr>
                                <th>Current Status</th>
                                <td>$($currentStatus.Trim("Current status:"))</td>
                                </tr>
                                "
     } elseif ($finalStatus -ne $null) {
        $restoredServicesTable += "
                                <tr>
                                <th>Final Update</th>
                                <td>$($finalStatus.Trim("Final status:"))</td>
                                </tr>
                                "
     } else {
         
        $restoredServicesTable += "
                                <tr>
                                <th>Final Update</th>
                                <td>There has been no recent update since last update.</td>
                                </tr>
                                "
     }

       $restoredServicesTable += "
                                <tr>
                                <th>End Time</th>
                                <td>$($endTime) UTC</td>
                                </tr>
                                "

    }

    $restoredServicesTable += " 
                            </table>
                            </p>
                            "

  }
}    

$endOfHtmlBody = "
        </body>
        </html>
"

$html = $htmlHeader + $degradationServicesTable + $restoringServicesTable + $restoredServicesTable + $endOfHtmlBody
$html | Out-File .\results.html
