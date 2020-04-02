# This script will require the Web Application and permissions setup in Azure Active Directory
$ClientID       = "d3032a0a-b2f0-42bc-a8a3-a16980e9983b"             # Should be a ~35 character string insert your info here
$ClientSecret   = "@QTB9?2-?ODnPfOnTR9ajDgbyazIwNA6"         # Should be a ~44 character string insert your info here
$loginURL       = "https://login.microsoftonline.com/"
$tenantdomain   = "williamfiddes.onmicrosoft.com"

# Leave blank if you dont want to filter by AppId
$SearchByAppId = ""

# Leave blank if you dont want to filter by User ID
$SearchByUserId = ""

$top            = 1000
$ResultsPerPage = 100000           

$ago = "{0:s}" -f (get-date).AddDays(-30) + "Z"
# or, AddMinutes(-5)


# Get an Oauth 2 access token based on client id, secret and tenant domain
function Get-Token
{
    Param()
    $resource = "https://graph.microsoft.com"
    $body       = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
    
    $AuthResponse = $null
    $AuthResponse = (Invoke-RestMethod -Method Post -Uri "$loginURL/$tenantdomain/oauth2/token" -Body $body -verbose)
    
    $result = $null
    $result = $AuthResponse.access_token
    return $result
}

function Get-Report
{
    Param()

    $AccessToken = Get-Token
    if (Get-Token -ne $null) {

        $url = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$top=$top&`$filter=createdDateTime ge $ago"
        if($SearchByAppId)
        {
            $url += " and appId eq '$SearchByAppId'"
        }
    
        if($SearchByUserId)
        {
            $url += " and userPrincipalName eq '$SearchByUserId'"
        }
    
        $i=0
    
        $Report = @()
    
        $RetryAttempts = 0
        $MaxRetryAttemps = 5
        Do{
            #Write-Output "Fetching data using Uri: $url"
            $response = $null
    
            try {
                $headerParams = $null
                $headerParams = @{'Authorization'="Bearer $($AccessToken)"}
                $response = (Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $url -Verbose)
                
            }
            catch {
                if($_.Exception.Response.StatusCode -eq "Unauthorized")
                {
                    $AccessToken = Get-Token
                    $RetryAttempts++
                    Continue
                }
    
                if($_.Exception.Response.StatusCode -eq "429")
                {
                    Write-Host "Throttling limit hit: StatusCode 429: Waiting 5 minutes." -ForegroundColor Yellow
                    if($RetryAttempts -gt 2)
                    {
                        Start-Sleep -Seconds 300
                    }
                    $RetryAttempts++
                    Continue
                }

                if($RetryAttempts -gt $MaxRetryAttemps)
                {
                    Write-Line "Script failed. Max retry attemps exceeded..." -foregroundcolor Red
                    throw $_
                }
            }

            $RetryAttempts = 0
    
            $content = $null
            $content = $response.Content | ConvertFrom-Json
    
            $nextLink = $content.'@odata.nextLink'
            $url = $nextLink
            $Signins = $content.value
            
            foreach($entry in $Signins)
            {
                $ReportItem = @{}
                foreach($Member in ($entry | Get-Member))
                {
                    if($Member.MemberType -eq "NoteProperty")
                    {
                        $MemberName = $null
                        $MemberName = $Member.Name
                        $Value = $entry.$($MemberName)
                        if($Value)
                        {
                            $Type = $null
                            $Type = $entry.$($MemberName).GetType()
                            if($Type.isArray)
                            {
                                $ReportItem.($MemberName) = $entry.($MemberName) | ConvertTo-Json -Compress -Depth 99
                            }
    
                            else {
                                $ReportItem.($MemberName) = $entry.($MemberName)
                            }
                        }
                    }
                }
                $Report += New-Object -TypeName psobject -Property $ReportItem
    
                if($Report.Count -ge $ResultsPerPage)
                {
                    Write-Output "Saving the output to a file SigninActivities$i.json"
                    $Report | Export-Csv "SigninActivities$i.csv" -Force
                    $Report = @()
                    $i = $i+1
                }
            }
    
    
            # Stop script if there are no other pages
            if(-not $url)
            {
                Write-Output "Saving the output to a file SigninActivities$i.json"
                $Report | Export-Csv "SigninActivities$i.csv" -Force
                break
            }
    
            Start-Sleep -Milliseconds 100
        } while($true)
    
    } else {
    
        Write-Host "ERROR: No Access Token"
    }
}



