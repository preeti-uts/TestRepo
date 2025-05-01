[string]$sAccessKey = $args[0]
[string]$sSecretKey = $args[1]
[string]$sAccountId = $args[2]
[string]$sProjectKey = $args[3]
[string]$sVersionName = $args[4]
[string]$sCycleName = $args[5]
[string]$sFolderName = $args[6]

Write-Host '1-' $sProjectKey
Write-Host '2-' $sVersionName
Write-Host '3-' $sCycleName

write-host 'Starting file conversion'

$sToday = Get-Date -Format "yyyyMMddHHmm"

if (Test-Path -Path ".\reporting\TestResult.xml")
{
    $xml = Resolve-Path ".\reporting\TestResult.xml"
    $output = Join-Path ($pwd) junit-results.xml
    $xslt = New-Object System.Xml.Xsl.XslCompiledTransform;
    $xslt.Load(".\reporting\nunit3-junit.xslt");
    $xslt.Transform($xml, $output);
    $newxml = Resolve-Path junit-results.xml
    write-host 'Path to new file :' $newxml

    # 
    Write-Host ''
    Write-Host '** Test Zyphyr API to create / push results **'

    $sBaseUri = 'https://prod-vortexapi.zephyr4jiracloud.com/api/v1'
    $sMethod = '/jwt/generate'
    $sUri = $sBaseUri+$sMethod #+$sQuery

    $sBody = '{
    "accessKey":"'+$sAccessKey +'",
    "secretKey":"'+ $sSecretKey +'",
    "accountId":"'+ $sAccountId +'"
    }'
    #Write-Host $sUri
    #Write-Host $sBody

    write-host 'Get Token'
    try {
    
        $sResponse = Invoke-WebRequest -Uri $sUri -Method post -Body $sBody -ContentType application/json
        $StatusCode = $sResponse.StatusCode
        write-host query response : $sResponse
    }
    catch {
        $StatusCode = $_.Exception.Response.StatusCode.value__
        Write-Host 'Response text: ' $_.Exception.Message
    }

    if ($StatusCode -eq 200) {
        write-host 'PASS: jwt/generate: Response Code: ' $StatusCode
        $sJWTToken = $sResponse
    } else {
        write-host 'FAILED: jwt/generate: Response Code: ' $StatusCode
        exit 99
    }

    # Create a new automation task
    $sBaseUri = 'https://prod-vortexapi.zephyr4jiracloud.com/api/v1'
    $sMethod =  '/automation/job/saveAndExecute' #'/automation/job/create' '/automation/job/saveAndExecute'
    $sUri = $sBaseUri+$sMethod

    $sForm = @{
        jobName = 'Job1'
        jobDescription = 'This is a created from a powershell script'
        automationFramework = 'junit'
        projectKey = ''+ $sProjectKey +''
        versionName = ''+ $sVersionName +''
        cycleName = ''+ $sCycleName +''
        createNewCycle = 'false'
        appendDateTimeInCycleName = 'false'
        folderName = ''+ $sFolderName + $sToday +''
        createNewFolder = 'true'
        appendDateTimeInFolderName = 'false'
        assigneeUser = ''+ $sAccountId +''
        file =  Get-Item -Path $newxml
        mandatoryFields = '{"reporter": {"label": "Andrew Welsh","name": "Andrew Welsh","id": "'+ $sAccountId +'"}}'
    }

    $myHeader = New-Object System.Collections.Generic.Dictionary"[String,String]"
    $myHeader.Add("accessKey",$sAccessKey)
    $myHeader.Add("jwt", $sJWTToken)

    #Write-Host $sUri
    #Write-Host $sRequestHeader

    write-host 'Save and execute results file in Zephyr'
    try {
    
        $sResponse = Invoke-WebRequest -Uri $sUri -Headers $myHeader -Method post -Form $sForm -ContentType multipart/form-data
        $StatusCode = $sResponse.StatusCode
        write-host query response : $sResponse
    }
    catch {
        $StatusCode = $_.Exception.Response.StatusCode.value__
        Write-Host 'Response text: ' $_.Exception.Message
    }

    if ($StatusCode -eq 200) {
        write-host 'PASS: /automation/job/saveAndExecute: Response Code: ' $StatusCode
    } else {
        write-host 'FAILED: /automation/job/saveAndExecute: Response Code: ' $StatusCode
        exit 99
    }
}
exit 0