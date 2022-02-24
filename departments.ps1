$c = $configuration | ConvertFrom-Json;
$body = @{
    "client_id"=$c.client_id
    "scope"="https://graph.microsoft.com/.default"
    "client_secret"=$c.client_secret
    "grant_type"="client_credentials"
}
$tokenquery = Invoke-RestMethod -uri https://login.microsoftonline.com/$($c.tenant_id)/oauth2/v2.0/token -body $body -Method Post -ContentType 'application/x-www-form-urlencoded'

$headers = @{
    "content-type" = "Application/Json"
    "authorization" = "Bearer $($tokenquery.access_token)"
}

$document = Invoke-RestMethod -uri https://graph.microsoft.com/v1.0/users/$($c.user_id)/drive/root:/$($c.document_path) -Method GET -Headers $headers

$bodysession = @{
    "persistChanges" = $false
}
$url = "https://graph.microsoft.com/v1.0/users/$($c.user_id)/drive/items/$($document.id)/workbook/createSession"
$documentsession = Invoke-RestMethod -uri $url -Method POST -Body ($bodysession | ConvertTo-Json) -Headers $headers

$headersworkbook = @{
    "workbook-session-id" = $documentsession.id
    "authorization" = "Bearer $($tokenquery.access_token)"
}

$url = "https://graph.microsoft.com/v1.0/users/$($c.user_id)/drive/items/$($document.id)/workbook/worksheets/$($c.table_name)/tables/$($c.table_name)/rows"
$rowsquery = Invoke-RestMethod -uri $url -Method GET -Headers $headersworkbook 

$departments  = [System.Collections.ArrayList]@();
foreach ($employee in $rowsquery.value.values)
{
    $department  = @{};
    $department['Name'] = $employee[3]
    $department['DisplayName'] = $employee[3]
    $department['ExternalId'] = $employee[3]
    if ([string]::IsNullOrEmpty($department['ExternalId']) -eq $true)
    {
        $department['ExternalId'] = $department['Name']
    }
    if ($departments.Contains($department['ExternalId']) -eq $false)
    {
        Write-Output ($department | ConvertTo-Json -Depth 20);
        $departments += $department['ExternalId'];
    }
}