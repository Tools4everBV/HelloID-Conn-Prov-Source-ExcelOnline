$c = $configuration | ConvertFrom-Json;
$body = @{
    "client_id"=$c.client_id
    "scope"="https://graph.microsoft.com/.default"
    "client_secret"=$c.client_secret
    "grant_type"="client_credentials"
}
$connected = $false
try {
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

	$connected = $true
}
catch {
    Throw "Could not get Excel data, message: $($_.Exception.Message)"	
}

if ($connected)
{
	foreach ($employee in $rowsquery.value.values)
	{
		$person  = @{};
		$person['ExternalId'] = $employee[0]
		$person['DisplayName'] = $employee[1] + " " + $employee[2]
		$person['LastName'] = $employee[1]
		$person['FirstName'] = $employee[2]
		
		$person['Contracts'] = [System.Collections.ArrayList]@();
		$contract = @{};
		$contract['SequenceNumber'] = "1";
		$contract['Department'] = $employee[3]
		$contract['StartDate'] = Get-Date([datetime]::FromOADate($employee[4])) -format 'o'
		$contract['EndDate'] = Get-Date([datetime]::FromOADate($employee[5])) -format 'o'
		[void]$person['Contracts'].Add($contract);
		Write-Output ($person | ConvertTo-Json -Depth 20);
	}
}