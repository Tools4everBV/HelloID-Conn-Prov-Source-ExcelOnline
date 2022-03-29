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

    if ($c.useSharepointOnline)
	{
		$sites = Invoke-RestMethod -uri https://graph.microsoft.com/v1.0/sites -Method GET -Headers $headers
		$filter_sites = $sites.value | where-object { $_.name -eq $c.site_name}
		if (($filter_sites | Measure-Object).Count -eq 1)
		{
			if ([string]::IsNullOrEmpty($c.list_name))
			{
				#no list
				$document = Invoke-RestMethod -uri "https://graph.microsoft.com/v1.0/sites/$($filter_sites.id)/drive/root:/$($c.document_path)" -Method GET -Headers $headers
				if (($document | Measure-Object).Count -eq 0)
				{
					throw ("No Document found")
				}

				$bodysession = @{
					"persistChanges" = $false
				}
				$url = "https://graph.microsoft.com/v1.0/sites/$($filter_sites.id)/drive/items/$($document.id)/workbook/createSession"
				$documentsession = Invoke-RestMethod -uri $url -Method POST -Body ($bodysession | ConvertTo-Json) -Headers $headers

				$headersworkbook = @{
					"workbook-session-id" = $documentsession.id
					"authorization" = "Bearer $($tokenquery.access_token)"
				}
				
				$url = "https://graph.microsoft.com/v1.0/sites/$($filter_sites.id)/drive/items/$($document.id)/workbook/worksheets/$($c.table_name)/tables"
				$tablesquery = Invoke-RestMethod -uri $url -Method GET -Headers $headers
				if (($tablesquery | Measure-Object).Count -eq 0)
				{
					throw ("No Tables found")
				}
				if (($tablesquery | Measure-Object).Count -eq 1)
				{
					$sub_table_name = $tablesquery.value.name
				}
				if (($tablesquery | Measure-Object).Count -ge 2)
				{
					$filter_tables = $tablesquery.value | where-object {$_.name -match $c.table_name}
					if (($filter_tables | Measure-Object).Count -eq 1)
					{
						$sub_table_name = $filter_tables.value.name
					}
					else 
					{
						throw ("No Sub-Table found - tables found: "+($tablesquery.value | Select name | ConvertTo-Json))
					}
				}
				$url = "https://graph.microsoft.com/v1.0/sites/$($filter_sites.id)/drive/items/$($document.id)/workbook/worksheets/$($c.table_name)/tables/$sub_table_name/rows"
			}
			else
			{
				#use list
				$lists = Invoke-RestMethod -uri https://graph.microsoft.com/v1.0/sites/$($filter_sites.id)/lists -Method GET -Headers $headers
				$filter_lists = $lists | where-object { $_.name -eq $c.list_name}
				if (($filter_lists | Measure-Object).Count -eq 1)
				{
					$document = Invoke-RestMethod -uri "https://graph.microsoft.com/v1.0/sites/$($filter_sites.id)/lists/$($filter_lists.id)/drive/root:/$($c.document_path)" -Method GET -Headers $headers
					if (($document | Measure-Object).Count -eq 0)
					{
						throw ("No Document found")
					}

					$bodysession = @{
						"persistChanges" = $false
					}
					$url = "https://graph.microsoft.com/v1.0/sites/$($filter_sites.id)/lists/$($filter_lists.id)/drive/items/$($document.id)/workbook/createSession"
					$documentsession = Invoke-RestMethod -uri $url -Method POST -Body ($bodysession | ConvertTo-Json) -Headers $headers

					$headersworkbook = @{
						"workbook-session-id" = $documentsession.id
						"authorization" = "Bearer $($tokenquery.access_token)"
					}
					
					$url = "https://graph.microsoft.com/v1.0/sites/$($filter_sites.id)/lists/$($filter_lists.id)/drive/items/$($document.id)/workbook/worksheets/$($c.table_name)/tables"
					$tablesquery = Invoke-RestMethod -uri $url -Method GET -Headers $headers
					if (($tablesquery | Measure-Object).Count -eq 0)
					{
						throw ("No Tables found")
					}
					if (($tablesquery | Measure-Object).Count -eq 1)
					{
						$sub_table_name = $tablesquery.value.name
					}
					if (($tablesquery | Measure-Object).Count -ge 2)
					{
						$filter_tables = $tablesquery.value | where-object {$_.name -match $c.table_name}
						if (($filter_tables | Measure-Object).Count -eq 1)
						{
							$sub_table_name = $filter_tables.value.name
						}
						if (($filter_tables | Measure-Object).Count -eq 0)
						{
							throw ("No Sub-Table found - tables found: "+($tablesquery.value | Select name | ConvertTo-Json))
						}
					}
					$url = "https://graph.microsoft.com/v1.0/sites/$($filter_sites.id)/lists/$($filter_lists.id)/drive/items/$($document.id)/workbook/worksheets/$($c.table_name)/tables/$sub_table_name/rows"
				}
			}
		}
	}
	else
	{
		$document = Invoke-RestMethod -uri https://graph.microsoft.com/v1.0/users/$($c.user_id)/drive/root:/$($c.document_path) -Method GET -Headers $headers
		if (($document | Measure-Object).Count -eq 0)
		{
			throw ("No Document found")
		}
		$bodysession = @{
			"persistChanges" = $false
		}
		$url = "https://graph.microsoft.com/v1.0/users/$($c.user_id)/drive/items/$($document.id)/workbook/createSession"
		$documentsession = Invoke-RestMethod -uri $url -Method POST -Body ($bodysession | ConvertTo-Json) -Headers $headers

		$headersworkbook = @{
			"workbook-session-id" = $documentsession.id
			"authorization" = "Bearer $($tokenquery.access_token)"
		}

		$url = "https://graph.microsoft.com/v1.0/users/$($c.user_id)/drive/items/$($document.id)/workbook/worksheets/$($c.table_name)/tables"
		$tablesquery = Invoke-RestMethod -uri $url -Method GET -Headers $headersworkbook
		if (($tablesquery | Measure-Object).Count -eq 0)
		{
			throw ("No Tables found")
		}
		if (($tablesquery | Measure-Object).Count -eq 1)
		{
			$sub_table_name = $tablesquery.name
		}
		if (($tablesquery | Measure-Object).Count -ge 2)
		{
			$filter_tables = $tablesquery | where-object {$_.name -match $c.table_name}
			if (($filter_tables | Measure-Object).Count -eq 1)
			{
				$sub_table_name = $filter_tables.value.name
			}
			if (($filter_tables | Measure-Object).Count -eq 0)
			{
				throw ("No Sub-Table found - tables found: "+($tablesquery.value | Select name | ConvertTo-Json))
			}
		}
		$url = "https://graph.microsoft.com/v1.0/users/$($c.user_id)/drive/items/$($document.id)/workbook/worksheets/$($c.table_name)/tables/$sub_table_name/rows"
	}

    $rowsquery = Invoke-RestMethod -uri $url -Method GET -Headers $headersworkbook 

	$connected = $true
}
catch {
    Throw "Could not get Excel data, message: $($_.Exception.Message)"	
}

if ($connected)
{
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
}