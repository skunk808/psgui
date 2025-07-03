Add-Type -AssemblyName "Microsoft.Data.SqlClient"
Add-Type -AssemblyName "System.Net.Http"
Add-Type -AssemblyName "System.Windows.Forms"

$settings = [PSCustomObject]@{
    PDF_Folder     = "C:\path\to\pdf"
    Draft_Email_Folder = "C:\path\to\sentemails"
    SQL_DB_Name    = "MyDatabase"
    SQL_Server_Name= "MyServer"
    Response_Due_Date = (Get-Date).AddDays(7)
}
$carbonCopyAppAdmins = "debwong@fhb.com"

function Get-Data {
    param(
        [string]$Query,
        [string]$ConnectionString
    )
    $conn = [Microsoft.Data.SqlClient.SqlConnection] $ConnectionString
    $conn.Open()
    $adapter = New-Object Microsoft.Data.SqlClient.SqlDataAdapter($Query, $conn)
    $ds = New-Object System.Data.DataSet
    $adapter.Fill($ds) | Out-Null
    $conn.Close()
    return $ds.Tables[0]
}

function Export-PDFs {
    param(
        [System.Data.DataTable]$Rows
    )
    $baseURI = "http://itd-rpt-dossrst/ReportServer?/Data%20Ops/Oracle%20Database%20User%20Account%20Reval%20-%20By%20Users%20Manager"

    $http = [System.Net.Http.HttpClient]::new((New-Object System.Net.Http.HttpClientHandler -Property @{ UseDefaultCredentials=$true }))
    foreach ($row in $Rows | Where-Object { $_.IsSelected }) {
        if ($row.Attachment -and $row.ManagerEmpNr) {
            $parm = "&HostName_Filter_Parameter=%%&Host_Name_Parameter=$($row.HostNm)&PARM_ManagerEmpNr=$($row.ManagerEmpNr)&rs:Command=Render&rs:Format=PDF"
            $uri = $baseURI + $parm
            $destination = Join-Path $settings.PDF_Folder $row.Attachment
            $response = $http.GetAsync($uri).Result
            $response.EnsureSuccessStatusCode()
            $stream = $response.Content.ReadAsByteArrayAsync().Result
            [IO.File]::WriteAllBytes($destination, $stream)
            $row.Status = "Report Exported"
            Start-Sleep -Milliseconds 500
        }
    }
    [System.Windows.Forms.MessageBox]::Show("Exports Completed","Completed")
}

function Draft-Emails {
    param(
        [System.Data.DataTable]$Rows
    )
    $outlook = New-Object -ComObject Outlook.Application
    $template = Join-Path $settings.PDF_Folder "EmailTemplate.oft"
    $priorGroup = $null
    $msg = $null

    foreach ($row in $Rows | Where-Object { $_.IsSelected }) {
        $currentGroup = $row.GroupName
        if ($currentGroup -ne $priorGroup) {
            if ($msg) { $msg.SaveAs((Join-Path $settings.Draft_Email_Folder ($priorGroup + ".msg"))) }
            $msg = $outlook.CreateItemFromTemplate($template)
            $dueDateStr = $settings.Response_Due_Date.ToLongDateString()
            $msg.Body = $msg.Body.Replace("##DUE_DATE", $dueDateStr)
            $msg.CC = if ($row.CC) { "$carbonCopyAppAdmins;$($row.CC)" } else { $carbonCopyAppAdmins }
            $msg.Recipients.Add($row.ManagerEmail) | Out-Null
            $priorGroup = $currentGroup
        }

        $attachmentPath = Join-Path $settings.PDF_Folder $row.Attachment
        $msg.Attachments.Add($attachmentPath) | Out-Null
        $msg.Save()
        $row.Status = "Notification Drafted"
    }

    if ($msg) { $msg.SaveAs((Join-Path $settings.Draft_Email_Folder ($priorGroup + ".msg"))) }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    [System.Windows.Forms.MessageBox]::Show("Check your Outlook drafts folder","Completed")
}

function Save-Rows {
    param(
        [System.Data.DataTable]$Data,
        [string]$ConnectionString
    )
    $conn = [Microsoft.Data.SqlClient.SqlConnection] $ConnectionString
    $conn.Open()
    $adapter = New-Object Microsoft.Data.SqlClient.SqlDataAdapter("SELECT ...", $conn)
    $builder = New-Object Microsoft.Data.SqlClient.SqlCommandBuilder($adapter)
    $adapter.UpdateCommand = $builder.GetUpdateCommand()
    $adapter.Update($Data) | Out-Null
    $conn.Close()
    # Save settings (optionally serialize $settings)
}

### Main ###
$connString = (New-Object Microsoft.Data.SqlClient.SqlConnectionStringBuilder -Property @{
    InitialCatalog = $settings.SQL_DB_Name
    DataSource      = $settings.SQL_Server_Name
    IntegratedSecurity = $true
    ConnectTimeout = 3
    Encrypt = $false
}).ToString()

$query = @"
SELECT GroupName, Status, IsSelected, HostNm, EnvCd, ManagerFullNm,
       Attachment, ManagerEmpNr, CC, ManagerEmail, RevRptID
FROM dbo.TReviewRptsGrouper
ORDER BY GroupName, HostNm, EnvCd
"@

$table = Get-Data -Query $query -ConnectionString $connString

Export-PDFs -Rows $table
Draft-Emails   -Rows $table
Save-Rows      -Data $table -ConnectionString $connString
