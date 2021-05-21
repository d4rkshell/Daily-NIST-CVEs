# Force TLS1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Get RSS Feed for NIST CVE's
$NIST = Invoke-WebRequest -Uri "https://nvd.nist.gov/feeds/xml/cve/misc/nvd-rss.xml" -UseBasicParsing -ContentType "application/xml"

If ($NIST.StatusCode -ne "200") {
    # Feed failed to respond.
    Write-Host "Message: $($NIST.StatusCode) $($NIST.StatusDescription)"
}

# Set feed content
$NISTFeedXml = [xml]$NIST.Content
$Now = Get-Date

# Extract NIST CVE's updated within the last 24 hours
$NISTCVE = @()
$items = $NISTFeedXml.RDF.item

ForEach ($item in $items) {
    If (($Now - [datetime]$item.date).TotalMinutes -le 1440) {
        $title = $item.title
        $link = $item.link
        $desc = $item.description
        $updated = [datetime]$item.date
        $source = $NISTFeedXml.RDF.channel.title

        $NISTvulns = new-object PSCustomObject -prop @{title=$title;description=$desc;link=$link;updated=$updated}
        $NISTCVE += $NISTvulns
    }
}
$NISTCVE | Select-Object -Property Title, Description, Link, Updated #| Out-GridView

# CSS 2.0 styling
$css = @'
<style>
table {
  font-family: arial, sans-serif;
  border-collapse: collapse;
  table-layout:fixed; 
  width: auto;
}

td, th {
  border: 1px solid #dddddd;
  text-align: left;
  padding: 8px;
}

tr td:first-child {
  width: 120px;
}

tr:nth-child(even) {
  background-color: #dddddd;
}
</style>
'@

$emailBody = $NISTCVE | Sort-Object updated -Descending | ConvertTo-Html -Head $css

# Replace some dodgy encodings
$emailBody = $emailBody -replace '&quot;', '"'
$emailBody = $emailBody -replace "&gt;", ">"
$emailBody = $emailBody -replace "&lt;", "<"
$emailBody = $emailBody -replace "&amp;", "&"

# Send email
$emailParameters = @{
    To         = "email-address-here"
    From       = "email-address-here"
    Subject    = "NIST CVE updates: " + $(Get-Date -UFormat "%A %d %B %Y")
    Body       = "<span style='font-family:Calibri;font-size:11pt'>" + "Hello xxxxxxx" + "<br><br>" + "Please find below NIST CVE updates from the past 24 hours: " + "<br><br>" + $emailBody + "</span>" | Out-String
    BodyAsHtml = $true
    SmtpServer = "smtp-server-here"
}
Send-MailMessage @emailParameters