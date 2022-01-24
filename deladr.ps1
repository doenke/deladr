<#
  Connect to a CardDAV Address book, search the email addresses for keywords and deletes the matching contacts
  Version 1.0, 15.1.2017

  Delete-Addresses -username "user@mailbox.org" -password "geheim"

  Sönke Simon 2018

You can also put a configuration file into the same folder, where also this script is. The filename must be "deladr.config".
Alternatively  provide a configuration filename with the -config parameter

Parameters:
-username "user"       default: ""                           if this is provided on the command line, then the command line is leading
-password "secret"     default: ""                           if this is provided on the command line, then the command line is leading
-server <user>         default: "https://dav.mailbox.org"    if this is provided on the config file, the config file is leading
-folder <user>         default: "/carddav/29/"               if this is provided on the config file, the config file is leading
-simulate                                                    Deactivates the deletion of the VCards. On the config file 
-config "filename"     default: "deladr.config"              Config filename for the following settings:

username=myuser
password=secret
server=https://myserver.com
folder=/user/folder
#>

param(
[switch]$simulate,
[switch]$verbose,
[string]$username = "",
[string]$password = "",
[string]$config = ($PSScriptRoot+"\deladr.config"),
[string]$server="https://dav.mailbox.org",
[string]$folder="/carddav/29/"
)
$oldverbose = $VerbosePreference
if($verbose) {
    $VerbosePreference = "continue" 
}
Write-Output "Delete Vcards from a CardDAV Addressbook"
$simulation = $simulate.ToBool()
if (Test-Path $config)
{
    Write-Verbose ("Config File " + $config + " found")
    $c = Get-Content $config | Out-String | ConvertFrom-StringData

    if ($username -eq "") {$username = $c["username"]}
    if ($password -eq "") {$password = $c["password"]}
    if ($c["simulate"]) {$simulation=$true}
    if ($c["server"]) {$server=$c["server"]}
    if ($c["folder"]) {$folder=$c["folder"]}
} else
{Write-Verbose ("Config File " + $config + " not found")}


if ($username -eq "" -or $password -eq "" )
{
    Write-Error 'Username or Password not provided'
    exit
}
 
$url = $server+$folder
Write-Verbose ('Delete VCards from Folder ' + $url)
Write-verbose ('User ' + $username)


# Load keywords
$begriffe = Get-Content ($PSScriptRoot+"\keywords.txt")
Write-Verbose ($begriffe.Length.ToString() + ' Keywords loaded')


# prepare user for curl - DELETE request
$user=New-Object System.Management.Automation.PSCredential ($username, (ConvertTo-SecureString $password -AsPlainText -Force))

# prepare PROPFIND request
$r = [System.Net.WebRequest]::Create($url)
$uri = new-object System.Uri ( $url )
$enc = [system.Text.Encoding]::UTF8
$credcache = new-object System.Net.CredentialCache
$credcache.Add($uri, "Basic", $user.GetNetworkCredential())
$r.Credentials = $credcache
$r.method="PROPFIND"
$r.Headers.Add("Depth",1)
$r.ContentType="text/xml"

# prepare body for PROPFIND request
$body = "<C:propfind xmlns:D='DAV:' xmlns:C='urn:ietf:params:xml:ns:carddav'><D:prop><C:address-data/></D:prop></C:propfind>"
$ByteQuery = $enc.GetBytes($body)
$r.ContentLength = $ByteQuery.Length
$QueryStream = $r.GetRequestStream()
$QueryStream.Write($ByteQuery, 0, $ByteQuery.Length)
$QueryStream.Close()

# Run request

$sr = new-object System.IO.StreamReader $r.GetResponse().GetResponseStream()
[xml]$daten = $sr.ReadToEnd()

write-Output ($daten.multistatus.response.count.ToString() + ' VCards received')

$count = 0
$erg  = @()
# walk through vcards
foreach ($line in $daten.multistatus.response)
{
    # split vcard into an array
    $address = $line.propstat.prop.'address-data'.InnerText -split "\n"
    
    # walk through keywords
    foreach ($i in $begriffe)
    {
        if ($address | Select-String -pattern ("EMAIL.*"+$i) -Quiet)
        {
            $count ++
            # Return VCARD properties
            $obj = new-object psobject
            $obj | Add-Member  noteproperty count($count )
            $obj | Add-Member  noteproperty name((($address | Select-String  -simplematch "FN")  -split ":")[1] )
            $obj | Add-Member  noteproperty mail((($address | Select-String  -pattern ("EMAIL.*"+$i))  -split ":")[1] )
            $obj | Add-Member  noteproperty keyword($i)
            #$obj | Add-Member  noteproperty href($line.href )
            
            $erg += $obj
            # delete VCARD
            if ($simulation)
            { $obj | Add-Member  noteproperty status('Simulated')} 
            else
            {
                $x = curl -method "DELETE" -Credential $user -uri ( $server+$line.href )
                if ($x.StatusCode -eq "204")
                    {$obj | Add-Member  noteproperty status('Deleted')}
                    else
                    {$obj | Add-Member  noteproperty status('error')}

            }           
            break
        }
    }
}
Write-Verbose ($count.ToString() + ' VCards to delete')
Write-output $erg | Format-Table -AutoSize
$VerbosePreference = $oldverbose
