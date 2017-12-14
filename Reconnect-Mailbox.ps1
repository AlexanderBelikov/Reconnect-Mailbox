Param(
  [Parameter(Mandatory=$True,Position=1)]
  $SourceUser,
  [Parameter(Mandatory=$True,Position=2)]
  $TargerUser
)

if ( !(Get-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue) ){
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

#Check SourceUser with UserMailbox 
$Source = Get-User $SourceUser -ErrorAction "SilentlyContinue";
if( -not($Source) -or ($Source.gettype().Name -ne "User") ){
    Write-Warning "SourceUser $SourceUser is not User type"
    Write-Warning "Aborting script..."
    Break;
}
if( $Source.RecipientType -ne "UserMailbox" ){
    Write-Warning "SourceUser $SourceUser is not UserMailbox RecipientType"
    Write-Warning "Aborting script..."
    Break;
}

$Mailbox = $Source | Get-Mailbox;
$DataBase = $Mailbox.Database;
$Server = Get-ExchangeServer (Get-MailboxDatabase $DataBase).Server;
$DCHostName = (Get-ADDomainController -Discover -ForceDiscover -SiteName (Get-ADSite (Get-ExchangeServer $Server).Site).Name).HostName[0]

#Check TargerUser is enabled and not UserMailbox 
$Targer = Get-User $TargerUser -DomainController $DCHostName -ErrorAction "SilentlyContinue";
if( -not($Targer) -or ($Targer.gettype().Name -ne "User") ){
    Write-Warning "TargerUser $TargerUser is not User type"
    Write-Warning "Aborting script..."
    Break;
}
if( -not($Targer.RecipientType -eq "User" -and $Targer.RecipientTypeDetails -eq "User") ){
    Write-Warning "TargerUser $TargerUser must be enabled and must not have a mailbox"
    Write-Warning "Aborting script..."
    Break;
}

$Mailbox | Select-Object ExchangeGuid, `
                PrimarySmtpAddress, `
                @{ n="EmailAddresses"; e={($_.EmailAddresses | ?{ $_.Prefix -like "SMTP"}).ProxyAddressString}}, `
                Database,
                @{ n="Server"; e={$Server.Name}}, `
                @{ n="DCHostName"; e={$DCHostName}} `
                | fl;

#Disable mailbox and Update
$Mailbox | Disable-Mailbox -DomainController $DCHostName;
Write-Output "Mailbox Disabled";

$i=0;
do{
    Write-Output "Try Update-StoreMailboxState -Database $DataBase -Identity $($Mailbox.ExchangeGuid)";
    Update-StoreMailboxState -Database $DataBase -Identity $Mailbox.ExchangeGuid;
    Start-Sleep -Seconds 5;
    $Disconnected = (Get-MailboxStatistics -StoreMailboxIdentity $Mailbox.ExchangeGuid -Database $DataBase -DomainController $DCName).DisconnectReason -eq "Disabled";
    $i++;
}while(($i -le 30) -and !($Disconnected));

if(!$Disconnected){
    Write-Warning "Update-StoreMailboxState failed";
    Write-Warning "Aborting script..."
    Break;
}

#Connect mailbox and set emailaddresses
Connect-Mailbox $Mailbox.ExchangeGuid -Database $DataBase -DomainController $DCHostName -User $Targer.DistinguishedName -Alias $Mailbox.Alias;

$Targer | Set-Mailbox -DomainController $DCHostName -EmailAddressPolicyEnabled $Mailbox.EmailAddressPolicyEnabled;
$Targer | Set-Mailbox -DomainController $DCHostName -PrimarySMTPAddress $Mailbox.PrimarySMTPAddress;

$TargerMailbox =  $Targer | Get-Mailbox -DomainController $DCHostName;
$EmailAddresses = $TargerMailbox.EmailAddresses | ?{ $_.PrefixString -ceq "smtp" } | Select-Object ProxyAddressString;
$EmailAddresses.ProxyAddressString | %{ $TargerMailbox.EmailAddresses.Remove($_) | Out-Null }
$Mailbox.EmailAddresses | ?{ $_.PrefixString -ceq "smtp" } | %{ $TargerMailbox.EmailAddresses.Add($_) }

$Targer | Set-Mailbox -DomainController $DCHostName -EmailAddresses $Mailbox.EmailAddresses;

if(($Targer | Get-Mailbox -DomainController $DCHostName).RecipientType -eq "UserMailbox"){
    Write-Output "Mailbox reconnected successfully";
} else {
    Write-Output "Mailbox reconnect failed";
}
