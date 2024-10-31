$Updates = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher().Search("").Updates
$Updates
$Updates | % {$_.Title}
$Updates = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher().Search("IsHidden=0").Updates
$Updates
$Updates | % {$_.Title}
$Updates = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher().Search("IsHidden=0 and IsInstalled=0").Updates
$Updates | % {$_.Title}
$Updates | % {$_.EulaAccepted=$true}
$WUUpdates = New-Object -ComObject Microsoft.Update.UpdateColl
$Updates | Out-GridView -OutputMode Multiple | % {$WUUpdates.Add($_)}
$Updates
$Updates.Count
$WUUpdates.Count
$WUInstaller = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateInstaller
$WUDownloader = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateDownloader()
$WUInstaller = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateInstaller()
$WUInstaller.Updates=$WUUpdates
$WUDownloader.Updates=$WUUpdates
$WUDownloader.Download()
$WUInstaller.Install()






$Updates = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher().Search("IsHidden=0 and IsInstalled=0").Updates
$SqlUpdates = $Updates | Where-Object {$_.Title -like "*SQL*"}
$WUUpdates = New-Object -ComObject Microsoft.Update.UpdateColl
$null = $SqlUpdates | Foreach-Object { $WUUpdates.Add($_) }
$WUInstaller = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateInstaller
$WUDownloader = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateDownloader()
$WUInstaller = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateInstaller()
$WUInstaller.Updates=$WUUpdates
$WUDownloader.Updates=$WUUpdates
$WUDownloader.Download()
$WUInstaller.Install()