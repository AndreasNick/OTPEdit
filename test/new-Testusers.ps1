$path="OU=testuser,OU=uran,DC=uran,DC=local"
$username="testuser"
$count=1..1500
foreach ($i in $count)
{ New-AdUser -Name $username$i -Path $path -Enabled $True -ChangePasswordAtLogon $true -UserPrincipalName $("$username"+$i+"Uran.local")  -AccountPassword (ConvertTo-SecureString "P@ssw0rd" -AsPlainText -force) -passThru }



 