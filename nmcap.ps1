#Requires -RunAsAdministrator
#$test = "$PSScriptRoot\testing.txt"
write-debug "display the networks "
$Network = "cmd.exe c:\Program Files\Microsoft Network Monitor 3>nmcap /displaynetwork"
[string[]]$DisplayNetworks = Invoke-expression -Command $Network

    foreach($NetworkItems in $DisplayNetworks)
        {
            if()
        }

# -command $DisplayNetworks'
#Invoke-expression -command $DisplayNetworks'
#$Network | Out-File $test
  