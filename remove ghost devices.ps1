    #List all hidden devices
     $unknown_devs = Get-PnpDevice | Where-Object{$_.Status -eq 'Unknown'} 

    #loop through all hidden devices to remove them using pnputil
     ForEach($dev in $unknown_devs){
     pnputil /remove-device $dev.InstanceId
     }