REG ADD  “HKLM\Software\policies\Microsoft\Windows NT\DNSClient”
REG ADD  “HKLM\Software\policies\Microsoft\Windows NT\DNSClient” /v ” EnableMulticast” /t REG_DWORD /d “0” /f 
netsh advfirewall firewall set rule group="Network Discovery" new enable=No