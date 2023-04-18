# MSExchangeMailForwardingScript
This simple Powershell script, automates mail forwarding creation on Microsoft Exchange Online platform using csv file. 
Script works in loop and automatically(depending on the dates in csv file)  creates or  deletes distribution groups and set/removes mail forwardings.   


>Script needs for proper working csv file, for example  "csvexo.csv"

>For proper working of the script you should install additional PowerShell MSExchangeOnline module using this command:
    Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.1.0


>For security reasons I suggest to create special Exchange account and add only certain permissions for it.
      
      




