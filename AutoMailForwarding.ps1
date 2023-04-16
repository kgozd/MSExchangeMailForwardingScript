function GettingDateAndCSVDirectory {

    $script:CurrentDate = Get-Date -Format "dd/MM/yyyy"
    $script:DirectoryToCSV = Join-Path $PSScriptRoot "csvexo.csv" 
}

function LogToExO {

    $password = ConvertTo-SecureString '<password>' -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential ('<script_login_mail>', $password)
    Connect-ExchangeOnline -Credential $credential
}

function CreatingDistroGroupName {
    $script:MailTrim = $MailToForward.split('@')[0] 
    $script:DistroGroup =  $MailTrim + '_forwarding@' + '<your_mail_domain>' + '.onmicrosoft.com'
}

function EnablingForwarding {

    if($DateON -eq $CurrentDate ){
                                 
        Try{New-DistributionGroup $DistroGroup -Members $MailF1}
        catch{Write-Host ("Distribution group already created")}
        
        Try{Add-DistributionGroupMember -Identity $DistroGroup -Members $MailF2}
        catch{Write-Host("User $MailF2 couldn't be added to $DistroGroup")}
        
        Try{Add-DistributionGroupMember -Identity $DistroGroup -Members $MailF3}
        catch{Write-Host("User $MailF3 couldn't be added to $DistroGroup")}
        
        Try{Set-Mailbox -Identity $MailToForward -ForwardingSMTPAddress $DistroGroup -DeliverToMailboxAndForward $true -force}
        catch{Write-Host("Mail forwarding already enabled")}
    }
}

function DisablingForwarding{

    if($DateOFF -eq $CurrentDate ){
                   
        Try{Remove-DistributionGroup   $DistroGroup -Confirm:$false}
        catch{Write-Host("Distribution group $DistroGroup cannot be deleted because, it doesnt exist")}

        Try{Set-Mailbox $MailToForward -ForwardingAddress $NULL -ForwardingSmtpAddress $NULL -force}
        catch{Write-Host("Mail $MailToForward forwarding cannot be disabled")}  
    } 
}

function RemovingCSVRow {

    Try{    
        $content = Get-Content $DirectoryToCSV | Where-Object { $_.Split(",")[2] -notmatch $CurrentDate }
        $content | Set-Content $DirectoryToCSV 
    }
    catch{"Rows could't be removed"}

}

function LoopingThroughRows{

    ForEach ($line in $CSV){

        $script:MailToForward = $($line.Mail)
        $script:DateON = $($line.Date1)
        $script:DateOFF = $($line.Date2)
        $script:MailF1 = $($line.MailF1)
        $script:MailF2 = $($line.MailF2)
        $script:MailF3 = $($line.MailF3)

        if($DateON -or $DateOFF -eq $CurrentDate ){
            &CreatingDistroGroupName
            &EnablingForwarding
            &DisablingForwarding    
        }
    }
}

function Main {
    &GettingDateAndCSVDirectory
    &LogToExO

    if (Test-Path -Path $DirectoryToCSV){
        $script:CSV =  Import-Csv -Path $DirectoryToCSV
        &LoopingThroughRows
        &RemovingCSVRow
        Disconnect-ExchangeOnline -Confirm:$false
        Start-Sleep -Seconds 60
    }  
}

while($true){
    &Main
}


  
