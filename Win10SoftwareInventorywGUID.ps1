New-Item -ItemType Directory -Force -Path C:\SoftwareInventory
Get-CimInstance -ClassName Win32_Product | Sort-Object -Property Name | Select IdentifyingNumber, Name, LocalPackage | Export-CSV "C:\SoftwareInventory\softwareInventory.csv" -NoTypeInformation
$Outlook = New-Object -ComObject Outlook.Application
        $Mail = $Outlook.CreateItem(0)
         $hostname = hostname
         $Mail.To = "email@domain.com"
         $date = Get-Date -Format g
         $Mail.Subject = "Sofware Inventory Results for $hostname $date"
         $Mail.Body = "Here are the software inventory results for $hostname"
         $file = "C:\SoftwareInventory\softwareInventory.csv"
        $Mail.Attachments.Add($file)
         $Mail.Send()