# Paste-dump
Just a paste Dump currently with Random Scripts TBD

____________________________________________________
Outlook Credentials login
# Connect to Outlook
$Outlook = New-Object -ComObject Outlook.Application
$NameSpace = $Outlook.GetNameSpace("MAPI")
$NameSpace.Logon("Outlook", "", $False, $True)

# Connect to Word
$Word = New-Object -ComObject Word.Application
$Word.Visible = $True
$Word.Activate()
$Word.UserName = "OutlookUserName"
$Word.UserInitials = "OU"
$Word.UserAddress = "outlook@example.com"

# Login to Outlook
$Username = "OutlookUsername"
$Password = "OutlookPassword" | ConvertTo-SecureString -AsPlainText -Force
$Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $Password
$Session = $Outlook.Session
$Session.AddStoreEx("Outlook", $OlStoreType::olMailbox)
$Session.Logon("Outlook", $Cred)
