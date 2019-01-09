On Error Resume Next
 

Set colAccounts = GetObject("LDAP://OU=Prof,OU=Compte,OU=ndchateau.fr,DC=ndchateau,DC=fr")
colAccounts.Filter = Array("user")
Set objLastLogon = objUser.Get("lastLogonTimestamp")

For Each objUser In colAccounts
'Set objUser = GetObject("LDAP://cn=username,dc=yourdomain,dc=com")
'Set objLastLogon = objUser.Get("lastLogonTimestamp")



If (i8Low < 0) Then
    i8High = i8High + 1
End If



name = objUser.name

Wscript.Echo name & "Last logon time: " & intLastLogonTime + #1/1/1601#
Next


