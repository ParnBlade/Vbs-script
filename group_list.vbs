Option Explicit
Dim objNetwork, strDomain, strUser, objUser, objGroup, strGroupMemberships, strGroup

' Get the domain and username from the WScript.Network object
Set objNetwork = CreateObject("WScript.Network")
strDomain = "ndchateau"'objNetwork.UserDomain
'wscript.echo strDomain
strUser = "p.p"'objNetwork.UserName
'wscript.echo strUser
strGroup = "Prof"

' Instanciate the user object from the data above
Set objUser = GetObject("WinNT://" & strDomain & "/" & strUser)


' Run through the users groups and put them in the string
For Each objGroup In objUser.Groups
if objGroup.Name = strGroup then
	wscript.echo "Dans groupe = "& strGroup
Exit For
Else
    'strGroupMemberships = strGroupMemberships & objGroup.Name & ","
	strGroupMemberships = objGroup.Name
'end if
wscript.echo strGroupMemberships
end if
Next

'MsgBox strGroupMemberships