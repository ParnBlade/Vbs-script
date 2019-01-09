Const ADS_PROPERTY_CLEAR = 1 

Set objOU = GetObject("LDAP://OU=Equipe,OU=Compte,OU=ndchateau.fr,DC=ndchateau,DC=fr")
objOU.Filter = Array("Group")

For Each objGroup in objOU
'group = replace(objGroup.Name,"CN=","")
    'Wscript.Echo objGroup.Name

strNameGroup = objGroup.Name


 
 
 'ex: strNameGroup = "CN=Equipe_1ES2"
Set objGroup = GetObject _
 ("LDAP://"& strNameGroup &",OU=Equipe,OU=Compte,OU=ndchateau.fr,DC=ndchateau,DC=fr") 
 
objGroup.PutEx ADS_PROPERTY_CLEAR, "member", 0
 
objGroup.SetInfo
Next