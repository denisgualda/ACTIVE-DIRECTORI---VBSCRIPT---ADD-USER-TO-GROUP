'AFEGIR GRUPS A USUARI

addusertogroup "LDAP://CN=GRUP1,OU=Grups,DC=domini,DC=lab","CN=usuari02,OU=Usuaris,DC=domini,DC=lab"


sub addusertogroup (dsgrup,dsusuari)
Const ADS_PROPERTY_APPEND = 3

Set objGroup = GetObject (dsgrup)
objGroup.PutEx ADS_PROPERTY_APPEND,"member", Array(dsusuari)
objGroup.SetInfo
wscript.echo "Afegit:" & dsgrup & "a usuari: " & dsusuari

End sub
