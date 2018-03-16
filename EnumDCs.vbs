Set objOU = GetObject("LDAP://ou=Domain Controllers, dc=smi-rps, dc=ajgco, dc=com")

objOU.Filter = Array("Computer")

For Each objComputer in objOU

    Wscript.Echo objComputer.CN

Next