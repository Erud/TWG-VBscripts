' List User Accounts on the Local Computer



Set colAccounts = GetObject("WinNT://0219-MAYERSKYL" )
colAccounts.Filter = Array("user")

For Each objUser In colAccounts
    Wscript.Echo objUser.Name 
Next