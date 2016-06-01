    SET WshShell = WScript.CreateObject("WScript.Shell")
    SET WshSysEnv = WshShell.Environment("SYSTEM")
    SET FSO = CreateObject("Scripting.FileSystemObject")
    IF WScript.Arguments.Count <> 0 Then
        FOR EACH arg IN WScript.Arguments
            iArgCount = iArgCount + 1
            strCmdArg = (arg)
            strCmdArray = Split(strCmdArg, " ", 2, 1)
            IF iArgCount = 1 THEN
            strExe = strCmdArray(0)
            ELSEIF iArgCount = 2 THEN
            strRun = strCmdArray(0)
            ELSE
            strParams = strParams&" "&strCmdArray(0)
            END IF
        NEXT
    END IF
'/t:0A && title ***** Admin ***** 
        strExt = LCase(Right(strExe, 3))

IF strExt <> "exe" AND strExt <> "bat" AND strExt <> "cmd" THEN
WshShell.Run "psexec.exe -d -i -e -u COMPUTERNAME\USER -p PASSWORD  cmd /c start "&strExe&" "&strRun&" "&strParams, 0, FALSE
ELSE
WshShell.Run "psexec.exe -d -i -e -u COMPUTERNAME\USER -p PASSWORD "&strExe&" "&strRun&" "&strParams, 0, FALSE
END IF

    SET WshShell = NOTHING
    SET WshSysEnv = NOTHING
    SET FSO = NOTHING
