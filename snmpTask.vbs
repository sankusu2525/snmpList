set objLocator = CreateObject("WbemScripting.SWbemLocator") 
set objServices = objLocator.connectServer("", "root/snmp/localhost") 
set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet") 
Set objFso = CreateObject("Scripting.FileSystemObject") 

'ファイルPath取得
strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")

'データファイルOpen
strOpenFile = objFso.BuildPath(strScriptPath,"ip-list.txt")
Set objTextData = objFso.OpenTextFile(strOpenFile ,1)

'日時取得
Dim wkNow
wkNow= Year(Now())
wkNow= wkNow & Right("0" & Month(Now()) , 2)
wkNow= wkNow & Right("0" & Day(Now()) , 2)
wkNow= wkNow & "_"
wkNow= wkNow & Right("0" & Hour(Now()) , 2)
wkNow= wkNow & Right("0" & Minute(Now()) , 2)
wkNow= wkNow & Right("0" & Second(Now()) , 2)
'wscript.echo wkNow

'書き込みファイルOpen
strOpenFile = objFso.BuildPath(strScriptPath, wkNow & ".csv")
Set objTextWrite = objFso.OpenTextFile(strOpenFile ,2,True)

'タイトル書き込みフラグ
flg = 1

'SNMP_IP取り込み
Do Until objTextData.AtEndOfLine = True
   strIP_Text = objTextData.ReadLine
   strText_name = chr(34)&  "IP" & chr(34) & "," & chr(34)
   strText_value = chr(34)& strIP_Text & chr(34) & "," & chr(34)

'ping実施
   dim timeOut
'タイムアウト設定(ミリ秒)
   timeOut = "1000"
   Set objWMIService = GetObject("winmgmts:\\.")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_PingStatus " & "Where Timeout = " & timeOut & " AND Address = '" & strIP_Text & "'")
   strText_name = strText_name & "Ping" & chr(34) & "," & chr(34)
   For Each objItem in colItems
      If objItem.StatusCode = 0 Then
         strText_value = strText_value & "OK" & chr(34) & "," & chr(34)
'SNMP取得
         objNamedValueSet.Add "AgentAddress", strIP_Text
'wscript.echo strIP_Text
         set objset = objServices.instancesof( "SNMP_RFC1213_MIB_system", ,objNamedValueSet ) 
'エラー停止禁止
         On Error Resume Next
         for each obj in objset
            for each prop in obj.properties_ 
'wscript.echo prop.name & " -- " & prop.value
               strText_name = strText_name & prop.name & chr(34) & "," & chr(34)
               strText_value = strText_value & prop.value & chr(34) & "," & chr(34)
            next
         next
         strText_name = strText_name & chr(34)
         strText_value = strText_value & chr(34)
         if flg = 1 then
            objTextWrite.WriteLine strText_name
            flg = 0
         end if
      else
         strText_value = strText_value & "NG" & chr(34)
      End If
   Next
   objTextWrite.WriteLine strText_value
Loop

objTextData.Close
objTextWrite.Close

Set objTextData = Nothing
Set objTextWrite = Nothing
Set objFileSys = Nothing

wscript.echo "snmp list end"
