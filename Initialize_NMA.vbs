Option Explicit

Dim oShell, oRegistry, sRegKey, sKeyPath, sValueName, NMA_APIKey,  NMAKeyExists

Const HKCU = &H80000001

Set oShell=CreateObject("WScript.Shell")
Set oRegistry=GetObject("winmgmts:\\.\root\default:StdRegProv")

sRegKey="HKCU\Software\boredazfcuk\mining\NMA_APIKey"
sKeyPath="Software\boredazfcuk\mining"
sValueName="NMA_APIKey"
NMA_APIKey=InputBox("Please enter your Prowl API Key (Cancel to delete):", "NMA API")

If ((NMA_APIKey=Null) Or (NMA_APIKey="")) Then
	oRegistry.GetStringValue HKCU, sKeyPath, sValueName, NMAKeyExists
	If IsNull (NMAKeyExists) Then
		WScript.Echo "Key does not exist"
	Else
		oShell.RegDelete sRegKey
		WScript.Echo "Registry key deleted"
	End If
Else
	oShell.RegWrite sRegKey, NMA_APIKey, "REG_SZ"
	WScript.Echo "API Key: " & NMA_APIKey & " written to registry."
End If
