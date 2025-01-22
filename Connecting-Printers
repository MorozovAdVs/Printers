Dim Window, GroupName, Pattern, ListNewPrinter

Set ListNewPrinter = CreateObject("Scripting.Dictionary")

const AddScriptPath = "cscript C:\Windows\System32\Printing_Admin_Scripts\ru-RU\prnmngr.vbs -ac -p "
const DelScriptPath = "cscript C:\Windows\System32\Printing_Admin_Scripts\ru-RU\prnmngr.vbs -xc "
const DefaultScriptPath = "cscript C:\Windows\System32\Printing_Admin_Scripts\ru-RU\prnmngr.vbs -t -p "
Pattern = "(?:mfp|prn)"

Set re = New RegExp
With re
	.Pattern    = Pattern 
	.IgnoreCase = False
	.Global     = False
End With

Set objSysInfo = CreateObject("ADSystemInfo")
Set objUser = GetObject("LDAP://" & objSysInfo.UserName)

For Each strCN In objUser.MemberOf
    Set objGroup = GetObject("LDAP://" & strCN)
	GroupName = replace(objGroup.name, "CN=", "")
	If re.Test( GroupName ) Then
		ListNewPrinter.Add GroupName, ""
	End If
Next

If IsEmpty(ListNewPrinter) Then
	MsgBox "Доступные группы принтеров не найдены",16,"Установка принтеров"	
Else
	Window = MsgBox ("Запустить установку этих принтеров?"&vbCrLf & vbCrLf & join(ListNewPrinter.Keys, vbCrLf) & vbCrLf & vbCrLf & "Если нужного принтера нет в списке - презагрузитесь или перезайдите в систему.",32+4,"Установка принтеров")
	If Window = 6 Then
		CreateObject("WScript.Shell").Run DelScriptPath, 0, True
		If ListNewPrinter.Count = 1 Then
			PrinterName = join(ListNewPrinter.Keys, "")
			If InStr(PrinterName, "str") > 0 Then
				ServerLink = "\\str-prn\"
			Else
				ServerLink = "\\spb-prn22\"
			End If
			CreateObject("WScript.Shell").Run AddScriptPath + ServerLink + PrinterName, 0, True
			CreateObject("WScript.Shell").Run DefaultScriptPath + ServerLink + PrinterName, 0, True
			MsgBox "По-умолчанию установлен принтер " & PrinterName & ".", 0, "Установка принтеров"
		Else
			For Each PrinterName In ListNewPrinter
				If InStr(PrinterName, "str") > 0 Then
					ServerLink = "\\str-prn\"
				Else
					ServerLink = "\\spb-prn22\"
				End If
				CreateObject("WScript.Shell").Run AddScriptPath + ServerLink + PrinterName, 0, True
				Next
			MsgBox "После нажатия 'ОК' будет открыто окно настройки принтеров." & vbCrLf & vbCrLf & "В нем выберите нужный принтер." & vbCrLf & "Нажмите на него правой кнопкой мыши." & vbCrLf & "В открывшемся меню выберите 'Использовать по умолчанию'.",32+0,"Установка принтеров"
			CreateObject("WScript.Shell").Run "control /name {A8A91A66-3A7D-4424-8D24-04E180695C7A}"
		End If
		CreateObject("WScript.Shell").Run "taskkill /IM mstsc.exe", 0, True
	Else
		MsgBox "Отмена выполнения",16,"Установка принтеров"
	End If
End If
