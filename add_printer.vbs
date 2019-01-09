Dim WSHShell, objNET, objSysInfo, objComputer, strComputerDN, strGroups, Group, GroupName
Set WSHShell = WScript.CreateObject("WScript.Shell")
Set objNET = WScript.CreateObject("WScript.Network")
Set objSysInfo = WScript.CreateObject("ADSystemInfo")
strComputerDN = objSysInfo.COMPUTERNAME
Set objComputer = GetObject("LDAP://" & strComputerDN) 'Binds the objComputer to the Distiguished Name of the Computer in reference   

strGroups = objComputer.GetEx("memberOf")   
For Each Group in strGroups
	Group = Mid(Group, 4, 330)
	arrGroup = Split(Group, "," )
	strList = arrGroup(0)

	
		Select Case strList
		
			Case "CDI"
            Set objNetwork = CreateObject("WScript.Network")
				objNetwork.AddWindowsPrinterConnection "\\SRV-FS-LYC\CDI"
				objNetwork.SetDefaultPrinter "\\SRV-FS-LYC\CDI"
				
			Case "102_1"
            Set objNetwork = CreateObject("WScript.Network")
				objNetwork.AddWindowsPrinterConnection "\\PEDAGO-LYC\102"
				objNetwork.SetDefaultPrinter "\\PEDAGO-LYC\102"
				
				Case "102_2"
            Set objNetwork = CreateObject("WScript.Network")
				objNetwork.AddWindowsPrinterConnection "\\PEDAGO-LYC\102"
				objNetwork.SetDefaultPrinter "\\PEDAGO-LYC\102"
				
			Case "104"
            Set objNetwork = CreateObject("WScript.Network")
				objNetwork.AddWindowsPrinterConnection "\\PEDAGO-LYC\104"
				objNetwork.SetDefaultPrinter "\\PEDAGO-LYC\104"
				
			Case "121"
            Set objNetwork = CreateObject("WScript.Network")
				objNetwork.AddWindowsPrinterConnection "\\121-PROF\121"
				objNetwork.SetDefaultPrinter "\\121-PROF\121"
				
			Case "121-local"
            'Set objNetwork = CreateObject("WScript.Network")
				'objNetwork.AddWindowsPrinterConnection "\\121-PROF\121"
				Wscript.quit
				
			Case "120"
            Set objNetwork = CreateObject("WScript.Network")
				objNetwork.AddWindowsPrinterConnection "\\PEDAGO-LYC\120"
				objNetwork.SetDefaultPrinter "\\PEDAGO-LYC\120"
				
							
			Case "123"
            Set objNetwork = CreateObject("WScript.Network")
				objNetwork.AddWindowsPrinterConnection "\\121-PROF\121"
				objNetwork.SetDefaultPrinter "\\121-PROF\121"
				
			Case "008"
            Set objNetwork = CreateObject("WScript.Network")
				objNetwork.AddWindowsPrinterConnection "\\PEDAGO-LYC\STI"
				objNetwork.SetDefaultPrinter "\\PEDAGO-LYC\STI"				
            
            Case "009"
            Set objNetwork = CreateObject("WScript.Network")
				objNetwork.AddWindowsPrinterConnection "\\PEDAGO-LYC\STI"
				objNetwork.SetDefaultPrinter "\\PEDAGO-LYC\STI"
 
            Case "010"
            Set objNetwork = CreateObject("WScript.Network")
				objNetwork.AddWindowsPrinterConnection "\\PEDAGO-LYC\STI"
				objNetwork.SetDefaultPrinter "\\PEDAGO-LYC\STI"				
		
		End Select

next

WScript.Quit 