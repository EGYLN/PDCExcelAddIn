The following steps assume an installation on a 64-Bit Windows system.
PDC is a 32-Bit Addon to Excel. Therefore the respective 32-Bit registration tools are used.
If the default 64-Bit versions are used, the registry entries may end up at the wrong place.

The registration needs administration rights.  The AddIn itsself is installed on user level (HKCU), 
while the former Setup program allowed to choose between user and machine level.
If a former (.Net 2) version of PDC is already installed, it should be uninstalled first.

Installationssteps:

1.) Register the (COM) assemblies PDCExcelAddIn.dll and OpenLib.dll in the installation directory. 
    The .Net Framework path (v4.0.30319) may differ, depending on the concrete version, which is installed.
    Important is, that is contains a v4 and that it is NOT the Framework64 path.
	c:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm PDCExcelAddIn.dll /tlb /codebase
	- cd <InstallationPath> (usually C:\Program Files (x86)\BayerBBS\PDCExcelAddIn)
	- c:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm PDCExcelAddIn.dll /tlb /codebase
	- c:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm OpenLib.dll /tlb /codebase
		
2.) Register the PDC addin in Excel via registry import. 
	a) 	Replace %INSTALLATIONPATH% in the files pdc_vsto_HKCR.reg pdc_vsto_HKCU.reg and  with the actual installation path 
		Directory separators must be doubled (e.g. something like C:\\Program Files (x86)\\BayerBBS\\PDCExcelAddIn.
	b) Register the assembly PDCExcelAddIn.dll under HKEY_CLASSES_ROOT. This needs administration privileges!
 		c:\Windows\SysWOW64\reg.exe import pdc_vsto_HKCR.reg
	c) Register the assembly PDCExcelAddIn.dll under HKEY_CURRENT_USER as a VSTO Excel Addin. 
		This must be done "as" the user, for whom the AddIn should be installed. Administration privileges are not necessary. 
		c:\Windows\SysWOW64\reg.exe import pdc_vsto_HKCU.reg	

Deinstallation:

1.) Unregister the (COM) assemblies PDCExcelAddIn.dll and OpenLib.dll in the installation directory.

	The .Net Framework path (v4.0.30319) may differ, depending on the concrete version, which is installed.
    Important is, that is contains a v4 and that it is NOT the Framework64 path.
	c:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm /u PDCExcelAddIn.dll 
	- cd <InstallationPath> (usually C:\Program Files (x86)\BayerBBS\PDCExcelAddIn)
	- c:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm /u PDCExcelAddIn.dll
	- c:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm /u OpenLib.dll
	
2.) Undo the registry import from step 2 of the installation routine. The unregister*.reg files can be imported with reg.exe to delete the registry keys from the installation.
	a) Unregister the assembly PDCExcelAddIn.dll under HKEY_CLASSES_ROOT. This needs administration privileges!
 		c:\Windows\SysWOW64\reg.exe import unregister_pdc_vsto_HKCR.reg
	b) Unregister the assembly PDCExcelAddIn.dll under HKEY_CURRENT_USER as a VSTO Excel Addin. 
		This must be done "as" the user, for whom the AddIn should be installed. Administration privileges are not necessary. 
		c:\Windows\SysWOW64\reg.exe import unregister_pdc_vsto_HKCU.reg
	

