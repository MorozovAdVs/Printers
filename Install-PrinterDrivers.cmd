set driverdirectory=\\dipaul.corp\NETLOGON\drv
for /r %driverdirectory% %%f in ("*.inf") do (pnputil /add-driver %%f /install)

rundll32.exe printui.dll,PrintUIEntry /ia /u /m "HP Universal Printing PCL 6 (v7.2.0)"
rundll32.exe printui.dll,PrintUIEntry /ia /u /m "Kyocera ECOSYS M2540dn KX"
rundll32.exe printui.dll,PrintUIEntry /ia /u /m "Kyocera ECOSYS M2235dn KX"
rundll32.exe printui.dll,PrintUIEntry /ia /u /m "Kyocera ECOSYS M2735dn KX"
rundll32.exe printui.dll,PrintUIEntry /ia /u /m "Kyocera ECOSYS M5526cdw KX"
