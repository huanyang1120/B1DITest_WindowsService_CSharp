# B1DITest_WindowsService_CSharp

Powershell:
Regiester the Windwos Service

New-Service -Name "B1DITest_WindowsService_CSharp" -BinaryPathName "C:\Temp\COM DI\CSharp\B1DITest_WindowsService_CSharp\bin\Debug\B1DITest_WindowsService_CSharp.exe" -DisplayName "B1DITest_WindowsService_CSharp" -StartupType Automatic

Remove the Windows Service
sc.exe delete "B1DITest_WindowsService_CSharp"
