# WindowsServiceTestB1DI

Powershell:
Regiester the Windwos Service

New-Service -Name "WindowsServiceTestB1DI" -BinaryPathName "C:\Temp\COM DI\CSharp\WindowsServiceTestB1DI\bin\Debug\WindowsServiceTestB1DI.exe" -DisplayName "My B1DI Test Windows" -StartupType Automatic

Remove the Windows Service
sc.exe delete "WindowsServiceTestB1DI"