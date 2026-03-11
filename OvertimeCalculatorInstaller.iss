; Inno Setup script for Overtime Calculator
; 1) Install Inno Setup from https://jrsoftware.org/
; 2) Open this .iss file in Inno Setup Compiler
; 3) Build to generate OvertimeCalculatorSetup.exe installer

[Setup]
AppName=Overtime Calculator
AppVersion=1.0.0
DefaultDirName={pf}\Overtime Calculator
DefaultGroupName=Overtime Calculator
OutputBaseFilename=OvertimeCalculatorSetup
Compression=lzma
SolidCompression=yes
DisableDirPage=no
DisableProgramGroupPage=no

[Files]
Source: "C:\Users\Tech\Desktop\overtime_calculator\dist\OvertimeCalculator.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Overtime Calculator"; Filename: "{app}\OvertimeCalculator.exe"
Name: "{commondesktop}\Overtime Calculator"; Filename: "{app}\OvertimeCalculator.exe"
