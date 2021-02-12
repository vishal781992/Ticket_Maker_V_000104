; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
AppName=TicketMaker
AppVersion=0.0.12
DefaultDirName={pf}\ExcelReadingApp
; Since no icons will be created in "{group}", we don't need the wizard
; to ask for a Start Menu folder name:
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\TicketMaker.exe
AppPublisher=Vision Metering
;OutputDir=userdocs:Inno Setup Examples Output

OutputBaseFilename=TicketMaker_0.0.12_Setup

[Files]
Source: "C:\Project\ExcelReadingAppVersions\ExcelReadingApp_10_v010_working\ExcelReadingApp\bin\Release\ExcelReadingApp.exe"; DestDir: "{app}"

[Icons]
Name: "{commonprograms}\TicketMaker"; Filename: "{app}\TicketMaker.exe"
Name: "{commondesktop}\TicketMaker"; Filename: "{app}\TicketMaker.exe"

[Run]
Filename: "{tmp}\dotNetFx40_Full_x86_x64.exe"; Check: FrameworkIsNotInstalled

[code]
function FrameworkIsNotInstalled: Boolean;
begin
  Result := not RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\.NETFramework\policy\v4.0');
end;