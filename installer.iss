; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{ED721D26-0DD1-4A69-A93C-4D6657512ADB}
AppName=alphaDrive
AppVersion=1.5
;AppVerName=alphaDrive 1.5
AppPublisher=LivingAnalytics, Inc.
AppPublisherURL=blog.livinganalytics.com
AppSupportURL=blog.livinganalytics.com
AppUpdatesURL=blog.livinganalytics.com
DefaultDirName={pf}\alphaDrive
DefaultGroupName=alphaDrive
LicenseFile=C:\Users\Chris Riley\LIVINGANALYTICS\Projects\alphaDrive\ShareDriveAnalyze\bin\Release\laeula.txt
OutputDir=C:\Users\Chris Riley\Desktop
OutputBaseFilename=alphaDrive_Setup
SetupIconFile=C:\Users\Chris Riley\LIVINGANALYTICS\Projects\alphaDrive\ShareDriveAnalyze\1260313293_drives_31.ico
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\Users\Chris Riley\LIVINGANALYTICS\Projects\alphaDrive\ShareDriveAnalyze\bin\Release\alphaDrive.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Chris Riley\LIVINGANALYTICS\Projects\alphaDrive\ShareDriveAnalyze\bin\Release\Interop.DSOFile.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Chris Riley\LIVINGANALYTICS\Projects\alphaDrive\ShareDriveAnalyze\bin\Release\ZedGraph.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Chris Riley\LIVINGANALYTICS\Projects\alphaDrive\ShareDriveAnalyze\bin\Release\ZedGraph.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Chris Riley\LIVINGANALYTICS\Projects\alphaDrive\ShareDriveAnalyze\bin\Release\de\*"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Chris Riley\LIVINGANALYTICS\Projects\alphaDrive\DsoFileSetup_KB224351_x86.exe"; DestDir:  "{tmp}"
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\alphaDrive"; Filename: "{app}\alphaDrive.exe"
Name: "{commondesktop}\alphaDrive"; Filename: "{app}\alphaDrive.exe"; Tasks: desktopicon

[Run]
Filename: "DsoFileSetup_KB224351_x86.exe"; Parameters: "/q ""{tmp}\DsoFileSetup_KB224351_x86.exe"""
Filename: "{app}\alphaDrive.exe"; Description: "{cm:LaunchProgram,alphaDrive}"; Flags: nowait postinstall skipifsilent

