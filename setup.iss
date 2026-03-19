;-----------------------------------------------------------
; Table Viewer -- Inno Setup Script
; Önce build.bat'ı çalıştırın, ardından bu dosyayı
; Inno Setup Compiler ile derleyin.
;-----------------------------------------------------------

#define AppName      "Table Viewer"
#define AppVersion   "1.0.0"
#define AppPublisher "Talha"
#define AppExeName   "TableViewer.exe"
#define BuildDir     "dist\TableViewer"

[Setup]
AppId={{F3A7C2B1-4E8D-4F9A-B2C3-1A2B3C4D5E6F}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisherURL=
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
OutputDir=installer
OutputBaseFilename=TableViewer-Setup-v{#AppVersion}
SetupIconFile=assets\favicon.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequiredOverridesAllowed=dialog
; Kurulum sırasında "Başlat Menüsü" ve "Masaüstü" seçenekleri
; kullanıcıya sunulur
AllowNoIcons=yes
; Windows 10 ve üstü
MinVersion=10.0

[Languages]
Name: "turkish";  MessagesFile: "compiler:Languages\Turkish.isl"
Name: "english";  MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";    Description: "Masaüstüne kısayol oluştur";    GroupDescription: "Ek kısayollar:"; Flags: unchecked
Name: "fileassoc_xlsx"; Description: ".xlsx dosyalarını Table Viewer ile aç"; GroupDescription: "Dosya ilişkilendirmeleri:"; Flags: unchecked
Name: "fileassoc_xls";  Description: ".xls dosyalarını Table Viewer ile aç";  GroupDescription: "Dosya ilişkilendirmeleri:"; Flags: unchecked
Name: "fileassoc_csv";  Description: ".csv dosyalarını Table Viewer ile aç";  GroupDescription: "Dosya ilişkilendirmeleri:"; Flags: unchecked

[Files]
; PyInstaller çıktısının tamamını kopyala
Source: "{#BuildDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Başlat menüsü
Name: "{group}\{#AppName}";       Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"
Name: "{group}\Kaldır";           Filename: "{uninstallexe}"
; Masaüstü (göreve bağlı)
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Registry]
; .xlsx ilişkilendirmesi
Root: HKCU; Subkey: "Software\Classes\.xlsx\OpenWithProgids"; ValueType: string; ValueName: "TableViewer.XLSX"; ValueData: ""; Flags: uninsdeletevalue; Tasks: fileassoc_xlsx
Root: HKCU; Subkey: "Software\Classes\TableViewer.XLSX";                        ValueType: string; ValueName: ""; ValueData: "{#AppName}";                        Flags: uninsdeletekey; Tasks: fileassoc_xlsx
Root: HKCU; Subkey: "Software\Classes\TableViewer.XLSX\DefaultIcon";            ValueType: string; ValueName: ""; ValueData: "{app}\{#AppExeName},0";             Tasks: fileassoc_xlsx
Root: HKCU; Subkey: "Software\Classes\TableViewer.XLSX\shell\open\command";     ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1""";    Tasks: fileassoc_xlsx

; .xls ilişkilendirmesi
Root: HKCU; Subkey: "Software\Classes\.xls\OpenWithProgids"; ValueType: string; ValueName: "TableViewer.XLS"; ValueData: ""; Flags: uninsdeletevalue; Tasks: fileassoc_xls
Root: HKCU; Subkey: "Software\Classes\TableViewer.XLS";                        ValueType: string; ValueName: ""; ValueData: "{#AppName}";                       Flags: uninsdeletekey; Tasks: fileassoc_xls
Root: HKCU; Subkey: "Software\Classes\TableViewer.XLS\DefaultIcon";            ValueType: string; ValueName: ""; ValueData: "{app}\{#AppExeName},0";            Tasks: fileassoc_xls
Root: HKCU; Subkey: "Software\Classes\TableViewer.XLS\shell\open\command";     ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1""";   Tasks: fileassoc_xls

; .csv ilişkilendirmesi
Root: HKCU; Subkey: "Software\Classes\.csv\OpenWithProgids"; ValueType: string; ValueName: "TableViewer.CSV"; ValueData: ""; Flags: uninsdeletevalue; Tasks: fileassoc_csv
Root: HKCU; Subkey: "Software\Classes\TableViewer.CSV";                        ValueType: string; ValueName: ""; ValueData: "{#AppName}";                       Flags: uninsdeletekey; Tasks: fileassoc_csv
Root: HKCU; Subkey: "Software\Classes\TableViewer.CSV\DefaultIcon";            ValueType: string; ValueName: ""; ValueData: "{app}\{#AppExeName},0";            Tasks: fileassoc_csv
Root: HKCU; Subkey: "Software\Classes\TableViewer.CSV\shell\open\command";     ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1""";   Tasks: fileassoc_csv

[Run]
; Kurulum bittikten sonra uygulamayı başlatma seçeneği
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Kaldırma sırasında uygulama klasörünü tamamen temizle
Type: filesandordirs; Name: "{app}"
