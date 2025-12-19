; สคริปต์ Inno Setup สำหรับโปรแกรม APLUS HR & Payroll
; สร้างโดย Gemini AI

#define MyAppName "APLUS Smart HR & Payroll"
#define MyAppVersion "1.0"
#define MyAppPublisher "APLUS Com"
#define MyAppExeName "APLUS_HR_Payroll.exe"
#define MyIcon "company_logo.png" 

[Setup]
; --- ข้อมูลพื้นฐาน ---
AppId={{8A29C8B3-F9D1-4A5E-9F8C-1234567890AB}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
; ลบไฟล์เดิมก่อนลงใหม่ (Clean Install)
DirExistsWarning=no

; --- การตั้งค่าไฟล์ Output (ตัวติดตั้งที่จะได้) ---
OutputDir=C:\Users\Nitro V15\Desktop\PayRoll\Output
OutputBaseFilename=APLUS_HR_Setup_v1.0
Compression=lzma
SolidCompression=yes
WizardStyle=modern

; --- ไอคอนของตัวติดตั้ง (ถ้ามีไฟล์ .ico ให้ใส่ตรงนี้ ถ้าไม่มีให้ลบบรรทัดนี้ออก) ---
; SetupIconFile=C:\Users\Nitro V15\Desktop\PayRoll\company_logo.ico 

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; --- ไฟล์หลัก .exe ---
Source: "C:\Users\Nitro V15\Desktop\PayRoll\dist\APLUS_HR_Payroll\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; --- กวาดทุกไฟล์และโฟลเดอร์ใน dist ไปด้วย (รวม resources, pdf, library ต่างๆ) ---
Source: "C:\Users\Nitro V15\Desktop\PayRoll\dist\APLUS_HR_Payroll\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; หมายเหตุ: recursesubdirs คือเอาโฟลเดอร์ย่อยไปด้วย (เช่น resources)

[Icons]
; --- สร้าง Shortcut ใน Start Menu ---
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

; --- สร้าง Shortcut บน Desktop ---
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; --- รันโปรแกรมทันทีหลังติดตั้งเสร็จ ---
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent