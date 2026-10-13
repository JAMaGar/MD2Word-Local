; Script do Inno Setup 6 para o MD2Word Local
; Compilado automaticamente pelo build.bat (ou abra no Inno Setup e clique em Compile).

#define AppNome "MD2Word Local"
#define AppVersao "1.0.0"
#define AppExe "MD2Word.exe"

[Setup]
AppId={{6F3B2C1E-8A4D-4E7B-9C2F-5D1A7E3B9C40}
AppName={#AppNome}
AppVersion={#AppVersao}
AppVerName={#AppNome} {#AppVersao}
AppPublisher=Uso pessoal
DefaultDirName={autopf}\MD2Word Local
DefaultGroupName={#AppNome}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=instalador
OutputBaseFilename=MD2Word_Local_Setup_{#AppVersao}
SetupIconFile=assets\md2word.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppNome}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ChangesAssociations=yes
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "ptbr"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Área de Trabalho"; GroupDescription: "Atalhos:"
Name: "menucontexto"; Description: "Adicionar ""Converter para Word (.docx)"" ao clicar com o botão direito em arquivos .md"; GroupDescription: "Integração com o Explorer:"
Name: "abrircom"; Description: "Mostrar o MD2Word em ""Abrir com"" para arquivos .md"; GroupDescription: "Integração com o Explorer:"

[Files]
Source: "dist\MD2Word\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "LEIA-ME.docx"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "exemplos\*"; DestDir: "{app}\exemplos"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{autoprograms}\{#AppNome}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppNome}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Registry]
; Menu de contexto do Explorer (no Windows 11 aparece em "Mostrar mais opções")
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.md\shell\MD2Word.docx"; ValueType: string; ValueName: ""; ValueData: "Converter para Word (.docx)"; Flags: uninsdeletekey; Tasks: menucontexto
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.md\shell\MD2Word.docx"; ValueType: string; ValueName: "Icon"; ValueData: """{app}\{#AppExe}"",0"; Tasks: menucontexto
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.md\shell\MD2Word.docx\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExe}"" --rapido docx ""%1"""; Tasks: menucontexto
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.markdown\shell\MD2Word.docx"; ValueType: string; ValueName: ""; ValueData: "Converter para Word (.docx)"; Flags: uninsdeletekey; Tasks: menucontexto
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.markdown\shell\MD2Word.docx"; ValueType: string; ValueName: "Icon"; ValueData: """{app}\{#AppExe}"",0"; Tasks: menucontexto
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.markdown\shell\MD2Word.docx\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExe}"" --rapido docx ""%1"""; Tasks: menucontexto
; "Abrir com"
Root: HKA; Subkey: "Software\Classes\Applications\{#AppExe}"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "{#AppNome}"; Flags: uninsdeletekey; Tasks: abrircom
Root: HKA; Subkey: "Software\Classes\Applications\{#AppExe}\SupportedTypes"; ValueType: string; ValueName: ".md"; ValueData: ""; Tasks: abrircom
Root: HKA; Subkey: "Software\Classes\Applications\{#AppExe}\SupportedTypes"; ValueType: string; ValueName: ".markdown"; ValueData: ""; Tasks: abrircom
Root: HKA; Subkey: "Software\Classes\Applications\{#AppExe}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExe}"" ""%1"""; Tasks: abrircom
Root: HKA; Subkey: "Software\Classes\.md\OpenWithProgids"; ValueType: string; ValueName: "MD2Word.md"; ValueData: ""; Flags: uninsdeletevalue; Tasks: abrircom
Root: HKA; Subkey: "Software\Classes\.markdown\OpenWithProgids"; ValueType: string; ValueName: "MD2Word.md"; ValueData: ""; Flags: uninsdeletevalue; Tasks: abrircom
Root: HKA; Subkey: "Software\Classes\MD2Word.md"; ValueType: string; ValueName: ""; ValueData: "Documento Markdown"; Flags: uninsdeletekey; Tasks: abrircom
Root: HKA; Subkey: "Software\Classes\MD2Word.md\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExe}"",0"; Tasks: abrircom
Root: HKA; Subkey: "Software\Classes\MD2Word.md\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExe}"" ""%1"""; Tasks: abrircom

[Run]
Filename: "{app}\{#AppExe}"; Description: "Abrir o {#AppNome} agora"; Flags: nowait postinstall skipifsilent
