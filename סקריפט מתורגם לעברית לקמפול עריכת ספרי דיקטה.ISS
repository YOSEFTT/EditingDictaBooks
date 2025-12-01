; ---------------------------
; קובץ התקנה עבור "עריכת ספרי דיקטה"
; כאן יש הערות שמסבירות כל שורה
; ---------------------------

; --- הגדרות ראשיות (משתנים שנשתמש בהם בהמשך) ---
#define MyAppName "עריכת ספרי דיקטה"
; שם האפליקציה שמופיע במסכי ההתקנה ובקיצורי דרך
#define MyAppVersion "3.7.1"
; מספר גרסה של האפליקציה
#define MyAppPublisher "מתנדבי אוצריא"
; שם המוציא לאור שמופיע בהתקנה
#define MyAppURL "https://github.com/YOSEFTT/EditingDictaBooks/releases"
; אתר הבית/עדכונים
#define MyAppExeName "עריכת ספרי דיקטה.exe"
; שם קובץ ההרצה שמותקן אצל המשתמש
#define MyAppSource "D:\ADMIN\Downloads\output\עריכת ספרי דיקטה.exe"
; מיקום הקובץ שאתה מצרף להתקנה

#expr Exec('pyinstaller', '--noconfirm --onefile --windowed --distpath "D:\ADMIN\Downloads\output" --icon "D:\ADMIN\Downloads\מידע לקימפול\לעריכת ספרי דיקטה\edit otzaria.ico" --name "עריכת ספרי דיקטה" --clean --version-file "D:\ADMIN\Downloads\מידע לקימפול\לעריכת ספרי דיקטה\version_info.txt" "C:\Users\ADMIN\EditingDictaBooks\EditingDictaBooks.py"', '', 1)
; אם INNO מותקן בפרוגרם פיילס שאין שם הרשאות כתיבה.
; #expr Exec('mkdir', 'C:\temp\build', '', 1, 1)
; #expr Exec('pyinstaller', '--workpath "C:\temp\build" --noconfirm --onefile --windowed --distpath "D:\ADMIN\Downloads\output" --icon "D:\ADMIN\Downloads\מידע לקימפול\לעריכת ספרי דיקטה\edit otzaria.ico" --name "EditingDictaBooks" --clean --version-file "D:\ADMIN\Downloads\מידע לקימפול\לעריכת ספרי דיקטה\version_info.txt" "C:\Users\ADMIN\EditingDictaBooks\EditingDictaBooks.py"', 'C:\temp', 1, 0)

; #expr Exec('signtool sign /f "C:\Users\ADMIN\certificate.pfx" /fd SHA256 /tr http://timestamp.digicert.com /v D:\ADMIN\Downloads\output\עריכת ספרי דיקטה.exe', '', 0)
; #expr Exec('signtool', 'sign /f "C:\Users\ADMIN\certificate.pfx" /fd SHA256 /tr http://timestamp.digicert.com /v "D:\ADMIN\Downloads\output\עריכת ספרי דיקטה.exe"', '', 1, 0)
; עם CMD 
#expr Exec('cmd', '/c signtool sign /f "C:\Users\ADMIN\certificate.pfx" /fd SHA256 /td SHA256 /a /tr http://timestamp.digicert.com /v "D:\ADMIN\Downloads\output\עריכת ספרי דיקטה.exe"', '', 1, 0)

; --- סקשן ההתקנה ---
[Setup]
AppId={{592C8ED4-7248-4CBA-81A8-1132FC786F65}
              ; מזהה ייחודי לאפליקציה, לא לשנות בין גרסאות
AppName={#MyAppName}
                     ; שם האפליקציה
AppVersion={#MyAppVersion}
               ; מספר גרסה
;AppVerName={#MyAppName} {#MyAppVersion}                   ; ניתן להציג גם שם+גרסה יחד (כאן מושבת)
AppPublisher={#MyAppPublisher}
           ; שם המוציא לאור
AppPublisherURL={#MyAppURL}
              ; קישור שיוצג למשתמש
AppSupportURL={#MyAppURL}
                ; קישור לעזרה/תמיכה
AppUpdatesURL={#MyAppURL}
                ; קישור לעדכונים
DefaultDirName={autopf}\{#MyAppName}
     ; תיקיית ברירת מחדל להתקנה (Program Files או AppData)
UninstallDisplayIcon={app}\{#MyAppExeName}
                 ; האייקון שיוצג בהסרת התקנה
ArchitecturesAllowed=x64compatible
                   ; התקנה מותרת רק על 64bit
ArchitecturesInstallIn64BitMode=x64compatible
; ההתקנה מתבצעת ב־Program Files (64bit) אם אפשר
DisableProgramGroupPage=yes
              ; לא מציג עמוד ליצירת קבוצת תפריט התחלה
PrivilegesRequired=lowest
                ; לא דורש אדמין (יתקין ב־AppData אם אין הרשאות)
PrivilegesRequiredOverridesAllowed=dialog
                  ; המשתמש יכול לבחור אם לעלות עם אדמין
OutputBaseFilename=עריכת ספרי דיקטה תוכנת התקנה
            ; שם קובץ ההתקנה שיווצר (Setup.exe)
SetupIconFile=D:\ADMIN\Downloads\מידע לקימפול\לעריכת ספרי דיקטה\edit otzaria.ico
    ; אייקון קובץ ההתקנה
SolidCompression=yes
                     ; דחיסה מקסימלית
WizardStyle=modern
                       ; עיצוב ויזארד מודרני
; 'signtool', mycert 'sign /f "C:\Users\ADMIN\certificate.pfx" /fd SHA256 /td SHA256 /v /a /tr http://timestamp.digicert.com' $f
; חתימה דיגיטלית על תוכנת ההתקנה
; SignTool=mycert=sign /f "C:\Users\ADMIN\certificate.pfx" /fd SHA256 /td SHA256 /v /a /tr http://timestamp.digicert.com $f
; SignTool=mycert=sign /f "C:\Users\ADMIN\certificate.pfx" /fd SHA256 /td SHA256 /v /a /tr http://timestamp.digicert.com $f
; SignedUninstaller=yes
; 'cmd', '/c signtool sign /f "C:\Users\ADMIN\certificate.pfx" /fd SHA256 /td SHA256 /a /tr http://timestamp.digicert.com $f
; SignTool=mycert=cmd /c signtool sign /f "C:\Users\ADMIN\certificate.pfx" /fd SHA256 /td SHA256 /v /a /tr http://timestamp.digicert.com $f
; SignTool=mycert="C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe" sign /f "C:\Users\ADMIN\certificate.pfx" /fd SHA256 /td SHA256 /v /a /tr http://timestamp.digicert.com $f

; --- שפות התקנה ---
[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "hebrew"; MessagesFile: "compiler:Languages\Hebrew.isl"

; --- משימות נוספות ---
[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
; יוצר אייקון על הדסקטופ אם המשתמש בוחר בזה

; --- קבצים שמתווספים להתקנה ---
[Files]
Source: "{#MyAppSource}"; DestDir: "{app}"; Flags: ignoreversion
; הקובץ שיצורף להתקנה (נלקח מהמחשב שלך, ויועתק ל־{app})

; --- קיצורי דרך ---
[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
        ; קיצור דרך בתפריט ההתחלה
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
 ; קיצור דרך בשולחן עבודה

; --- הרצה בסיום ---
[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
; מציע להריץ את התוכנה מיד בסיום ההתקנה
