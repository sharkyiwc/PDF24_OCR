'==========================================================================
'
' NAME: ScanPDF2OCR.vbs
'
' AUTHOR: Sharkyiwc
' DATE  : 23.12.2021
'
' COMMENT: OCR a PDF file from a given directory using PDF24.
' Doc : https://creator.pdf24.org/manual/10/#command-line
'==========================================================================
Dim ProgFilesx86, ProgFiles, PDF24
Set fso = CreateObject("Scripting.FileSystemObject")
Set WSHShell = WScript.CreateObject("wscript.shell")
OCRlang = "fra" 'Set lang for OCR: Lang list https://creator.pdf24.org/tesseract/4.0/traindata/local-list.txt
PDF24OCRArg = " -language " & OCRlang & " -dpi 200 -autoRotatePages -skipFilesWithText -skipPagesWithText "
workpath = fso.GetAbsolutePathName(".")
workpath2 = workpath&"\"
Set Afolder = fso.GetFolder (workpath)
Set AllFiles = Afolder.Files

ProgFilesx86 = WSHShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
ProgFiles = WSHShell.ExpandEnvironmentStrings("%ProgramFiles%")

'Check PDF42 is installed
If (fso.FileExists(ProgFilesx86 & "\PDF24\pdf24-Ocr.exe")) Then
    PDF24EXE = """" & ProgFilesx86 & "\PDF24\pdf24-Ocr.exe" & """"
    ProgFilesPath = """" & ProgFiles & "\PDF24" & """"
End if
If (fso.FileExists(ProgFiles & "\PDF24\pdf24-Ocr.exe")) Then
    PDF24EXE = """" & ProgFiles & "\PDF24\pdf24-Ocr.exe" & """"
    ProgFilesPath = """" & ProgFilesx86 & "\PDF24" & """"
End if
if PDF24EXE = Empty Then
    Message = MsgBox ("PDF24 n'est pas install"&chr(233)&" sur ce poste. Veuillez assurez que PDF24 soit install"&chr(233)&" sur votre PC", 64, "Pr"&chr(233)&"requis manquant")
    WScript.Quit 
end if

For Each Afile In AllFiles
    If UCase(fso.GetExtensionName(Afile.Name))="PDF" Then
        NotMakeOCR = "NO"
        filename = fso.GetFileName (ucase(Afile))
        OCRfile = left(right(filename,8),4)
        filenameWithoutExt = Left(filename,Len(filename) - 4)
        NewFileName = filenameWithoutExt & "_OCR.pdf"
        InpoutFile = """" & workpath2&filename & """"
        OutputFile = """" & workpath2&NewFileName & """"
        if Afile.Name=NewFileName Then 'check if output file is allready present
            NotMakeOCR = "YES"
        End if
        'WScript.Echo NotMakeOCR & " Etape 1"
        if fso.FileExists(workpath2&filenameWithoutExt & "_OCR.pdf") Then 'check if PDF already OCR
            NotMakeOCR = "YES"
        End if
        if NotMakeOCR <> "YES" then
            if not (fso.FileExists(OutputFile)) then
                if not OCRfile = "_OCR" Then
                    WSHShell.Run "cmd /C cd &" & PDF24EXE & " -outputFile " & OutputFile & PDF24OCRArg & InpoutFile, 0, true
                End if
            End if
        End if
    End if
Next

AskDeleteFile = Msgbox("Voulez-vous supprimer les fichiers originaux apr"&chr(232)&"s l'OCR?" & Chr(13) & Chr(10) & "Les fichiers originaux seront supprim"&chr(233)&"s" & Chr(13) & Chr(10) & "Les nouveaux seront renomm"&chr(233)&" comme les originaux.", 4, "Suppression apr"&chr(232)&"s converssion")
If AskDeleteFile=6 Then
    ' If Yes, 1 delete original file
    For Each Afile In AllFiles
        If UCase(fso.GetExtensionName(Afile.Name))="PDF" Then
            filenameToDel = fso.GetFileName (Afile)
            'wscript.Echo filenameToDel
            if not Ucase(left(right(filenameToDel,8),4))="_OCR" then
                'wscript.Echo Ucase(left(right(filenameToDel,8),4))
                fso.DeleteFile workpath2&filenameToDel
            End if
        End if
    Next
    '2. rename OCR file to original file
    For each Afile In AllFiles
        filenameToRename = fso.GetFileName (Afile)
        If instr(Afile, "_OCR") > 0 THEN
            Afile.name = replace(Afile.name, "_OCR.", ".")
        End IF
    Next

End If

'End
Set fso = Nothing
Set WSHShell = Nothing
WScript.Quit
