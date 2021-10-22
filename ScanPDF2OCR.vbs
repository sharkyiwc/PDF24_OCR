'==========================================================================
'
' NAME: ScanPDF2OCR.vbs
'
' AUTHOR: Sharkyiwc
' DATE  : 22.10.2021
'
' COMMENT: OCR a PDF file from a given directory using PDF24.
' Doc : https://creator.pdf24.org/manual/10/#command-line
'==========================================================================
Dim ProgFilesx86, ProgFiles, PDF24
Set fso = CreateObject("Scripting.FileSystemObject")
Set WSHShell = WScript.CreateObject("wscript.shell")
OCRlang = "fra" 'Set lang for OCR
' Lang list https://creator.pdf24.org/tesseract/4.0/traindata/local-list.txt
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
End if
If (fso.FileExists(ProgFiles & "\PDF24\pdf24-Ocr.exe")) Then
    PDF24EXE = """" & ProgFiles & "\PDF24\pdf24-Ocr.exe" & """"
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
                    'OCR PDF files
                    'WScript.Echo PDF24EXE & " -outputFile " & OutputFile & PDF24OCRArg & InpoutFile 'Only for try/debug
                    WSHShell.Run PDF24EXE & " -outputFile " & OutputFile & PDF24OCRArg & InpoutFile, 1, true
                    WshShell.SendKeys "{ENTER}"
                    'If you want delete source file after ocr, uncomment next line
                    'fso.DeleteFile (InpoutFile)
                End if
            End if
        End if
    End if
Next

'End
Set fso = Nothing
Set WSHShell = Nothing
WScript.Quit
