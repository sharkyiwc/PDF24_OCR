'==========================================================================
'
' NAME: ScanPDF2OCR.vbs
'
' AUTHOR: Sharkyiwc
' DATE  : 18.10.2021
'
' COMMENT: OCR a PDF file from a given directory using PDF24.
' Doc : https://creator.pdf24.org/manual/10/#command-line
'==========================================================================
Dim ProgFilesx86, ProgFiles, PDF24
Set fso = CreateObject("Scripting.FileSystemObject")
Set WSHShell = WScript.CreateObject("wscript.shell")
workpath = fso.GetAbsolutePathName(".")
workpath2 = workpath&"\"
Set Afolder = fso.GetFolder (workpath)
Set AllFiles = Afolder.Files

ProgFilesx86 = WSHShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
ProgFiles = WSHShell.ExpandEnvironmentStrings("%ProgramFiles%")

'Check PDF42 installed
'PDF24EXE = ProgFiles & "\PDF24\pdf24-Ocr.exe" 'For try
If (fso.FileExists(ProgFilesx86 & "\PDF24\pdf24-Ocr.exe")) Then
    PDF24EXE = ProgFilesx86 & "\PDF24\pdf24-Ocr.exe"
End if
If (fso.FileExists(ProgFiles & "\PDF24\pdf24-Ocr.exe")) Then
    PDF24EXE = ProgFiles & "\PDF24\pdf24-Ocr.exe"
End if
'WScript.Echo PDF24EXE 'for control
if PDF24EXE = Empty Then
    Message = MsgBox ("PDF24 n'est pas install"&chr(233)&" sur ce poste. Veuillez assurez que PDF24 soit install"&chr(233)&" sur votre PC", 64, "Pr"&chr(233)&"requis manquant")
    WScript.Quit 
end if

For Each Afile In AllFiles
    If UCase(fso.GetExtensionName(Afile.Name))="PDF" Then
        filename = fso.GetFileName (ucase(Afile))
        OCRfile = left(right(filename,8),4)
        filenameWithoutExt = Left(filename,Len(filename) - 4)
        NewFileName = filenameWithoutExt & "_OCR.pdf"
        PDF24OCRArg = " -language fra -dpi 200 -autoRotatePages -skipFilesWithText -skipPagesWithText "
        'WScript.Echo filename
        'WScript.Echo filenameWithoutExt
        'WScript.Echo NewFileName
        InpoutFile = """" & workpath2&filename & """"
        OutputFile = """" & workpath2&NewFileName & """"

        if not OCRfile = "_OCR" Then
            WScript.Echo """" & PDF24EXE & " -outputFile " & OutputFile & PDF24OCRArg & InpoutFile & """"
            'WSHShell.Run """" & PDF24EXE & " -outputFile " & OutputFile & PDF24OCRArg & InpoutFile & """"
           ' OCR
        End if
    End if
Next

'End
Set fso = Nothing
Set WSHShell = Nothing
WScript.Quit 
