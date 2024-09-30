'pptex for Microsoft Powerpoint for Mac
'https://github.com/v-joe/pptex
'License: GPLv3
Sub pptex()

Dim tex As String
Dim LF As String

LF = Chr(10)
tex = "\documentclass[varwidth=true,border=2pt]{standalone}" & LF & _
    "\usepackage{color}" & LF & _
    "\begin{document}" & LF & _
    LF & LF & LF & LF & LF & LF & _
    "\end{document}"

Dim tmpdir As String
Dim file As String
Dim dtex As String
Dim ntex As String
Dim wscale As Single


tmpdir = ""
On Error Resume Next
' MacScript is deprecated in PowerPoint 2016, but running mktemp with it
' will create a sub directory in PowerPoint's temporary space it can read from
' and write to
tmpdir = MacScript("do shell script " & Chr(34) & "/usr/bin/mktemp -d 2>/dev/null" & Chr(34))
On Error GoTo 0
If tmpdir = "" Then
    MsgBox "Cannot create temporary directory", Title:="pptex error"
    Exit Sub
End If

file = tmpdir & "/input.tex"
pdf = tmpdir & "/input.pdf"

' Check if there is previous LaTeX code as alternative text in the
' selected shape
dtex = ""
On Error Resume Next
dtex = ActiveWindow.Selection.ShapeRange.Item(1).AlternativeText
On Error GoTo 0
If InStr(dtex, "\documentclass") > 0 Then
    tex = dtex
End If

' Write LaTeX input
Open file For Output As #1
Print #1, tex;
Close #1

' Launch editor and pdfLaTeX
On Error GoTo pdfLaTeXerror
pdfLaTeXOutput = AppleScriptTask("pptexLaunch.applescript", "ppLauncher", tmpdir)
On Error GoTo 0

' If input wasn't changed, quit here
Open file For Input As #1
ntex = Input(LOF(1), 1)
Close #1
If tex = ntex Then GoTo texcontinue

' Insert generated pdf
If dtex = "" Then
    On Error GoTo texerror
    Set x = ActiveWindow.Selection.SlideRange.Shapes.AddPicture( _
        fileName:=pdf, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, Left:=200, Top:=100)
    x.ScaleWidth 2, msoTrue
    x.ScaleHeight 2, msoTrue
Else
' In case of selected pptex-generated shape, replace it with new one
' at same position and same relative scale factor
    Set s = ActiveWindow.Selection.ShapeRange.Item(1)
    On Error GoTo texerror
    Set x = ActiveWindow.Selection.SlideRange.Shapes.AddPicture( _
        fileName:=pdf, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, Left:=s.Left, Top:=s.Top)
    wscale = s.Width
    s.ScaleWidth 1, msoTrue
    wscale = wscale / s.Width
    x.ScaleWidth wscale, msoTrue
    x.ScaleHeight wscale, msoTrue
    s.Delete
End If
x.Select
' Store LaTeX source code as alternative text of pdf figure
x.AlternativeText = ntex

' Remove temporary sub directory
texcontinue:
On Error Resume Next
MacScript "do shell script " & Chr(34) & "rm -r " & tmpdir & " 2>/dev/null" & Chr(34)
On Error GoTo 0

Exit Sub

texerror:
MsgBox "Error in LaTeX input", vbCritical, "pptex"
GoTo texcontinue

' create message box with pdfLaTeX log in case of error
pdfLaTeXerror:
On Error Resume Next
MsgBox "pdfLaTeX error:" & LF & MacScript("do shell script " & Chr(34) & "tail -25 " & tmpdir & "/input.log" & Chr(34))
GoTo texcontinue

End Sub
