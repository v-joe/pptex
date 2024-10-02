# pptex
[pptex](https://github.com/v-joe/pptex) is a LaTeX Add-In for
PowerPoint for Mac allowing for insertion of equations rendered as PDF
with `pdflatex`. The LaTeX source is stored as alternative text with the
equations in the slides, so that running the [pptex](https://github.com/v-joe/pptex)
macro on a selected [pptex](https://github.com/v-joe/pptex) equation
object allows for re-editing of the existing equation starting from the
original LaTeX equation source.

## Installation
Get `Installpptex.tbz` from the [releases](https://github.com/v-joe/pptex/releases)
and run the unpacked installer.

A PowerPoint Add-in file will be stored in `~/Documents/pptex` (OSX might ask for
permission to access the Documents folder) and an
`applescript` to launch an external editor and run pdflatex in
`~/Library/Application Scripts/com.microsoft.PowerPoint`.

## Prerequisites
- A working installation of `pdflatex` including the [standalone package](https://ctan.org/pkg/standalone)
(which often is already included in standard installations). Choices for
`pdflatex` installations include [MacPorts](https://www.macports.org/),
[Fink](http://www.finkproject.org/), and [MacTeX](http://www.tug.org/mactex/).

- An editor to create LaTeX source for equations to be inserted into PowerPoint
presentations. So far, due to the need at the scripting side to detect if the
editor has closed the LaTeX source file, only [TextMate](https://github.com/textmate/textmate)
seems to work reliably with [pptex](https://github.com/v-joe/pptex).
