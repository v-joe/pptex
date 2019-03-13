# AppleScript part of pptex
# https://github.com/v-joe/pptex
# License: GPLv3
on ppLauncher(dirname)

# ------------ Make adjustments for your system here --------------
#
# Edit the following variable to pick your editor for LaTeX
# TextMate seems to be working well for the pptex macro and is recommended.
# TextEdit, SubEthaEdit, and Sublime do not seem to work: lsof cannot detect
# open file status with these three editors, and the pptex macro won't be
# able to detect when editing the LaTeX snippet is done.
# Atom does show open file status with lsof, but the during tests, the
# pptex macro ran into errors.
set Editor to "TextMate"
# Edit the following variable to point to the full path of pdflatex
# E.g., for MacPorts: "/opt/local/bin/pdflatex"
# or for Fink: "/sw/bin/pdflatex"
# or for MacTeX: "/Library/TeX/texbin/pdflatex"
set pdfLatex to "/opt/local/bin/pdflatex"
#
# -----------------------------------------------------------------


#
# main part of the script
#
set filename to dirname & "/input.tex"
set pFile to (POSIX file filename)
set lsofStr to "sleep 0.01;lsof " & filename & "|grep -i " & Editor & "|wc -l|tr -d ' '"
tell application Editor
	activate
	open file pFile
	# wait until editor has opened the file
	repeat
		do shell script lsofStr
		if the result is "1" then exit repeat
	end repeat
	# wait until editor has closed the file
	repeat
		do shell script lsofStr
		if the result is "0" then exit repeat
	end repeat
end tell
# run pdflatex
do shell script "cd " & dirname & ";" & pdfLatex & " input.tex"
end ppLauncher
