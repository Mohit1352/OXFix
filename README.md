# OXFix
A tool that tries to fix a PPTX/DOCX file when it cannot be read or fixed by presentation software other than PowerPoint/Word.
(I faced an issue with opening a few properly downloaded ppt files that failed to open in software like LibreOffice, and only opened in PowerPoint/Word. This is my solution of the issue.)

The tool is built using Python3, and makes use of os,shutil and sys modules. It replaces all PPTX/DOCX files in the given directories with fixed versions of the same file.

A. USING THE PYTHON FILE:

The ways this program can be used:<br>
  1. Plain run from the shell:<br>
     a. With no added directory names:<br><br>
            <code>python3 OXFix.py</code><br><br>
        This scans the current working directory for any ppt files, and fixes them.<br>
     b. With folder names in the current working directory having ppt files:<br><br>
            <code>python3 OXFix.py Folder1 Folder2 "Folder 3" ...</code><br><br>
        This scans the current working directory, and also the folders mentioned.<br>

  2. With a pipe from another command that outputs names of directories.<br><br>
          <code>ls | python3 OXFix.py</code><br>
          <code>ls | awk <do some filtering, analyze some pattern> | python XFix.py</code><br>
        (and others...)<br><br>

Note: The programs checks sys.stdin to receive input from the pipes. In cases where there is no piping, you might have to press the Enter key to make the program fix files instead of waiting for input in stdin.<br>

B. USING THE FILE WITH NO EXTENSION:

A symlink of the file can be added to /usr/bin and will run with the command <code>OXFix</code>.
The syntax is same as the python file syntax, just without python.
