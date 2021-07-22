# PPTFix
A tool that tries to fix a PPT file when it cannot be read or fixed by presentation software other than PowerPoint.
(I faced an issue with opening a few properly downloaded ppt files that failed to open in software like LibreOffice, and only opened in PowerPoint. This is my solution of the issue.)

The tool is built using Python, and makes use of os,shutil and sys modules. It replaces all PPT files in the given directories with fixed versions of the same file.

The ways this program can be used:<br>
  1. Plain run from the shell:<br>
     a. With no added directory names:<br><br>
            <code>python PPTFix.py</code><br><br>
        This scans the current working directory for any ppt files, and fixes them.<br>
     b. With folder names in the current working directory having ppt files:<br><br>
            <code>python PPTFix.py Folder1 Folder2 "Folder 3" ...</code><br><br>
        This scans the current working directory, and also the folders mentioned.<br>

  2. With a pipe from another command that outputs names of directories.<br><br>
          <code>ls folder | python PPTFix.py</code><br>
          <code>ls | awk <do some filtering, analyze some pattern> | python PPTFix.py</code><br>
        (and others...)<br><br>

Note: The programs checks sys.stdin to receive input from the pipes. In cases where there is no piping, you might have to press the Enter key to make the program fix files instead of waiting for input in stdin.<br>
