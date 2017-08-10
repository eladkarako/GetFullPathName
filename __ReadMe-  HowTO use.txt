1. add module STD_in_out.bas to vb project.
2. work, use the ReadStdIn and WriteStdOut as you needed.
3. save project, make exe.
4. patch (relink) project exe with shortcut ("C:\Program Files\Microsoft Visual Studio\vb98\LINK.EXE" /EDIT /SUBSYSTEM:CONSOLE <yourfile.exe>)
5. your exe will work now.