# WordToPdfConverter
This vb script project will convert your all .doc or .docx files to .pdf in one go.
Also it can be placed into the windows sendto menu to convert files without running a command.
The .docx/doc files will be remain same.
The new files will be created with the same name with .pdf extension.
It can convert multiple files at one go.
If you try to convert the same file twice, it will replace the old pdf with the new one.

Steps:
1. Download the project and cd into the project directory.
2. Now Create a shortcut of the file (eg. Doc2Pdf.vbs).
3. Now copy the shortcut and cd into "C:\Users\%username%\AppData\Roaming\Microsoft\Windows\SendTo" directory.
4. Now paste the shortcut into this directory.
5. Now right click on the shortcut and go to the properties.
6. Now in target section paste below commandNote that vb script file location can be different in your case.)
7. C:\Windows\System32\wscript.exe "X:\DocToPdf.vbs"
8. Additionally you can add a icon for your shortcut give a look and recognize it.
9. Now click on Apply then OK and you are good to go.
10. Now to convert a word file right click on the file, if there are multiple file then  select them and right click then choose send to menu and select the shortcut name.
11. Once you selected the shortcut it will start creating the .pdf files.
