'This vb script will convert your all .doc or .docx files to .pdf in one go.'
'The .docx/doc files will be remain same.'
'More details are in README.md file.

Set FileSysOb = CreateObject("Scripting.FileSystemObject")
For p= 0 To WScript.Arguments.Count -1
   docPath = WScript.Arguments(p)
   docPath = FileSysOb.GetAbsolutePathName(docPath)
   If LCase(Right(docPath, 4)) = ".doc" Or LCase(Right(docPath, 5)) = ".docx" Then
      Set objWord = CreateObject("Word.Application")
      pdfPath = FileSysOb.GetParentFolderName(docPath) & "\" & _
	FileSysOb.GetBaseName(docpath) & ".pdf"
      objWord.Visible = False
      Set objDoc = objWord.documents.open(docPath)
      objDoc.saveas pdfPath, 17
      objDoc.Close
      objWord.Quit   
   End If   
Next