sFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWord = CreateObject("Word.Application")
oWord.Visible = False

Set oFolder2 = oFSO.GetFolder(sFolder)


ConvertFolder(oFolder2)
oWord.Quit

Sub ConvertFolder(oFldr)
  For Each oFile In oFldr.Files
    If LCase(oFSO.GetExtensionName(oFile.Name)) = "docx" Then
        Set oDoc = oWord.Documents.Open(oFile.path)
        Str = left(oFile,instr(1,oFile,".")-1) 
        oWord.ActiveDocument.SaveAs Str & ".pdf", 17
        oDoc.Close
    End If
Next

End Sub
