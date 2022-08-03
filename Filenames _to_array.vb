'put filenames from a directory in a string array
'Colocar os nomes dos arquivos em um array
'language: vba

Sub ListarArquivosnapasta() 
 
  Dim ObjScriptArquivo As Object 
  Dim ObjPasta As Object 
  Dim ObjArquivo As Object 
  Dim strpathFiles() As String 
  Dim i As Integer 
 
  i = 0 
   
  Set ObjScriptArquivo = CreateObject("Scripting.FileSystemObject") 
 
  Set ObjPasta = ObjScriptArquivo.GetFolder("C:\Users\Public\Pictures")

  Debug.Print ObjPasta.Path

  For Each ObjArquivo In ObjPasta.Files

    ReDim Preserve strpathFiles(0 To i)
    strpathFiles(i) = ObjPasta.Path & "\" & ObjArquivo.Name

    Cells(i + 1, 1).Value = ObjArquivo.Name
    Cells(i + 1, 2).Value = strpathFiles(i)

    i = i + 1

  Next ObjArquivo

End Sub
'Language: vba
