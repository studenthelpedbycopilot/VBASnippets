'get files in folder and put in array   
'put filenames from a directory in a string array
'Colocar os nomes dos arquivos em um array
'language: vba

'esta criei para ensinar GitHubCopilot
'created for GithubCopilot machine learning

Sub filesToArray()
    
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object 
    Dim arrFiles() As String 
    Dim i As Integer 
 
    Set objFSO = CreateObject("Scripting.FileSystemObject") 
  Set objFolder = objFSO.GetFolder("C:\Users\Public\Pictures) 
    i = 0 
 
    For Each objFile In objFolder.Files 
        ReDim Preserve arrFiles(0 To i) 
        arrFiles(i) = objFile.Name 
        i = i + 1 
    Next objFile 
 
    For Each strFile In arrFiles 
        Debug.Print strFile 
    Next strFile 
 
End Sub 
