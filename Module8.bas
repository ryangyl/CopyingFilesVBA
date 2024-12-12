Attribute VB_Name = "Module8"
Sub moveWithErrorHandling()
Dim fso As Object
Dim sourceFolder As String
Dim destinationFolder As String

Set fso = CreateObject("Scripting.FileSystemObject")

sourceFolder = "C:\Users\rg413939\OneDrive - GSK\General - RMCB Forum_Ju&QR\RML\Slide_Deck\*.*"
destinationFolder = "C:\Users\rg413939\OneDrive - GSK\General - RMCB Forum_Ju&QR\RMCB 2025\01 Jan25 RMCB\"

fso.CopyFile sourceFolder, destinationFolder, True
End Sub

