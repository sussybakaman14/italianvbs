Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")

' Creazione di una nuova cartella
If Not fso.FolderExists("C:\TestVBScript") Then
    fso.CreateFolder("C:\TestVBScript")
    MsgBox "Cartella creata con successo!", vbInformation, "Info"
End If

' Creazione di un nuovo file e scrittura di contenuto
Set file = fso.CreateTextFile("C:\TestVBScript\esempio.txt", True)
file.WriteLine("Questo è un file creato con VBScript!")
file.Close

MsgBox "File creato e contenuto aggiunto!", vbInformation, "Info"
