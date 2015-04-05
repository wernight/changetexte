Attribute VB_Name = "Dir"


Public Function DirFichier(Répertoire As String, Nommé As String, Attrib As Integer)
    DirFichier = Répertoire & Dir(Répertoire & "\" & Nommé, Attrib)
End Function
