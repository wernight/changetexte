Attribute VB_Name = "Dir"


Public Function DirFichier(R�pertoire As String, Nomm� As String, Attrib As Integer)
    DirFichier = R�pertoire & Dir(R�pertoire & "\" & Nomm�, Attrib)
End Function
