Sub Creation_PNG()
    Dim Pict As Picture
    Dim chrt As ChartObject
    Dim Dossier As String
    Dim CheminQR As String
    Dim i As Integer
    Dim fs As Object

    i = 1
    Dossier = "QRCODES"
    
    Worksheets("Liens").Range("A2:B65536").ClearContents
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    CheminQR = ThisWorkbook.Path & "\" & Dossier
    If Not fs.FolderExists(CheminQR) Then
        fs.CreateFolder CheminQR
    End If
    Set fs = Nothing

    If Dir(ThisWorkbook.Path & "\" & Dossier & "\" & "*.*") <> "" Then
        Kill ThisWorkbook.Path & "\" & Dossier & "\" & "*.*"
    End If

    For Each Pict In ThisWorkbook.Sheets("Createur-QR-Codes").Pictures
        Pict.CopyPicture
        W = Pict.Width
        H = Pict.Height
        Set chrt = ActiveSheet.ChartObjects.Add(0, 0, W, H)
        chrt.Chart.Paste
        chrt.Chart.Export Filename:=ThisWorkbook.Path & "\" & Dossier & "\" & Pict.Name & ".png", FilterName:="PNG"
        CheminPublipostage = Replace(ThisWorkbook.Path & "\" & Dossier & "\" & Pict.Name & ".png", "\", "\\")
        i = i + 1
        ThisWorkbook.Sheets("Liens").Cells(i, 1).Value = CheminPublipostage
        ThisWorkbook.Sheets("Liens").Cells(i, 2).Value = Pict.Name
        chrt.Delete
    Next Pict
End Sub

Sub MAJ_Etiquettes()
    Dim CheminClasseur As String
    Dim CheminPublipostage As String
    Dim NomClasseur As String
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim Tableau1() As String

    NomClasseur = ThisWorkbook.Name
    CheminClasseur = ThisWorkbook.Path & "\" & NomClasseur
    NomPublipostage = "etiquettes.docx"
    
    Tableau1 = Split(CheminClasseur, "\")
    ReDim Preserve Tableau1(UBound(Tableau1) - 1)
    CheminPublipostage = Join(Tableau1, "\") & "\" & NomPublipostage

    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    Set WordDoc = WordApp.Documents.Open(CheminPublipostage)
    WordApp.Visible = True

    WordDoc.MailMerge.MainDocumentType = wdFormLetters
    WordDoc.MailMerge.OpenDataSource Name:=CheminClasseur, _
        ConfirmConversions:=False, ReadOnly:=False, LinkToSource:=True, _
        AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:="", _
        WritePasswordDocument:="", WritePasswordTemplate:="", Revert:=False, _
        Format:=wdOpenFormatAuto, Connection:= _
        "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=" & CheminClasseur & ";Mode=Read;Extended Properties=""HDR=YES;IMEX=1;"";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=3" _
        , SQLStatement:="SELECT * FROM `Liens$`", SQLStatement1:="", SubType:= _
        wdMergeSubTypeAccess
        
    With WordDoc.MailMerge
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        With .DataSource
            .FirstRecord = wdDefaultFirstRecord
            .LastRecord = wdDefaultLastRecord
        End With
        .Execute Pause:=False
    End With
    
    Set WordDoc = Nothing
    Set WordApp = Nothing

    Message1 = "Les étiquettes ont été générées. Rendez-vous dans le fichier Word ouvert 'Lettres types1.docx' puis actionnez les touches 'Ctrl+A' puis la touche 'F9' afin de mettre à jour tous les champs."
    rep = MsgBox(Message1, 64, "Création des étiquettes")
End Sub
