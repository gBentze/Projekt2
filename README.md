
Option Compare Database

'Tabelle Abteilung, Mitkreis, typeFzgGrp
Const dateipfad = "D:\DantenbankAccess"
Const xlBlattName As String = "Tabelle1"

Private Type typeAbteilung
kuerzel As String
lngKuerzel As Long
End Type

Private Type typeMitKreis
mitKreis As String
End Type

Private Type typeFzgGp
fzgGrp As String
End Type

Private Type StammdatenStruc
dKXKennung As String
nachname As String
vorname As String
email As String
mitKreis As String
lngMitKreis As Long
re As String
stammNr As Long
kstelle As String


End Type

'Private VerknuepftMitFzgGrpXLSXdata() As VerknuepftMitFzgGrp
Private AbtXLSXdata() As typeAbteilung
Private MitKreisXLSXdata() As typeMitKreis
Private FzgGrpXLSXdata() As typeFzgGp
Private StammdatenXLSXdata() As StammdatenStruc
'Anzahl Zeilen in Excel Tabelle

Private XLSXmax As Integer

'##############################################################################
'##############################################################################

Private Sub MainDataImport()             

    Call ImportDaten                        
    Call WriteXLSXDaten                      
    Call CloseXLSXApp(True)     
    
End Sub   

'#####################################
'##############################################################################
'##############################################################################

Private Sub ImportDaten()
Dim xlpfad As String
xlpfad = dateipfad & "\Stammdaten.xlsx"
Dim vKuerzel As Variant, vMitKreis, vFzgGrp
Dim vDKXKennung As Variant, vNachname, vVorname, vEmail, vRE, vStammNr, vKostenstelle

Dim i As Integer
Dim iRowS As Integer
Dim iRowL As Long
Dim iCol As Integer
Dim sCol As String

' Verweis auf Excel-Bibliothek muss gesetzt sein
Dim xlsApp As Excel.Application
Dim Blatt As Excel.Worksheet
Dim MsgAntw As Integer
' Konstante: Name des einzulesenden Arbeitsblattes

' Excel vorbereiten
On Error Resume Next
Set xlsApp = GetObject(, "Excel.Application")
If xlsApp Is Nothing Then
    Set xlsApp = CreateObject("Excel.Application")
End If

On Error GoTo 0
' Exceldatei readonly öffnen
xlsApp.Workbooks.Open xlpfad, , True
    ' Erste Zeile wird statisch angegeben
    iRowS = 2
    ' Letzte Zeile auf Tabellenblatt wird dynamisch ermittelt
    iRowL = xlsApp.Worksheets(xlBlattName).Cells(xlsApp.Rows.Count, 1).End(xlUp).Row
                    
    ' Spalte mit Kuerzel einlesen
    'Stammnummer
    sCol = "A"
    vStammNr = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Nachname
    sCol = "B"
    vNachname = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Vorname
    sCol = "C"
    vVorname = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'DKX-Kennung
    sCol = "D"
    vDKXKennung = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Email
    sCol = "E"
    vEmail = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Abteilung
    sCol = "F"
    vKuerzel = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    ' Rechtseinheit
    sCol = "G"
    vRE = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Kostenstelle
    sCol = "H"
    vKostenstelle = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Mitarbeiterkreis
    sCol = "J"
    vMitKreis = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'FzgGrp
    sCol = "M"
    vFzgGrp = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)

    '--------------------
    ' Werte in globaler Variable speichern
    XLSXmax = iRowL - 1
    ReDim AbtXLSXdata(1 To XLSXmax)
    ReDim MitKreisXLSXdata(1 To XLSXmax)
    ReDim FzgGrpXLSXdata(1 To XLSXmax)
    ReDim StammdatenXLSXdata(1 To XLSXmax)
    
    With xlsApp.Worksheets(xlBlattName)
        For i = 1 To XLSXmax
            AbtXLSXdata(i).kuerzel = vKuerzel(i, 1)
            MitKreisXLSXdata(i).mitKreis = vMitKreis(i, 1)
            FzgGrpXLSXdata(i).fzgGrp = vFzgGrp(i, 1)
            StammdatenXLSXdata(i).dKXKennung = vDKXKennung(i, 1)
            StammdatenXLSXdata(i).email = vEmail(i, 1)
            StammdatenXLSXdata(i).nachname = vNachname(i, 1)
            StammdatenXLSXdata(i).vorname = vVorname(i, 1)
            StammdatenXLSXdata(i).re = vRE(i, 1)
            StammdatenXLSXdata(i).kstelle = vKostenstelle(i, 1)
            StammdatenXLSXdata(i).stammNr = vStammNr(i, 1)

        Next i
    End With

Set xlsApp = Nothing


End Sub

Private Sub WriteXLSXDaten()
' Daten aus Variable in Tabelle übertragen
Dim i As Integer
'Dim lngMitKreisID As Long
Dim sSQLAbt As String
Dim sSQLKreis As String
Dim sSQLGrp As String
Dim sSQLMit As String
Dim sSQLStNr As String
Dim sSQLKst As String
Dim sSqlMitGrp As String

Dim abtID As Long, grpID As Long, kreisID As Long, mitID As Long, reID As Long, StammNrID As Long, kstID As Long
Dim MsgAntw As Integer
' Schleife über alle Datensätze in der Variablen
For i = 1 To XLSXmax
    ' SQL-String erstellen und Daten schreiben
    reID = Nz(DLookup("REID", "tblRechtseinheit", "RE = '" & StammdatenXLSXdata(i).re & "'"), 0)
    StammNrID = Nz(DLookup("StammNrID", "tblStammnummer", "StammNr = " & StammdatenXLSXdata(i).stammNr), 0)
    kstID = Nz(DLookup("KstID", "tblKostenstelle", "Kostenstelle = '" & StammdatenXLSXdata(i).kstelle & "'"), 0)
    
    abtID = Nz(DLookup("AbtID", "tblAbteilung", "OrgEh= '" & AbtXLSXdata(i).kuerzel & "'"), 0)
    kreisID = Nz(DLookup("MitKrID", "tblMitKreis", "MitKreis= '" & MitKreisXLSXdata(i).mitKreis & "'"), 0)
    grpID = Nz(DLookup("FzgGrpID", "tblFzgGrp", "FzgGrp= '" & FzgGrpXLSXdata(i).fzgGrp & "'"), 0)
    mitID = Nz(DLookup("MitID", "tblMitarbeiter", "DKXKennung= '" & StammdatenXLSXdata(i).dKXKennung & "'"), 0)
    
    
'    Debug.Print sSQLAbt
'    Debug.Print sSQLKreis
'    Debug.Print sSQLGrp
    
    DoCmd.SetWarnings False


'1.######################################## tblMitKreis #########################################

    If kreisID = 0 Then
    
        sSQLKreis = "INSERT INTO tblMitKreis (MitKreis ) VALUES ('" & MitKreisXLSXdata(i).mitKreis & "');"
        
        DoCmd.RunSQL sSQLKreis
        
        kreisID = Nz(DLookup("MitKrID", "tblMitKreis", "MitKreis= '" & MitKreisXLSXdata(i).mitKreis & "'"))

    End If

'2.######################################## tblfzgGrp ########################################

    If grpID = 0 Then
    
        sSQLGrp = "INSERT INTO tblFzgGrp (FzgGrp) VALUES ('" & FzgGrpXLSXdata(i).fzgGrp & "');"
        
        DoCmd.RunSQL sSQLGrp
        
        grpID = Nz(DLookup("FzgGrpID", "tblFzgGrp", "FzgGrp= '" & FzgGrpXLSXdata(i).fzgGrp & "'"))
        
    End If
'3.######################################## tblRE ############################################

    If reID = 0 Then
    
        sSQLKreis = "INSERT INTO tblRechtseinheit (RE ) VALUES ('" & StammdatenXLSXdata(i).re & "');"
        
        DoCmd.RunSQL sSQLKreis
        
        reID = Nz(DLookup("REID", "tblRechtseinheit", "RE= '" & StammdatenXLSXdata(i).re & "'"))

    End If


'3.######################################## tblAbteilung ####################################

    If abtID = 0 Then
    
         sSQLAbt = "INSERT INTO tblAbteilung (OrgEh ) VALUES ('" & AbtXLSXdata(i).kuerzel & "');"

         DoCmd.RunSQL sSQLAbt

         abtID = Nz(DLookup("AbtID", "tblAbteilung", "OrgEh= '" & AbtXLSXdata(i).kuerzel & "'"))

    End If


'4.######################################## tblMitarbeiter #################################

     ' Abteilungskuerzel mit den Schlüsselwerten in Variablen ersetzen
     
    If mitID = 0 Then
    
         sSQLMit = "INSERT INTO tblMitarbeiter (MitKreisID, Nachname, Vorname, DKXKennung, Email)VALUES(" & kreisID & ", '" & _
         StammdatenXLSXdata(i).nachname & "','" & StammdatenXLSXdata(i).vorname & "','" & StammdatenXLSXdata(i).dKXKennung & _
         "', '" & StammdatenXLSXdata(i).email & "');"
         
        Debug.Print sSQLMit
        
        DoCmd.RunSQL sSQLMit
        
        mitID = Nz(DLookup("MitID", "tblMitarbeiter", "DKXKennung= '" & StammdatenXLSXdata(i).dKXKennung & "'"))
        
    End If

'5.######################################## tblVerkMitGrp ######################################
    
    If IsNull(DLookup("MitID", "tblVerknuepftMit_FzgGrp", "MitID=" & mitID & " AND FzgGrpID= " & grpID)) Then
    
        sSQLMit = "INSERT INTO tblVerknuepftMit_FzgGrp (MitID,FzgGrpID,Datum) VALUES (" _
        & mitID & "," & grpID & ", '#" & Date & "#');"
        
        Debug.Print sSQLMit
        
        DoCmd.RunSQL sSQLMit
        
    End If

'6######################################## tblKostenstelle #################################

    If kstID = 0 Then
    
        sSQLKst = "INSERT INTO tblKostenstelle (AbtID, Kostenstelle) " & _
        "VALUES (" & abtID & ", '" & StammdatenXLSXdata(i).kstelle & "');"
        DoCmd.RunSQL sSQLKst
        
        kstID = Nz(DLookup("KstID", "tblKostenstelle", "Kostenstelle= '" & StammdatenXLSXdata(i).kstelle & "'"), 0)
        
        Debug.Print sSQLKst
    End If


'7######################################## tblStammnummer ##################################

    If StammNrID = 0 Then
    
        sSQLStNr = "INSERT INTO tblStammnummer (StammNr, REID, KstID, MitID) " & _
        "VALUES (" & StammdatenXLSXdata(i).stammNr & ", " & reID & ", " & kstID & "," & mitID & ");"
        
        Debug.Print sSQLStNr
        
        DoCmd.RunSQL sSQLStNr
        
    End If
    
DoCmd.SetWarnings True

Next i

'###############################Feststellungen und Vorschläge ################################################

    'eine Kostenstelle kann mehrere Abteilungen haben und eine Abteilung kann mehrere Kostenstellen haben.
    'also m:n Beziehung

End Sub

Public Sub CloseXLSXApp(bShowInfo As Boolean)
''' Excel-Instanz beenden
' Verweis auf Excel-Bibliothek muss gesetzt sein
Dim xlsApp As Excel.Application
Dim MsgAntw As Integer

' Excel-Instanz suchen
  On Error Resume Next
    Set xlsApp = GetObject(, "Excel.Aplication")
    If xlsApp Is Nothing Then
        ' keine Excel-Instanz vorhanden
        ' Meldung
        If bShowInfo Then
            MsgAntw = MsgBox("Es wurde keine Excel-Instanz gefunden.", vbInformation, "Excel-Instanz beenden")
        End If
        ' Ende
        Exit Sub
    End If

On Error GoTo 0
    ' Excel schließen und resetten
    xlsApp.Quit
    Set xlsApp = Nothing
    ' Meldung
    If bShowInfo Then
        MsgAntw = MsgBox("Die Excel-Instanz wurde beendet.", vbInformation, "Excel-Instanz beenden")
    End If
End Sub





