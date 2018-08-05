
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

###############################################################################################################################
##############################################################################################################################

Option Compare Database
Option Explicit

Private Sub cboUsername_AfterUpdate()
  Me.txtPassword.SetFocus
End Sub


Sub versteckePasswort()
  DoCmd.OpenForm "LoginForm"
 
  With Forms("LoginForm")
    .txtPassword.InputMask = "Password"
  End With
End Sub

Sub fuelleKombinationsfeld()
  Dim strSql As String
DoCmd.OpenForm "LoginForm"
strSql = "SELECT tblBenutzer.bzrLogin, tblBenutzer.bzrPass FROM tblBenutzer;"
With Forms("LoginForm").cboUsername
    .RowSource = strSql
    .ColumnCount = 1
  End With
End Sub

Private Sub cmd_close_Click()
 DoCmd.Close 'Schließt das Login-Formular
 DoCmd.Quit 'Schließt die komplette Access-Umgebung
End Sub
End Sub

Private Sub cmd_login_Click()

Dim logID As Long
Dim strCboPass As String
Dim strPass As String
Dim username As String

On Error GoTo Error_Handler

If Len(Trim(Me.txtPassword)) > 0 Then
  strCboPass = Me.cboUsername.Column(1)
  strPass = Me.txtPassword.Value
  username = Me.cboUsername

  If strCboPass = strPass Then
   DoCmd.Close acForm, Me.Name
  Else
   'Me.lblStatus.Visible = True
   With Me.txtPassword
    .Value = vbNullString
    .SetFocus
   End With
  End If
 ElseIf Len(Trim(Me.txtPassword)) = 0 Then
  MsgBox "Sie haben Ihr Passwort nicht eingegeben", vbInformation, "Passwort eingeben, bitte!"
 End If

Exit_Procedure:
 DoCmd.SetWarnings True
 Exit Sub

Error_Handler:
If IsNull(Me.cboUsername) Then
  MsgBox ("Sie haben Ihren Benutzernamen nicht eingegeben")
 Else
  MsgBox (Me.cboUsername.Value & " ist kein berechtigter Benutzer")
 End If
 Me.cboUsername.Value = vbNullString 'Null
 Me.cboUsername.SetFocus
 Me.txtPassword.Value = vbNullString
 Resume Exit_Procedure
End Sub


Private Sub Form_Load()
Call versteckePasswort
Call fuelleKombinationsfeld

End Sub

#############################################################################################################################
#############################################################################################################################
Option Compare Database
Option Explicit

Sub tabelleAnlegen()
    Dim qdf As DAO.QueryDef
    Dim db As DAO.Database
    Dim strSql As String
 
    'Sql Anweisung
    strSql = "CREATE TABLE tblBenutzer(" & _
             "bzrIdPk Autoincrement " & _
             "Constraint PrimaryKey PRIMARY KEY, " & _
             "bzrName Text(55), " & _
             "bzrVorname Text(50), " & _
             "bzrLogin Text(50), " & _
             "bzrPass Text(50))"
    Debug.Print strSql
    'Fehlerbehandlung
    On Error GoTo Fehler_Behandlung
 
    'Tabelle anlegen
    Set db = CurrentDb()
    db.Execute strSql
 
    MsgBox "tblBenutzer wurde angelegt!"
ExitSub:
    Set qdf = Nothing
    Set db = Nothing
    Exit Sub
 
Fehler_Behandlung:
    MsgBox "tblBenutzer konnte nicht angelegt werden!"
    Resume ExitSub
End Sub

#############################################################################################################################
#############################################################################################################################

Option Compare Database
Option Explicit

Sub FormularErstellen()
 Dim frm As Form
 Dim ctlLabel_Text As Control
 Dim ctlLabel_Kombi As Control
 Dim ctlText As Control
 Dim ctlKombi As Control
 Dim ctlButon1 As Control
 Dim ctlButon2 As Control
 Dim ctlRahmen As Control
 Dim formName As String
 
 ' Konstanten zur Positionierung der Controls auf dem Formular
 Const ctlBreite As Integer = 1450
 Const ctlMargeWaagerecht As Integer = 1500
 Const ctlMargeSenkrecht As Integer = 400
 Const ctlHoehe As Integer = 350
 Const waagerecht As Integer = 3000
 Const senkrecht As Integer = 2000
 
'Fehlerbehandlung
On Error Resume Next
 
 ' Ein neues Formular erstellen
 Set frm = CreateForm
 formName = frm.Name
 
 'Rahmen erstellen
 Set ctlRahmen = CreateControl(frm.Name, acRectangle, , , , 100, 150, 8500, 5000)
 ' ungebundene TextBox im DetailBereich erstellen
 
'Kombinationsfeld für den Benutzernamen
 Set ctlKombi = CreateControl(frm.Name, acComboBox, acDetail, , , _
 waagerecht + ctlMargeWaagerecht, senkrecht, ctlBreite + 100, ctlHoehe)
 ctlKombi.Name = "cboUsername"
 
 'Label wird Am Kombinationsfeld gebunden
 Set ctlLabel_Kombi = CreateControl(frm.Name, acLabel, acDetail, _
 ctlKombi.Name, "Benutzername", waagerecht, senkrecht, ctlBreite, ctlHoehe)
 
'Textfeld fürs Passwort
 Set ctlText = CreateControl(frm.Name, acTextBox, acDetail, "", "", _
 waagerecht + ctlMargeWaagerecht, senkrecht + ctlMargeSenkrecht, ctlBreite + 100, ctlHoehe)
 ctlText.Name = "txtPassword"
 
 'Label wird am TextFeld  gebunden
 Set ctlLabel_Text = CreateControl(frm.Name, acLabel, acDetail, _
 ctlText.Name, , waagerecht, senkrecht + ctlMargeSenkrecht, ctlBreite, ctlHoehe)
 ctlLabel_Text.Caption = "Passwort"
 
'Buton zum Schliessen des Formulars
 Set ctlButon1 = CreateControl(frm.Name, acCommandButton, acDetail, , , _
 waagerecht, senkrecht + ctlMargeSenkrecht * 3, ctlBreite, ctlHoehe)
 ctlButon1.Name = "cmd_close"
 ctlButon1.Caption = "Abbrechen"
 
'Buton zum Einlogen
 Set ctlButon2 = CreateControl(frm.Name, acCommandButton, acDetail, , , _
 waagerecht + ctlMargeWaagerecht, senkrecht + ctlMargeSenkrecht * 3, ctlBreite, ctlHoehe)
 ctlButon2.Caption = "Einloggen"
 ctlButon2.Name = "cmd_login"
 
 ' rufe Funktion zur Formatierung des Formulars auf
 Call formularFormatieren(frm)
 DoCmd.Restore
 
 'Formular wird umbennant
 DoCmd.Close acForm, formName, acSaveYes
 DoCmd.SelectObject acForm, formName, True
 DoCmd.Rename "LoginForm", acForm, formName
 
 'Formular wird normal geöffnet
 DoCmd.OpenForm "LoginForm", acNormal
 Set frm = Nothing
End Sub

Sub formularFormatieren(frm As Form)
 On Error Resume Next
   With frm
     .Detail.Height = 5200
     .BorderStyle = acDialog
     .PopUp = True
     .Modal = True
     .ScrollBars = 0
     .RecordSelectors = False
     .NavigationButtons = False
     .Detail.BackColor = 14741230
     .Caption = "Loggen Sie sich bitte ein!"
 End With
End Sub





