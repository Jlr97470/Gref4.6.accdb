Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Public SQLBase As String
Public SQLWhere As String
Public SQLWhereGlobal As String
Public SQLOrder As String

Public Function ToutRien(ByVal BolSelection As Boolean)
   Dim StrSQL As String

   On Error GoTo Err_ToutRien

   Select Case CadMode
      Case 1

         Select Case BolSelection
            Case True

               StrSQL = "INSERT INTO TMP_DEBITEUR (TDéb_Code) SELECT DEBITEUR.Déb_Code FROM DEBITEUR LEFT JOIN TMP_DEBITEUR ON DEBITEUR.Déb_Code = TMP_DEBITEUR.TDéb_Code WHERE TDéb_Code Is Null " & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " AND ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & ";"

            Case False

               StrSQL = "DELETE TMP_DEBITEUR.* FROM TMP_DEBITEUR WHERE TDéb_Code In (SELECT DEBITEUR.Déb_Code FROM DEBITEUR LEFT JOIN TMP_DEBITEUR ON DEBITEUR.Déb_Code = TMP_DEBITEUR.TDéb_Code WHERE Not TDéb_Code Is Null " & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " AND ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & ";);"

         End Select

      Case 2

         Select Case BolSelection
            Case True

               StrSQL = "INSERT INTO TMP_DEBITEUR (TDéb_Code) SELECT ARCHIVE_DEBITEUR.Déb_Code FROM ARCHIVE_DEBITEUR LEFT JOIN TMP_DEBITEUR ON ARCHIVE_DEBITEUR.Déb_Code = TMP_DEBITEUR.TDéb_Code WHERE TDéb_Code Is Null " & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " AND ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & ";"

            Case False

               StrSQL = "DELETE TMP_DEBITEUR.* FROM TMP_DEBITEUR WHERE TDéb_Code In (SELECT ARCHIVE_DEBITEUR.Déb_Code FROM ARCHIVE_DEBITEUR LEFT JOIN TMP_DEBITEUR ON ARCHIVE_DEBITEUR.Déb_Code = TMP_DEBITEUR.TDéb_Code WHERE Not TDéb_Code Is Null " & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " AND ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & ";);"

         End Select

    End Select

   CurrentDb.Execute StrSQL, dbFailOnError

   Me.Requery

Exit_ToutRien:

   Exit Function

Err_ToutRien:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ToutRien"

   Resume Exit_ToutRien
End Function

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   SQLBase = "SELECT * FROM DEBITEUR"

   SQLWhere = ""

   SQLOrder = "Déb_Code"

   Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

   CurrentDb.Execute "DELETE TMP_DEBITEUR.* FROM TMP_DEBITEUR;", dbFailOnError

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub CadMode_AfterUpdate()

   On Error GoTo Err_CadMode_AfterUpdate

   Select Case CadMode
     Case 1

       Me.Caption = "Liste des Débiteurs à archiver"

       CmdArchiver.Enabled = True

       CmdDesarchiver.Enabled = False

       SQLBase = "SELECT * FROM DEBITEUR"

     Case 2

       Me.Caption = "Liste des Débiteurs à désarchiver"

       CmdArchiver.Enabled = False

       CmdDesarchiver.Enabled = True

       SQLBase = "SELECT * FROM ARCHIVE_DEBITEUR"

   End Select

   CurrentDb.Execute "DELETE TMP_DEBITEUR.* FROM TMP_DEBITEUR;", dbFailOnError

   Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

Exit_CadMode_AfterUpdate:

   Exit Sub

Err_CadMode_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CadMode_AfterUpdate"

   Resume Exit_CadMode_AfterUpdate
End Sub

Private Sub ChkArchiver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim StrSQL As String

   On Error GoTo Err_ChkArvhiver_MouseDown

   Select Case ChkArchiver
      Case True

         StrSQL = "DELETE TMP_DEBITEUR.* FROM TMP_DEBITEUR WHERE TDéb_Code='" & Déb_Code & "';"

      Case False

         StrSQL = "INSERT INTO TMP_DEBITEUR (TDéb_Code) VALUES('" & Déb_Code & "');"

   End Select

   CurrentDb.Execute StrSQL, dbFailOnError

   ChkArchiver.Requery

Exit_ChkArvhiver_MouseDown:

   Exit Sub

Err_ChkArvhiver_MouseDown:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ChkArvhiver_MouseDown"

   Resume Err_ChkArvhiver_MouseDown
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdRechercher_Click()

   On Error GoTo Err_CmdRechercher_Click

   DoCmd.OpenForm "FrmFicheDebiteurRechercher", acNormal, , , , acDialog, "FrmDebiteurArchivage"

   CurrentDb.Execute "DELETE TMP_DEBITEUR.* FROM TMP_DEBITEUR;", dbFailOnError

   Select Case SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"
     Case Me.RecordSource

     Case Else

        Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

   End Select

   RecToutRien.Visible = SQLWhere <> ""

   EtiRien.Visible = SQLWhere <> ""

   EtiTout.Visible = SQLWhere <> ""

Exit_CmdRechercher_Click:

   Exit Sub

Err_CmdRechercher_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdRechercher_Click"

   Resume Exit_CmdRechercher_Click
End Sub

Private Sub CmdArchiver_Click()
   Dim StrSQL As String
   Dim DBBase As DAO.Database
   Dim WksEspace As DAO.Workspace

   On Error GoTo Err_CmdArchiver_Click

   Screen.MousePointer = 11

   Set WksEspace = DBEngine.Workspaces(0)

   Set DBBase = WksEspace.Databases(0)

   WksEspace.BeginTrans

   StrSQL = "INSERT INTO ARCHIVE_DEBITEUR SELECT DEBITEUR.* FROM DEBITEUR INNER JOIN TMP_DEBITEUR ON DEBITEUR.Déb_Code = TMP_DEBITEUR.TDéb_Code;"

   DBBase.Execute StrSQL, dbFailOnError

   StrSQL = "INSERT INTO ARCHIVE_ADRESSE_DEBITEUR SELECT ADRESSE_DEBITEUR.* FROM ADRESSE_DEBITEUR INNER JOIN TMP_DEBITEUR ON ADRESSE_DEBITEUR.AD_Code_Débiteur = TMP_DEBITEUR.TDéb_Code;"

   DBBase.Execute StrSQL, dbFailOnError

   StrSQL = "INSERT INTO ARCHIVE_ENCAISSEMENT SELECT ENCAISSEMENT.* FROM ENCAISSEMENT INNER JOIN TMP_DEBITEUR ON ENCAISSEMENT.Enc_Code_Débiteur = TMP_DEBITEUR.TDéb_Code;"

   DBBase.Execute StrSQL, dbFailOnError

   StrSQL = "DELETE ADRESSE_DEBITEUR.* FROM ADRESSE_DEBITEUR WHERE ADRESSE_DEBITEUR.AD_Code_Débiteur In (SELECT TDéb_Code FROM TMP_DEBITEUR;);"

   DBBase.Execute StrSQL, dbFailOnError

   StrSQL = "DELETE ENCAISSEMENT.* FROM ENCAISSEMENT WHERE ENCAISSEMENT.Enc_Code_Débiteur In (SELECT TDéb_Code FROM TMP_DEBITEUR;);"

   DBBase.Execute StrSQL, dbFailOnError

   StrSQL = "DELETE DEBITEUR.* FROM DEBITEUR WHERE DEBITEUR.Déb_Code In (SELECT TDéb_Code FROM TMP_DEBITEUR;);"

   DBBase.Execute StrSQL, dbFailOnError

   StrSQL = "DELETE TMP_DEBITEUR.* FROM TMP_DEBITEUR;"

   DBBase.Execute StrSQL, dbFailOnError

   WksEspace.CommitTrans

   Me.Requery

Exit_CmdArchiver_Click:

   Screen.MousePointer = 0

   Exit Sub

Err_CmdArchiver_Click:

   WksEspace.Rollback

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdArchiver_Click"

   Resume Exit_CmdArchiver_Click
End Sub

Private Sub CmdDesarchiver_Click()
   Dim StrSQL As String
   Dim DBBase As DAO.Database
   Dim WksEspace As DAO.Workspace

   On Error GoTo Err_CmdArchiver_Click

   Screen.MousePointer = 11

   Set WksEspace = DBEngine.Workspaces(0)

   Set DBBase = WksEspace.Databases(0)

   WksEspace.BeginTrans

   StrSQL = "INSERT INTO DEBITEUR SELECT ARCHIVE_DEBITEUR.* FROM ARCHIVE_DEBITEUR INNER JOIN TMP_DEBITEUR ON ARCHIVE_DEBITEUR.Déb_Code = TMP_DEBITEUR.TDéb_Code;"

   DBBase.Execute StrSQL, dbFailOnError

   StrSQL = "INSERT INTO ADRESSE_DEBITEUR SELECT ARCHIVE_ADRESSE_DEBITEUR.* FROM ARCHIVE_ADRESSE_DEBITEUR INNER JOIN TMP_DEBITEUR ON ARCHIVE_ADRESSE_DEBITEUR.AD_Code_Débiteur = TMP_DEBITEUR.TDéb_Code;"

   DBBase.Execute StrSQL, dbFailOnError

   StrSQL = "INSERT INTO ENCAISSEMENT SELECT ARCHIVE_ENCAISSEMENT.* FROM ARCHIVE_ENCAISSEMENT INNER JOIN TMP_DEBITEUR ON ARCHIVE_ENCAISSEMENT.Enc_Code_Débiteur = TMP_DEBITEUR.TDéb_Code;"

   DBBase.Execute StrSQL, dbFailOnError

   StrSQL = "DELETE ARCHIVE_ADRESSE_DEBITEUR.* FROM ARCHIVE_ADRESSE_DEBITEUR WHERE ARCHIVE_ADRESSE_DEBITEUR.AD_Code_Débiteur In (SELECT TDéb_Code FROM TMP_DEBITEUR;);"

   DBBase.Execute StrSQL, dbFailOnError

   StrSQL = "DELETE ARCHIVE_ENCAISSEMENT.* FROM ARCHIVE_ENCAISSEMENT WHERE ARCHIVE_ENCAISSEMENT.Enc_Code_Débiteur In (SELECT TDéb_Code FROM TMP_DEBITEUR;);"

   DBBase.Execute StrSQL, dbFailOnError

   StrSQL = "DELETE ARCHIVE_DEBITEUR.* FROM ARCHIVE_DEBITEUR WHERE ARCHIVE_DEBITEUR.Déb_Code In (SELECT TDéb_Code FROM TMP_DEBITEUR;);"

   DBBase.Execute StrSQL, dbFailOnError

   StrSQL = "DELETE TMP_DEBITEUR.* FROM TMP_DEBITEUR;"

   DBBase.Execute StrSQL, dbFailOnError

   WksEspace.CommitTrans

   Me.Requery

Exit_CmdArchiver_Click:

   Screen.MousePointer = 0

   Exit Sub

Err_CmdArchiver_Click:

   WksEspace.Rollback

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdArchiver_Click"

   Resume Exit_CmdArchiver_Click
End Sub
