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

Private Function ToutRien(ByVal BolSelection As Boolean)
   Dim StrSQL As String

   On Error GoTo Err_ToutRien

   Screen.MousePointer = 11

   StrSQL = "DELETE * FROM TMP_CLIENT;"

   CurrentDb.Execute StrSQL, dbFailOnError

   Select Case BolSelection
      Case True

         StrSQL = "INSERT INTO TMP_CLIENT ( TCli_Code ) SELECT Cli_Code FROM CLIENT" & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & ";"

      Case False

   End Select

   CurrentDb.Execute StrSQL, dbFailOnError

   ChkCli_Flag.Requery

   Screen.MousePointer = 0

Exit_ToutRien:

   Exit Function

Err_ToutRien:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ToutRien"

   Resume Exit_ToutRien
End Function

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   SQLBase = "SELECT CLIENT.* FROM CLIENT"

   SQLWhere = ""

   SQLWhereGlobal = ""

   SQLOrder = "Cli_Code"

   Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

   ToutRien False

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   BolListeClientSelection = False

   ToutRien False

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

    Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

    Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdRechercher_Click()

    On Error GoTo Err_CmdRechercher_Click

    DoCmd.OpenForm "FrmFicheClientRechercher", acNormal, , , acFormEdit, acDialog, "FrmClientFactureSelectionClientListe"

    ToutRien False

    Select Case SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"
      Case Me.RecordSource

      Case Else

          Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

   End Select

Exit_CmdRechercher_Click:

    Exit Sub

Err_CmdRechercher_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdRechercher_Click"

    Resume Exit_CmdRechercher_Click
End Sub

Private Sub CmdOK_Click()

   On Error GoTo Err_CmdOK_Click

   BolListeClientSelection = True

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdOK_Click:

   Exit Sub

Err_CmdOK_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOK_Click"

   Resume Exit_CmdOK_Click
End Sub

Private Sub ChkCli_Flag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim StrSQL As String

   On Error GoTo Err_ChkCli_Flag_MouseDown

   Select Case ChkCli_Flag
      Case True

         StrSQL = "DELETE * FROM TMP_CLIENT WHERE TMP_CLIENT.TCli_Code='" & Cli_Code & "';"

      Case False

         StrSQL = "INSERT INTO TMP_CLIENT ( TCli_Code ) VALUES('" & Cli_Code & "');"

   End Select

   CurrentDb.Execute StrSQL, dbFailOnError

   ChkCli_Flag.Requery

Exit_ChkCli_Flag_MouseDown:

   Exit Sub

Err_ChkCli_Flag_MouseDown:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ChkCli_Flag_MouseDown"

   Resume Exit_ChkCli_Flag_MouseDown
End Sub
