Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private SQLBase As String
Private SQLWhereGlobal As String
Private SQLWhere As String
Private SQLOrder As String

Private Function ToutRien(ByVal BolSelection As Boolean)
   Dim StrSQL As String

   On Error GoTo Err_ToutRien

   Screen.MousePointer = 11

   StrSQL = "UPDATE FACTURE SET Fac_Editer = " & IIf(BolSelection, "True", "False") & " WHERE Fac_Id In (SELECT FACTURE.Fac_Id FROM CLIENT INNER JOIN FACTURE ON CLIENT.Cli_Code=FACTURE.Fac_CodeCli " & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & ";);"

   CurrentDb.Execute StrSQL, dbFailOnError

   Me.Requery

   Screen.MousePointer = 0

Exit_ToutRien:

   Exit Function

Err_ToutRien:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ToutRien"

   Resume Exit_ToutRien
End Function

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   SQLBase = "SELECT FACTURE.Fac_Num, FACTURE.Fac_CodeCli, CLIENT.Cli_Nom, FACTURE.Fac_Date, FACTURE.Fac_Editer, FACTURE.Fac_Deja_Editer FROM CLIENT INNER JOIN FACTURE ON CLIENT.Cli_Code=FACTURE.Fac_CodeCli"

   SQLWhereGlobal = ""

   SQLWhere = ""

   SQLOrder = "Fac_Num"

   Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

   CurrentDb.Execute "UPDATE FACTURE SET Fac_Editer=False;", dbFailOnError

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub CmdImprimer_Click()

   On Error GoTo Err_CmdImprimer_Click

   DoCmd.OpenForm "FrmClientFactureDefinitifSelectionFiltreEdition", acNormal, , , , acDialog, "Déb_Nom;FacD_Déb_Code;Déb_Réf1_Client"

Exit_CmdImprimer_Click:

   Exit Sub

Err_CmdImprimer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdImprimer_Click"

   Resume Exit_CmdImprimer_Click
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

Private Sub ChkFac_Editer_AfterUpdate()

   On Error GoTo Err_ChkFac_Editer_AfterUpdate

   DoCmd.RunCommand acCmdSaveRecord

Exit_ChkFac_Editer_AfterUpdate:

   Exit Sub

Err_ChkFac_Editer_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ChkFac_Editer_AfterUpdate"

   Resume Exit_ChkFac_Editer_AfterUpdate
End Sub

Private Sub CmbGroupe_AfterUpdate()

   On Error GoTo Err_CmbGroupe_AfterUpdate

   CmbCli_Groupe.Requery

   CmbCli_Groupe = Null

Exit_CmbGroupe_AfterUpdate:

   Exit Sub

Err_CmbGroupe_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbGroupe_AfterUpdate:"

   Resume Exit_CmbGroupe_AfterUpdate
End Sub

Private Sub CmdOK_Click()
   Dim BolWhere As Boolean
   Dim StrFac_CodeCli As String
   Dim StrCli_Nom As String
   Dim StrFac_DateDebut As String
   Dim StrFac_DateFin As String
   Dim StrFac_Num As String
   Dim LngCG_Parent_Id As Long
   Dim LngCli_Groupe As Long

   On Error GoTo Err_CmdOK_Click

   StrFac_Num = Nz(TxtRech_Fac_Num)

   StrFac_CodeCli = Nz(TxtRech_Fac_CodeCli)

   StrCli_Nom = Nz(TxtRech_Cli_Nom)

   StrFac_DateDebut = Nz(TxtRech_Fac_DateDebut)

   StrFac_DateFin = Nz(TxtRech_Fac_DateFin)

   LngCG_Parent_Id = Nz(CmbGroupe)

   LngCli_Groupe = Nz(CmbCli_Groupe)

   BolWhere = False

   SQLWhere = ""

   Select Case StrFac_Num
      Case ""

      Case Else

            Select Case InStr(StrFac_Num, "*")
               Case Is > 0

               Case Else

                 StrFac_Num = StrFac_Num & "*"

            End Select

            SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Fac_Num] LIKE """ & StrFac_Num & """"

            BolWhere = True

   End Select

   Select Case StrFac_CodeCli
      Case ""

      Case Else

            Select Case InStr(StrFac_CodeCli, "*")
               Case Is > 0

               Case Else

                 StrFac_CodeCli = StrFac_CodeCli & "*"

            End Select

            SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Fac_CodeCli] LIKE """ & StrFac_CodeCli & """"

            BolWhere = True

   End Select

   Select Case StrCli_Nom
      Case ""

      Case Else

         Select Case InStr(StrCli_Nom, "*")
            Case Is > 0

            Case Else

               StrCli_Nom = StrCli_Nom & "*"

         End Select

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Cli_Nom] LIKE """ & StrCli_Nom & """"

         BolWhere = True

   End Select

   Select Case StrFac_DateDebut
      Case ""

      Case Else

         StrFac_DateDebut = ConvDate(StrFac_DateDebut)

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Fac_Date]>=" & StrFac_DateDebut

         BolWhere = True

   End Select

   Select Case StrFac_DateFin
      Case ""

      Case Else

         StrFac_DateFin = ConvDate(StrFac_DateFin)

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Fac_Date]<=" & StrFac_DateFin

         BolWhere = True

   End Select

   Select Case LngCG_Parent_Id
      Case 0

      Case Else

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Cli_Groupe] In (SELECT CG_Id FROM CLIENT_GROUPE WHERE CG_Parent_Id=" & LngCG_Parent_Id & ";)"

         BolWhere = True

   End Select

   Select Case LngCli_Groupe
      Case 0

      Case Else

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Cli_Groupe]=" & LngCli_Groupe

         BolWhere = True

   End Select

   Select Case CadClass
      Case 1

         SQLOrder = "Fac_Num"

      Case 2

         SQLOrder = "Fac_CodeCli"

      Case 3

         SQLOrder = "Cli_Nom"

      Case 4

         SQLOrder = "Fac_Date"

   End Select

   CurrentDb.Execute "UPDATE FACTURE SET Fac_Editer=False;", dbFailOnError

   Select Case SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"
      Case Me.RecordSource

      Case Else

         Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

   End Select

Exit_CmdOK_Click:

   Exit Sub

Err_CmdOK_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

   Resume Exit_CmdOK_Click
End Sub
