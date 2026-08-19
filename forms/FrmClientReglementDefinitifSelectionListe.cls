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

   StrSQL = "UPDATE REGLEMENT SET Reg_Editer = " & IIf(BolSelection, "True", "False") & " WHERE Reg_Id In (SELECT REGLEMENT.Reg_Id FROM CLIENT INNER JOIN REGLEMENT ON CLIENT.Cli_Code=REGLEMENT.Reg_Cli_Code" & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & ";);"

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

   SQLBase = "SELECT REGLEMENT.Reg_Cli_Code, CLIENT.Cli_Nom, REGLEMENT.Reg_Date, REGLEMENT.Reg_Editer, REGLEMENT.Reg_Deja_Editer FROM CLIENT INNER JOIN REGLEMENT ON CLIENT.Cli_Code=REGLEMENT.Reg_Cli_Code"

   SQLWhereGlobal = ""

   SQLWhere = ""

   SQLOrder = ""

   Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

   CurrentDb.Execute "UPDATE REGLEMENT SET Reg_Editer=False;", dbFailOnError

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub CmdImprimer_Click()

   On Error GoTo Err_CmdImprimer_Click

   DoCmd.OpenForm "FrmClientReglementDefinitifSelectionFiltreEdition", acNormal, , , , acDialog, "Déb_Nom;RegD_Déb_Code;Déb_Réf1_Client"

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

Private Sub ChkReg_Editer_AfterUpdate()

   On Error GoTo Err_ChkReg_Editer_AfterUpdate

   DoCmd.RunCommand acCmdSaveRecord

Exit_ChkReg_Editer_AfterUpdate:

   Exit Sub

Err_ChkReg_Editer_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ChkReg_Editer_AfterUpdate"

   Resume Exit_ChkReg_Editer_AfterUpdate
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
   Dim StrReg_Cli_Code As String
   Dim StrCli_Nom As String
   Dim StrReg_DateDebut As String
   Dim StrReg_DateFin As String
   Dim LngCG_Parent_Id As Long
   Dim LngCli_Groupe As Long

   On Error GoTo Err_CmdOK_Click

   StrReg_Cli_Code = Nz(TxtRech_Reg_Cli_Code)

   StrCli_Nom = Nz(TxtRech_Cli_Nom)

   StrReg_DateDebut = Nz(TxtRech_Reg_DateDebut)

   StrReg_DateFin = Nz(TxtRech_Reg_DateFin)

   LngCG_Parent_Id = Nz(CmbGroupe)

   LngCli_Groupe = Nz(CmbCli_Groupe)

   BolWhere = False

   SQLWhere = ""

   Select Case StrReg_Cli_Code
      Case ""

      Case Else

            Select Case InStr(StrReg_Cli_Code, "*")
               Case Is > 0

               Case Else

                 StrReg_Cli_Code = StrReg_Cli_Code & "*"

            End Select

            SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Reg_Cli_Code] LIKE """ & StrReg_Cli_Code & """"

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

   Select Case StrReg_DateDebut
      Case ""

      Case Else

         StrReg_DateDebut = ConvDate(StrReg_DateDebut)

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Reg_Date]>=" & StrReg_DateDebut

         BolWhere = True

   End Select

   Select Case StrReg_DateFin
      Case ""

      Case Else

         StrReg_DateFin = ConvDate(StrReg_DateFin)

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Reg_Date]<=" & StrReg_DateFin

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

         SQLOrder = "Reg_Cli_Code"

      Case 2

         SQLOrder = "Cli_Nom"

      Case 3

         SQLOrder = "Reg_Date"

   End Select

   CurrentDb.Execute "UPDATE REGLEMENT SET Reg_Editer=False;", dbFailOnError

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
