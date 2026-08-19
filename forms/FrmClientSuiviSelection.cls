Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Function ImpVisu(ByVal IntVisu As Integer)
   Dim QdfRequete As DAO.QueryDef
   Dim StrSQL As String

   On Error GoTo Err_ImpVisu

   For Each QdfRequete In CurrentDb.QueryDefs

      Select Case QdfRequete.Name
         Case "SelClientSuivi"

            CurrentDb.QueryDefs.Delete QdfRequete.Name

         Case Else

      End Select

   Next

   StrSQL = "SELECT CLIENT.Cli_Code, CLIENT.Cli_Nom, CLIENT.Cli_Date_contrat, CLIENT_TYPE.CT_Nom, CLIENT_GROUPE.CG_Nom As SousGroupe, CLIENT_GROUPE_1.CG_Nom As Groupe, CLIENT.Cli_Mont_Hon_m2, CLIENT.Cli_Pourcentage, CLIENT.Cli_Date_der_cré, " & _
                  "DLast(" & Chr(34) & "Reg_Date" & Chr(34) & "," & Chr(34) & "REGLEMENT" & Chr(34) & "," & Chr(34) & "Reg_Cli_Code='" & Chr(34) & " & [Cli_Code] & " & Chr(34) & "' AND Reg_Date>=" & Chr(34) & " & ConvDate([Cli_Date_contrat])) AS Cli_Dern_Reg_Date, " & _
                  "CCur(Nz(DLast(" & Chr(34) & "Reg_Montant_Encaisser_m2" & Chr(34) & "," & Chr(34) & "REGLEMENT" & Chr(34) & "," & Chr(34) & "Reg_Cli_Code='" & Chr(34) & " & [Cli_Code] & " & Chr(34) & "' AND Reg_Date>=" & Chr(34) & " & ConvDate([Cli_Date_contrat])),0)) AS Cli_Dern_Reg_Montant, " & _
                  "CCur(Nz(Sum(IIf([CG_Code_Action]=7 And [Déb_Date_Cré]>=[Cli_Date_contrat],Round([Déb_Mont_Prin_m2],2),0)),0)) AS Cli_Reg_Direct, " & _
                  "CCur(Nz(Sum(IIf([CG_Code_Action]=10,Round([Déb_Mont_Prin_m2],2),0)),0)) AS Cli_Representation, " & _
                  "CCur(Nz(Sum(IIf([CG_Code_Action]<>7 And [Déb_Date_Cré]>=[Cli_Date_contrat] And [Enc_Situation] Is Not Null,[Enc_Prin_m2],0)),0)) AS Cli_Reverser, " & _
                  "CCur(Nz(Sum(IIf([CG_Code_Action]<>7 And [Déb_Date_Cré]>=[Cli_Date_contrat] And [Enc_Situation] Is Null,[Enc_Prin_m2],0)),0)) AS Cli_ResteAVerser, " & _
                  "CCur(Nz(Sum(IIf([CG_Code_Action]<>7 And [Enc_Situation] Is Null,[Enc_Prin_m2],0)),0)) AS Cli_ResteAVerserGlobal " & _
                  "FROM (CLIENT_GROUPE RIGHT JOIN (CODE_GESTION RIGHT JOIN (((CLIENT_TYPE RIGHT JOIN CLIENT ON CLIENT_TYPE.CT_Id = CLIENT.Cli_Type) LEFT JOIN (DEBITEUR LEFT JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur) ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client) INNER JOIN TMP_REPRESENTANT_CLIENT ON CLIENT.Cli_Code = TMP_REPRESENTANT_CLIENT.TRCCli_Code) ON CODE_GESTION.CG_Code = DEBITEUR.Déb_Situation) ON CLIENT_GROUPE.CG_Id = CLIENT.Cli_Groupe) LEFT JOIN CLIENT_GROUPE AS CLIENT_GROUPE_1 ON CLIENT_GROUPE.CG_Parent_Id = CLIENT_GROUPE_1.CG_Id " & _
                  "WHERE (((TMP_REPRESENTANT_CLIENT.TRCRep_Code) = '" & Me.CmbTRCRep_Code & "') And ((TMP_REPRESENTANT_CLIENT.TRC_Flag) = True)) " & _
                  "GROUP BY CLIENT.Cli_Code, CLIENT.Cli_Nom, CLIENT.Cli_Date_contrat, CLIENT_TYPE.CT_Nom, CLIENT_GROUPE.CG_Nom, CLIENT_GROUPE_1.CG_Nom, CLIENT.Cli_Mont_Hon_m2, CLIENT.Cli_Pourcentage, CLIENT.Cli_Date_der_cré, " & _
                  "DLast(" & Chr(34) & "Reg_Date" & Chr(34) & "," & Chr(34) & "REGLEMENT" & Chr(34) & "," & Chr(34) & "Reg_Cli_Code='" & Chr(34) & " & [Cli_Code] & " & Chr(34) & "' AND Reg_Date>=" & Chr(34) & " & ConvDate([Cli_Date_contrat])), " & _
                  "Nz(DLast(" & Chr(34) & "Reg_Montant_Encaisser_m2" & Chr(34) & "," & Chr(34) & "REGLEMENT" & Chr(34) & "," & Chr(34) & "Reg_Cli_Code='" & Chr(34) & " & [Cli_Code] & " & Chr(34) & "' AND Reg_Date>=" & Chr(34) & " & ConvDate([Cli_Date_contrat])),0);"

   Set QdfRequete = CurrentDb.CreateQueryDef("SelClientSuivi", StrSQL)

   DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "SelClientSuivi", Format(Year(Date), "00") & Format(Month(Date), "00") & Format(Day(Date), "00") & "SelClientSuivi.XLS", True

   MsgBox "EXPORATION TERMINEE", vbCritical

Exit_ImpVisu:

   Exit Function

Err_ImpVisu:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ImpVisu"

   Resume Exit_ImpVisu
End Function

Public Function ToutRien(ByVal BolSelection As Boolean)
   Dim StrSQL As String

   On Error GoTo Err_ToutRien

   StrSQL = "UPDATE TMP_REPRESENTANT_CLIENT SET TRC_Flag = " & IIf(BolSelection, "True", "False") & _
                 " WHERE TRCRep_Code='" & CmbTRCRep_Code & "';"

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

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub CmbTRCRep_Code_AfterUpdate()

   On Error GoTo Err_CmbTRCRep_Code_AfterUpdate

   Me.Requery

Exit_CmbTRCRep_Code_AfterUpdate:

   Exit Sub

Err_CmbTRCRep_Code_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " CmbTRCRep_Code_AfterUpdate"

   Resume Exit_CmbTRCRep_Code_AfterUpdate
End Sub

Private Sub ChkTRC_Flag_AfterUpdate()

   On Error GoTo Err_ChkTRC_Flag_AfterUpdate

   DoCmd.RunCommand acCmdSaveRecord

Exit_ChkTRC_Flag_AfterUpdate:

   Exit Sub

Err_ChkTRC_Flag_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ChkTRC_Flag_AfterUpdate"

   Resume Exit_ChkTRC_Flag_AfterUpdate
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

Private Sub CmdSupprimer_Click()

   On Error GoTo Err_CmdSupprimer_Click

   CurrentDb.Execute "DELETE TMP_REPRESENTANT_CLIENT.* FROM TMP_REPRESENTANT_CLIENT WHERE TRCRep_Code='" & CmbTRCRep_Code & "' AND TRC_Flag=True;"

   Me.Requery

Exit_CmdSupprimer_Click:

   Exit Sub

Err_CmdSupprimer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdSupprimer_Click"

   Resume Exit_CmdSupprimer_Click
End Sub

Private Sub CmdClientListe_Click()
   Dim WksEspace As DAO.Workspace
   Dim DBBase As DAO.Database
   Dim StrSQL As String

   On Error GoTo Err_CmdClientListe_Click

   DoCmd.OpenForm "FrmClientSuiviSelectionClientListe", acNormal, , , , acDialog

   Select Case BolListeClientSelection
      Case True

         Set WksEspace = DBEngine.Workspaces(0)    ' Cherche la zone de travail par défaut.

         Set DBBase = WksEspace.Databases(0) ' Cherche la base de données active.

         WksEspace.BeginTrans

         Screen.MousePointer = 11

         StrSQL = "INSERT INTO TMP_REPRESENTANT_CLIENT ( TRCRep_Code, TRCCli_Code ,TRC_Flag) SELECT '" & CmbTRCRep_Code & "', TCli_Code, True FROM TMP_CLIENT " & _
                       "WHERE TCli_Code Not In (SELECT TRCCli_Code FROM TMP_REPRESENTANT_CLIENT WHERE TRCRep_Code='" & CmbTRCRep_Code & "');"

         DBBase.Execute StrSQL, dbFailOnError

         WksEspace.CommitTrans

         Me.Requery

      Case False

   End Select

Exit_CmdClientListe_Click:

   Screen.MousePointer = 0

   Exit Sub

Err_CmdClientListe_Click:

   WksEspace.Rollback

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdClientListe_Click"

   Resume Exit_CmdClientListe_Click
End Sub
