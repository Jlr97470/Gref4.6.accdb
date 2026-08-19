Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private iMonnaieVisu As Integer

Public Sub Inverseur(ByVal IntMonnaie As Integer)
    'Cette fonction permet de visualiser une fiche dans une monnaie particulière (argMonnaie)
   On Error GoTo Err_Inverseur

   Select Case IntMonnaie
       Case cFRANC

           TxtEnc_Prin_m.ControlSource = "=DLookUp(" & Chr(34) & "Enc_Prin_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Id=" & Chr(34) & " & [TReg_Enc_Id])"

       Case cEURO

           TxtEnc_Prin_m.ControlSource = "=DLookUp(" & Chr(34) & "Enc_Prin_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Id=" & Chr(34) & " & [TReg_Enc_Id])"

   End Select

   Me.Filter = "Cli_MonnaieEnCours = " & IntMonnaie

   Me.FilterOn = True

   CmbMonnaie = IntMonnaie

   iMonnaieVisu = IntMonnaie

Exit_Inverseur:

    Exit Sub

Err_Inverseur:

    MsgBox Err.Number & " : " & Err.description, vbCritical, stSource & " : Inverseur"

    Resume Exit_Inverseur
End Sub

Public Function ToutRien(ByVal BolSelection As Boolean)
   Dim StrSQL As String

   On Error GoTo Err_ToutRien

   Screen.MousePointer = 11

   StrSQL = "UPDATE TMP_REGLEMENT SET TMP_REGLEMENT.TReg_Flag = " & IIf(BolSelection, "True", "False") & _
                 " WHERE (((TMP_REGLEMENT.TReg_Enc_ID) In (SELECT TMP_REGLEMENT.TReg_Enc_ID FROM (CLIENT INNER JOIN (DEBITEUR INNER JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur) ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client) INNER JOIN TMP_REGLEMENT ON ENCAISSEMENT.Enc_ID = TMP_REGLEMENT.TReg_Enc_ID " & _
                 " WHERE (((CLIENT.Cli_MonnaieEnCours)=" & iMonnaieVisu & "));)));"

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

   Select Case DCount("TReg_Enc_ID", "TMP_REGLEMENT")
      Case 0

      Case Else

         MsgBox "Fonction en cours d'utilisation !!!", vbCritical

         Cancel = True

   End Select

   Inverseur cEURO

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub CmdFermer_Click()
   Dim StrSQL As String

   On Error GoTo Err_CmdFermer_Click

   StrSQL = "DELETE TMP_REGLEMENT.* FROM TMP_REGLEMENT;"

   CurrentDb.Execute StrSQL, dbFailOnError

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdImprimer_Click()

   On Error GoTo Err_CmdImprimer_Click

   DoCmd.OpenForm "FrmFiltreEdition", acNormal, , , , acDialog, "EtaBrouillardReglement;" & iMonnaieVisu & ";Déb_Nom;Déb_Code;Déb_Réf1_Client;"

Exit_CmdImprimer_Click:

   Exit Sub

Err_CmdImprimer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdImprimer_Click"

   Resume Exit_CmdImprimer_Click
End Sub

Private Sub CmbMonnaie_AfterUpdate()

   On Error GoTo Err_CmbMonnaie_AfterUpdate

   Inverseur CmbMonnaie

Exit_CmbMonnaie_AfterUpdate:

   Exit Sub

Err_CmbMonnaie_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbMonnaie_AfterUpdate"

   Resume Exit_CmbMonnaie_AfterUpdate
End Sub

Private Sub ChkTReg_Flag_AfterUpdate()

   On Error GoTo Err_ChkTReg_Flag_AfterUpdate

   DoCmd.RunCommand acCmdSaveRecord

Exit_ChkTReg_Flag_AfterUpdate:

   Exit Sub

Err_ChkTReg_Flag_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, stSource & " : ChkTReg_Flag_AfterUpdate"

   Resume Exit_ChkTReg_Flag_AfterUpdate
End Sub

Private Sub CmdClientListe_Click()
   Dim WksEspace As DAO.Workspace
   Dim DBBase As DAO.Database
   Dim StrSQL As String

   On Error GoTo Err_CmdClientListe_Click

   Select Case DCount("TReg_Enc_ID", "TMP_REGLEMENT") > 0 And BolListeClientSelection = False
      Case True

         MsgBox "Fonction en cours d'utilisation !!!", vbCritical

         GoTo Exit_CmdClientListe_Click

      Case False

   End Select

   DoCmd.OpenForm "FrmClientReglementSelectionClientListe", acNormal, , , , acDialog

   Select Case BolListeClientSelection
      Case True

         Set WksEspace = DBEngine.Workspaces(0)    ' Cherche la zone de travail par défaut.

         Set DBBase = WksEspace.Databases(0) ' Cherche la base de données active.

         WksEspace.BeginTrans

         Screen.MousePointer = 11
         ' *** Efface le contenu de la table REGLEMENT
         StrSQL = "DELETE * FROM TMP_REGLEMENT;"

         DBBase.Execute StrSQL, dbFailOnError
         ' *** Recherche tous les encaissements non encore réglé au client et dont le principal est différent de 0
         StrSQL = "INSERT INTO TMP_REGLEMENT ( TReg_Enc_Id ) SELECT ENCAISSEMENT.Enc_ID FROM (CLIENT INNER JOIN (DEBITEUR INNER JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur) ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client) INNER JOIN TMP_CLIENT ON CLIENT.Cli_Code = TMP_CLIENT.TCli_Code " & _
                       "WHERE ENCAISSEMENT.Enc_Situation Is Null AND ENCAISSEMENT.Enc_Prin_m1<>0;"

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

Private Sub CmdReglementDefinitif_Click()

   On Error GoTo Err_CmdReglementDefinitif_Click

   CurrentDb.Execute "UPDATE REGLEMENT SET Reg_Editer=False;", dbFailOnError

   Select Case GenererReglement
      Case True

         Me.Requery

         DoCmd.OpenForm "FrmClientReglementDefinitifSelectionFiltreEdition", acNormal, , , , acDialog, "Déb_Nom;RegD_Déb_Code;Déb_Réf1_Client"

      Case False

   End Select

Exit_CmdReglementDefinitif_Click:

   Exit Sub

Err_CmdReglementDefinitif_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdReglementDefinitif_Click"

   Resume Exit_CmdReglementDefinitif_Click
End Sub
