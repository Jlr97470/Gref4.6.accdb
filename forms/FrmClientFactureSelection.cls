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

           TxtEnc_Prin_m.ControlSource = "=DLookUp(" & Chr(34) & "Enc_Prin_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Id=" & Chr(34) & " & [TFac_Enc_Id])"

       Case cEURO

           TxtEnc_Prin_m.ControlSource = "=DLookUp(" & Chr(34) & "Enc_Prin_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Id=" & Chr(34) & " & [TFac_Enc_Id])"

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

   StrSQL = "UPDATE TMP_FACTURE SET TMP_FACTURE.TFac_Flag = " & IIf(BolSelection, "True", "False") & _
                 " WHERE (((TMP_FACTURE.TFac_Enc_ID) In (SELECT TMP_FACTURE.TFac_Enc_ID FROM (CLIENT INNER JOIN (DEBITEUR INNER JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur) ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client) INNER JOIN TMP_FACTURE ON ENCAISSEMENT.Enc_ID = TMP_FACTURE.TFac_Enc_ID " & _
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

   Select Case DCount("TFac_Enc_ID", "TMP_FACTURE")
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

   StrSQL = "DELETE TMP_FACTURE FROM TMP_FACTURE;"

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

Private Sub ChkTFac_Flag_AfterUpdate()

   On Error GoTo Err_ChkTFac_Flag_AfterUpdate

   DoCmd.RunCommand acCmdSaveRecord

Exit_ChkTFac_Flag_AfterUpdate:

   Exit Sub

Err_ChkTFac_Flag_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, stSource & " : ChkTFac_Flag_AfterUpdate"

   Resume Exit_ChkTFac_Flag_AfterUpdate
End Sub

Private Sub CmdClientListe_Click()
   Dim WksEspace As DAO.Workspace
   Dim DBBase As DAO.Database
   Dim StrSQL As String

   On Error GoTo Err_CmdClientListe_Click

   Select Case DCount("TFac_Enc_ID", "TMP_FACTURE") > 0 And BolListeClientSelection = False
      Case True

         MsgBox "Fonction en cours d'utilisation !!!", vbCritical

         GoTo Exit_CmdClientListe_Click

      Case False

   End Select

   DoCmd.OpenForm "FrmClientFactureSelectionClientListe", acNormal, , , , acDialog

   Select Case BolListeClientSelection
      Case True

         Set WksEspace = DBEngine.Workspaces(0)    ' Cherche la zone de travail par défaut.

         Set DBBase = WksEspace.Databases(0) ' Cherche la base de données active.

         WksEspace.BeginTrans

         Screen.MousePointer = 11
         ' *** Efface le contenu de la table REGLEMENT
         StrSQL = "DELETE * FROM TMP_FACTURE;"

         DBBase.Execute StrSQL, dbFailOnError
         ' *** Recherche tous les encaissements non encore réglé au client et dont le principal est différent de 0
         StrSQL = "INSERT INTO TMP_FACTURE ( TFac_Enc_Id ) SELECT ENCAISSEMENT.Enc_ID FROM (CLIENT INNER JOIN (DEBITEUR INNER JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur) ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client) INNER JOIN TMP_CLIENT ON CLIENT.Cli_Code = TMP_CLIENT.TCli_Code " & _
                       "WHERE ENCAISSEMENT.Enc_Situation='R' AND ENCAISSEMENT.Enc_Prin_m1<>0;"

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

Private Sub CmdFactureDefinitif_Click()

   On Error GoTo Err_CmdFactureDefinitif_Click

   DoCmd.OpenForm "FrmClientFactureSelectionEdition", acNormal, , , , acDialog

Exit_CmdFactureDefinitif_Click:

   Exit Sub

Err_CmdFactureDefinitif_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFactureDefinitif_Click"

   Resume Exit_CmdFactureDefinitif_Click
End Sub
