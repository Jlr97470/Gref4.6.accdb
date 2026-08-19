Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdOK_Click()

   On Error GoTo Err_CmdOK_Click

   Select Case IsNull(TxtHonoraire) Or IsNull(TxtFactureNumero)
      Case True

      Case False

         CurrentDb.Execute "UPDATE FACTURE SET Fac_Editer=False;", dbFailOnError

         Select Case GenererFacture(TxtHonoraire, TxtFactureNumero)
            Case True

               Forms!FrmClientFactureSelection.Form.Requery

               DoCmd.OpenForm "FrmClientFactureDefinitifSelectionFiltreEdition", acNormal, , , , acDialog, "Déb_Nom;FacD_Déb_Code;Déb_Réf1_Client"

            Case False

         End Select

   End Select

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdOK_Click:

   Exit Sub

Err_CmdOK_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

   Resume Exit_CmdOK_Click
End Sub
