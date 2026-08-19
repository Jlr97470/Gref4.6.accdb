Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
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

   Select Case DCount("TIrr_Déb_Code", "TMP_IRRECOUVRABLE") > 0
      Case True

         MsgBox "Fonction en cours d'utilisation !!!", vbCritical

         GoTo Exit_CmdOK_Click

      Case False

   End Select

   Select Case GenererIrrecouvrable(TxtCli_Code, TxtDateFin)
      Case True

         DoCmd.Close acForm, Me.Name, acSavePrompt

         DoCmd.OpenForm "FrmDebiteurIrrecouvrableSelection", acNormal, , , , acWindowNormal

      Case False

         DoCmd.Close acForm, Me.Name, acSavePrompt

   End Select

Exit_CmdOK_Click:

   Exit Sub

Err_CmdOK_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

   Resume Exit_CmdOK_Click
End Sub
