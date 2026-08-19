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

Private Sub CmdRelancer_Click()

   On Error GoTo Err_CmdRelancer_Click

   Select Case GenererRelance(TxtDateDebut, TxtDateFin)
      Case True

         DoCmd.OpenForm "FrmDebiteurRelanceSelection", acNormal, , , , acWindowNormal

      Case False

   End Select

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdRelancer_Click:

   Exit Sub

Err_CmdRelancer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdRelancer_Click"

   Resume Exit_CmdRelancer_Click
End Sub
