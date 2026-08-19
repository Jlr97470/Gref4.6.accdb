Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Sub CmdDetail_Click()

    On Error GoTo Err_CmdDetail_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdDetail_Click

   End Select

   On Error GoTo Err_CmdDetail_Click

    Select Case Not IsNull(AD_TypeAdresse_ID)
      Case True

         DoCmd.OpenForm "FrmFicheDebiteurDetailAdresse", acNormal, , , , acDialog, AD_Code_Débiteur & ";" & AD_TypeAdresse_ID

      Case False

   End Select

Exit_CmdDetail_Click:

    Exit Sub

Err_CmdDetail_Click:

    MsgBox Error

    Resume Exit_CmdDetail_Click
End Sub
