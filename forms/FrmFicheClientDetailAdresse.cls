Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Sub Form_Open(Cancel As Integer)
   Dim StrParam() As String

   On Error GoTo Err_Form_Open

   StrParam = Split(Nz(Me.OpenArgs), ";")

   Select Case UBound(StrParam)
      Case Is < 1

         Me.RecordSource = "SELECT ADRESSE_CLIENT.* FROM ADRESSE_CLIENT;"

      Case Else

         Me.RecordSource = "SELECT ADRESSE_CLIENT.* FROM ADRESSE_CLIENT WHERE [AC_Code_Client]= '" & StrParam(0) & "' AND [AC_TypeAdresse_ID]=" & StrParam(1) & ";"

         TxtAC_Code_Client.DefaultValue = "'" & StrParam(0) & "'"

         TxtAC_TypeAdresse_ID.DefaultValue = StrParam(1)

   End Select

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_Form_BeforeUpdate

   Select Case IsNull(AC_CodePostal)
      Case True

         Cancel = True

         MsgBox "Le Code Postal est obligatoire."

         TxtAC_CodePostal.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case False

   End Select

   Select Case IsNull(AC_Ville)
      Case True

        Cancel = True

        MsgBox "Le nom de la ville est obligatoire."

        TxtAC_Ville.SetFocus

        GoTo Exit_Form_BeforeUpdate

      Case False

    End Select

Exit_Form_BeforeUpdate:

   Exit Sub

Err_Form_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_BeforeUpdate"

   Resume Exit_Form_BeforeUpdate
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdUndo

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdFermer_Click

   End Select

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdOK_Click()
   Dim IntEtat As Integer

   On Error GoTo Err_CmdOK_Click

  On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 3021, 2501

         Err.Clear

         GoTo Exit_CmdOK_Click

      Case Else

         GoTo Err_CmdOK_Click

   End Select

   Select Case Me.NewRecord
      Case True

         Form_BeforeUpdate False

         GoTo Exit_CmdOK_Click

      Case False

   End Select

   On Error GoTo Err_CmdOK_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

   Select Case DonneEtatFormulaire("FrmFicheClient", IntEtat)
      Case True

         Forms!FrmFicheClient.Form.SFrmFicheClientAdresse.Requery

      Case False

   End Select

Exit_CmdOK_Click:

   Exit Sub

Err_CmdOK_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

   Resume Exit_CmdOK_Click
End Sub
