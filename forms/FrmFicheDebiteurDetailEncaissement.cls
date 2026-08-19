Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private iMonnaieVisu As Integer

Public Sub Inverseur(ByVal IntMonnaie As Integer)
    'Cette fonction permet de visualiser le formulaire dans une monnaie particulière (argMonnaie)
   On Error GoTo Err_Inverseur

   Select Case IntMonnaie
       Case cFRANC

           TxtEnc_Montant_m.ControlSource = "Enc_Montant_m1"

           TxtEnc_Prin_m.ControlSource = "Enc_Prin_m1"

           TxtEnc_Frais_m.ControlSource = "Enc_Frais_m1"

       Case cEURO

           TxtEnc_Montant_m.ControlSource = "Enc_Montant_m2"

           TxtEnc_Prin_m.ControlSource = "Enc_Prin_m2"

           TxtEnc_Frais_m.ControlSource = "Enc_Frais_m2"

   End Select

   iMonnaieVisu = IntMonnaie

Exit_Inverseur:

    Exit Sub

Err_Inverseur:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Inverseur"

    Resume Exit_Inverseur
End Sub

Private Sub Form_Open(Cancel As Integer)
   Dim StrParam() As String

   On Error GoTo Err_Form_Open

   StrParam = Split(Nz(Me.OpenArgs), ";")

   Select Case UBound(StrParam)
      Case Is < 1

         Me.RecordSource = "SELECT ENCAISSEMENT.* FROM ENCAISSEMENT;"

         Inverseur cEURO

      Case Else

         Me.RecordSource = "SELECT ENCAISSEMENT.* FROM ENCAISSEMENT WHERE [Enc_Id]= " & StrParam(0) & ";"

         Inverseur Nz(StrParam(1), cEURO)

   End Select

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
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

   DoCmd.Close acForm, "FrmFicheDebiteurDetailEncaissement", acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdOK_Click()

   On Error GoTo Err_CmdOK_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

         GoTo Exit_CmdOK_Click

      Case Else

         GoTo Err_CmdOK_Click

   End Select

   On Error GoTo Err_CmdOK_Click

   DoCmd.Close acForm, "FrmFicheDebiteurDetailEncaissement", acSavePrompt

Exit_CmdOK_Click:

    Exit Sub

Err_CmdOK_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

    Resume Exit_CmdOK_Click
End Sub
