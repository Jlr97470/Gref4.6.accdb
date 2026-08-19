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

           TxtDéb_Mont_Prin_m.ControlSource = "Déb_Mont_Prin_m1"

           TxtDéb_Enc_Prin_m.ControlSource = "Déb_Enc_Prin_m1"

       Case cEURO

           TxtDéb_Mont_Prin_m.ControlSource = "Déb_Mont_Prin_m2"

           TxtDéb_Enc_Prin_m.ControlSource = "Déb_Enc_Prin_m2"

   End Select

   CmbMonnaie = IntMonnaie

   Me.Filter = "Déb_MonnaieEnCours = " & IntMonnaie

   Me.FilterOn = True

   iMonnaieVisu = IntMonnaie

Exit_Inverseur:

    Exit Sub

Err_Inverseur:

    MsgBox Err.Number & " : " & Err.description, vbCritical, stSource & " : Inverseur"

    Resume Exit_Inverseur
End Sub

Private Function ImpVisu(ByVal IntVisu As Integer)

    On Error GoTo Err_ImpVisu

    DoCmd.OpenReport "EtaListeDebiteurEncaissementIncorrect", IntVisu

Exit_ImpVisu:

    Exit Function

Err_ImpVisu:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & ""

    Resume Exit_ImpVisu
End Function

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   Inverseur cEURO

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
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

   Me.Requery

Exit_CmdOK_Click:

   Exit Sub

Err_CmdOK_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOK_Click"

   Resume Exit_CmdOK_Click
End Sub

Private Sub CmdDetail_Click()
   Dim IntEtat As Integer

   On Error GoTo Err_CmdDetail_Click

   Select Case DonneEtatFormulaire("FrmFicheDebiteur", IntEtat)
      Case True

         Forms("FrmFicheDebiteur").Filter = "Déb_Code='" & Déb_Code & "'"

         Forms("FrmFicheDebiteur").FilterOn = True

      Case False

         DoCmd.OpenForm "FrmFicheDebiteur", acNormal, , "Déb_Code='" & Déb_Code & "'", , acWindowNormal

   End Select

Exit_CmdDetail_Click:

   Exit Sub

Err_CmdDetail_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdDetail_Click"

   Resume Exit_CmdDetail_Click
End Sub
