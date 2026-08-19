Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Public StrOrigine_Mot_de_Passe As String

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

    StrOrigine_Mot_de_Passe = TxtMot_de_passe

    Select Case CadMode_Calcul_Frais
      Case 1

        EtiTx_Interet_legal.ForeColor = 8421504

        TxtTx_Interet_legal.Enabled = False

        EtiFraisAccessoires.ForeColor = 8421504

        TxtFrais_Accessoires.Enabled = False

      Case 2, 3

        EtiTx_Interet_legal.ForeColor = 0

        TxtTx_Interet_legal.Enabled = True

        EtiFraisAccessoires.ForeColor = 0

        TxtFrais_Accessoires.Enabled = True

    End Select

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub CadMode_Calcul_Frais_AfterUpdate()

   On Error GoTo Err_CadMode_Calcul_Frais_AfterUpdate

    Select Case CadMode_Calcul_Frais
      Case 1

        EtiTx_Interet_legal.ForeColor = 8421504

        TxtTx_Interet_legal.Enabled = False

        EtiFraisAccessoires.ForeColor = 8421504

        TxtFrais_Accessoires.Enabled = False

      Case 2, 3

        EtiTx_Interet_legal.ForeColor = 0

        TxtTx_Interet_legal.Enabled = True

        EtiFraisAccessoires.ForeColor = 0

        TxtFrais_Accessoires.Enabled = True

    End Select

Exit_CadMode_Calcul_Frais_AfterUpdate:

   Exit Sub

Err_CadMode_Calcul_Frais_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CadMode_Calcul_Frais_AfterUpdate"

   Resume Exit_CadMode_Calcul_Frais_AfterUpdate
End Sub

Private Sub TxtFrais_Déplacement_m1_AfterUpdate()

   On Error GoTo Err_TxtFrais_Déplacement_m1_AfterUpdate

   TxtFrais_Déplacement_m2 = Round(TxtFrais_Déplacement_m1 * DonnerTauxConvertion(cFRANC), 2)

Exit_TxtFrais_Déplacement_m1_AfterUpdate:

   Exit Sub

Err_TxtFrais_Déplacement_m1_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtFrais_Déplacement_m1_AfterUpdate"

   Resume Exit_TxtFrais_Déplacement_m1_AfterUpdate
End Sub

Private Sub TxtFrais_Déplacement_m2_AfterUpdate()

   On Error GoTo Err_TxtFrais_Déplacement_m2_AfterUpdate

   TxtFrais_Déplacement_m1 = Round(TxtFrais_Déplacement_m2 * DonnerTauxConvertion(cEURO), 2)

Exit_TxtFrais_Déplacement_m2_AfterUpdate:

   Exit Sub

Err_TxtFrais_Déplacement_m2_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtFrais_Déplacement_m2_AfterUpdate"

   Resume Exit_TxtFrais_Déplacement_m2_AfterUpdate
End Sub

Private Sub TxtMot_de_passe_AfterUpdate()

   On Error GoTo Err_TxtMot_de_passe_AfterUpdate

   Select Case InputBox("Veuillez confirmer le nouveau mot de passe :", "Nouveau mot de passe")
     Case TxtMot_de_passe

       MsgBox "Le nouveau mot de passe est enregistré.", 64

       StrOrigine_Mot_de_Passe = TxtMot_de_passe

     Case Else

       MsgBox "Mot de passe incorrect !" & Chr(10) & "Le nouveau mot de passe n'a pas été enregistré.", 48

       TxtMot_de_passe = StrOrigine_Mot_de_Passe

   End Select

   TxtMot_de_passe.Locked = True

Exit_TxtMot_de_passe_AfterUpdate:

   Exit Sub

Err_TxtMot_de_passe_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtMot_de_passe_AfterUpdate"

   Resume Exit_TxtMot_de_passe_AfterUpdate
End Sub

Private Sub CmdTypeAdresse_Click()

   On Error GoTo Err_CmdTypeAdresse_Click

   DoCmd.OpenForm "FrmParam_TypeAdresse", acNormal, , , , acDialog

   CmbAdresse_Courrier.Requery

Exit_CmdTypeAdresse_Click:

   Exit Sub

Err_CmdTypeAdresse_Click:

   MsgBox Err.Number & "  " & Err.description, vbCritical, stSource & " : CmdTypeAdresse_Click"

   Resume Exit_CmdTypeAdresse_Click
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

Private Sub TxtMot_de_passe_DblClick(Cancel As Integer)

   On Error GoTo Err_TxtMot_de_passe_DblClick

   Select Case InputBox("Cette zone est protégée par un mot de passe :", "Mot de passe")
      Case TxtMot_de_passe

         TxtMot_de_passe.Locked = False

      Case Else

          TxtMot_de_passe.Locked = True

    End Select

Exit_TxtMot_de_passe_DblClick:

   Exit Sub

Err_TxtMot_de_passe_DblClick:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtMot_de_passe_DblClick"

   Resume Exit_TxtMot_de_passe_DblClick
End Sub
