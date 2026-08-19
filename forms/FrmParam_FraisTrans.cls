Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private iMonnaieVisu As Integer 'Monnaie dans laquelle est visualisée la fiche en cours

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   iMonnaieVisu = cEURO

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
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

Private Sub CmdInverseur_Click()

   On Error GoTo Err_CmdInverseur_Click

   Select Case iMonnaieVisu
       Case cFRANC

           iMonnaieVisu = cEURO

       Case cEURO

           iMonnaieVisu = cFRANC

   End Select

   Me![SFrmParam_FraisTrans1].Form.Inverseur iMonnaieVisu

   Me![SFrmParam_FraisTrans2].Form.Inverseur iMonnaieVisu

Exit_CmdInverseur_Click:

   Exit Sub

Err_CmdInverseur_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdInverseur_Click"

   Resume Exit_CmdInverseur_Click
End Sub

Private Sub CmdFraisTransGrille_Click()

   On Error GoTo Err_CmdFraisTransGrille_Click

   DoCmd.OpenForm "FrmParam_FraisTransGrille", acNormal, , , , acDialog

   CmbFrt_Grille.Requery

Exit_CmdFraisTransGrille_Click:

   Exit Sub

Err_CmdFraisTransGrille_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFraisTransGrille_Click"

   Resume Exit_CmdFraisTransGrille_Click
End Sub

Private Sub CmbFrt_Grille_AfterUpdate()

   On Error GoTo Err_CmbFrt_Grille_AfterUpdate

    Me![SFrmParam_FraisTrans1].Form.Inverseur iMonnaieVisu

    Me![SFrmParam_FraisTrans2].Form.Inverseur iMonnaieVisu

Exit_CmbFrt_Grille_AfterUpdate:

   Exit Sub

Err_CmbFrt_Grille_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbFrt_Grille_AfterUpdate"

   Resume Exit_CmbFrt_Grille_AfterUpdate
End Sub
