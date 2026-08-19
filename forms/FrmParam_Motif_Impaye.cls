Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   Me.Requery

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Load"

   Resume Exit_Form_Load
End Sub

Private Sub CmbMIG_Grille_AfterUpdate()

   On Error GoTo Err_CmbMIG_Grille_AfterUpdate

   Me.Requery

Exit_CmbMIG_Grille_AfterUpdate:

   Exit Sub

Err_CmbMIG_Grille_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbMIG_Grille_AfterUpdate"

   Resume Exit_CmbMIG_Grille_AfterUpdate
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdMotif_ImpayeGrille_Click()

   On Error GoTo Err_CmdMotif_ImpayeGrille_Click

   DoCmd.OpenForm "FrmParam_Motif_ImpayeGrille", acNormal, , , , acDialog

   CmbMIG_Grille.Requery

Exit_CmdMotif_ImpayeGrille_Click:

   Exit Sub

Err_CmdMotif_ImpayeGrille_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdMotif_ImpayeGrille_Click"

   Resume Exit_CmdMotif_ImpayeGrille_Click
End Sub
