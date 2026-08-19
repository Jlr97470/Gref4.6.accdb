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

Private Sub Form_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_Form_AfterUpdate

   CG_Parent_Id = CmbGroupe

Exit_Form_AfterUpdate:

   Exit Sub

Err_Form_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_AfterUpdate"

   Resume Exit_Form_AfterUpdate

End Sub

Private Sub CmbGroupe_AfterUpdate()

   On Error GoTo Err_CmbGroupe_AfterUpdate

   Me.AllowAdditions = Not IsNull(CmbGroupe)

   Me.Requery

Exit_CmbGroupe_AfterUpdate:

   Exit Sub

Err_CmbGroupe_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbGroupe_AfterUpdate"

   Resume Exit_CmbGroupe_AfterUpdate
End Sub

Private Sub CmbGroupe_GotFocus()

   On Error GoTo Err_CmbGroupe_GotFocus

   DoCmd.RunCommand acCmdSaveRecord

Exit_CmbGroupe_GotFocus:

   Exit Sub

Err_CmbGroupe_GotFocus:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbGroupe_GotFocus"

   Resume Exit_CmbGroupe_GotFocus
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.RunCommand acCmdSaveRecord

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub
