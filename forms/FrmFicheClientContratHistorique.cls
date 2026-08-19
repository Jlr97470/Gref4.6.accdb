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

           TxtCC_Mont_Hon_m.ControlSource = "CC_Mont_Hon_m1"

       Case cEURO

           TxtCC_Mont_Hon_m.ControlSource = "CC_Mont_Hon_m2"

   End Select

   Me.Filter = "CC_Mont_Hon_mo = " & IntMonnaie

   Me.FilterOn = True

   CmbMonnaie = IntMonnaie

   iMonnaieVisu = IntMonnaie

Exit_Inverseur:

    Exit Sub

Err_Inverseur:

    MsgBox Err.Number & " " & Err.Number, vbCritical, stSource & " : Inverseur"

    Resume Exit_Inverseur
End Sub

Private Sub Form_Open(Cancel As Integer)
   Dim StrParam() As String

   On Error GoTo Err_Form_Open

   StrParam = Split(Nz(Me.OpenArgs), ";")

   Select Case UBound(StrParam)
      Case Is < 1

         Me.RecordSource = "SELECT CLIENT_CONTRAT.* FROM CLIENT_CONTRAT;"

         Inverseur cEURO

      Case Else

         Me.RecordSource = "SELECT CLIENT_CONTRAT.* FROM CLIENT_CONTRAT WHERE CC_Code_Client='" & StrParam(0) & "';"

         Inverseur Nz(StrParam(1), cEURO)

   End Select

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
