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

         TxtResteAverser_m.ControlSource = "=Round(DLookUp(" & Chr(34) & "Déb_Mont_Prin_m1" & Chr(34) & "," & Chr(34) & "DEBITEUR" & Chr(34) & "," & Chr(34) & "Déb_Code='" & Chr(34) & " & [TRel_Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),2)-Round(nz(DSum(" & Chr(34) & "Enc_Prin_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [TRel_Déb_Code] & " & Chr(34) & "'" & Chr(34) & ")),2)"

         TxtDéb_Frais_m.ControlSource = "=DLookUp(" & Chr(34) & "Déb_Frais_m1" & Chr(34) & "," & Chr(34) & "DEBITEUR" & Chr(34) & "," & Chr(34) & "Déb_Code='" & Chr(34) & " & [TRel_Déb_Code] & " & Chr(34) & "'" & Chr(34) & ")"

      Case cEURO

         TxtResteAverser_m.ControlSource = "=Round(DLookUp(" & Chr(34) & "Déb_Mont_Prin_m2" & Chr(34) & "," & Chr(34) & "DEBITEUR" & Chr(34) & "," & Chr(34) & "Déb_Code='" & Chr(34) & " & [TRel_Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),2)-Round(nz(DSum(" & Chr(34) & "Enc_Prin_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [TRel_Déb_Code] & " & Chr(34) & "'" & Chr(34) & ")),2)"

         TxtDéb_Frais_m.ControlSource = "=DLookUp(" & Chr(34) & "Déb_Frais_m2" & Chr(34) & "," & Chr(34) & "DEBITEUR" & Chr(34) & "," & Chr(34) & "Déb_Code='" & Chr(34) & " & [TRel_Déb_Code] & " & Chr(34) & "'" & Chr(34) & ")"

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

Private Function ToutRien(ByVal BolSelection As Boolean)
   Dim StrSQL As String

   On Error GoTo Err_ToutRien

   StrSQL = "UPDATE TMP_RELANCE INNER JOIN DEBITEUR ON TMP_RELANCE.TRel_Déb_Code = DEBITEUR.Déb_Code SET TRel_Editer = " & IIf(BolSelection, "True", "False") & " WHERE Déb_MonnaieEnCours = " & iMonnaieVisu

   CurrentDb.Execute StrSQL, dbFailOnError

   Me.Requery

Exit_ToutRien:

   Exit Function

Err_ToutRien:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ToutRien"

   Resume Exit_ToutRien
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

Private Sub Form_Load()

   On Error GoTo Err_Form_Load

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Load"

   Resume Exit_Form_Load
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

Private Sub ChkTRel_Editer_AfterUpdate()

   On Error GoTo Err_ChkTRel_Editer_AfterUpdate

   DoCmd.RunCommand acCmdSaveRecord

Exit_ChkTRel_Editer_AfterUpdate:

   Exit Sub

Err_ChkTRel_Editer_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ChkTRel_Editer_AfterUpdate"

   Resume Exit_ChkTRel_Editer_AfterUpdate
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & ": CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdImprimer_Click()

    On Error GoTo Err_CmdImprimer_Click

    DoCmd.OpenForm "FrmDebiteurRelanceSelectionFiltreEdition", acNormal, , , , acDialog

Exit_CmdImprimer_Click:

   Exit Sub

Err_CmdImprimer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdImprimer_Click"

   Resume Exit_CmdImprimer_Click
End Sub

Private Sub CmdSupprimer_Click()
    Dim StrSQL As String
    Dim IntMes As Integer

    On Error GoTo Err_CmdSupprimer_Click

    IntMes = MsgBox("Suppression des relances sélectionnées." & vbCrLf & "Voulez-vous continuer ?", vbYesNo + vbDefaultButton2 + vbQuestion, stSource)

    Select Case IntMes
      Case vbYes

        StrSQL = "DELETE * FROM TMP_RELANCE WHERE TRel_Editer = True;"

        CurrentDb.Execute StrSQL, dbFailOnError

        Me.Requery

      Case Else

    End Select

Exit_CmdSupprimer_Click:

   Exit Sub

Err_CmdSupprimer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & ": CmdSupprimer_Click"

   Resume Exit_CmdSupprimer_Click
End Sub
