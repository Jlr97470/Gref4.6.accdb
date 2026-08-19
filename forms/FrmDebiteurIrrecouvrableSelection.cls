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

           TxtMontant_m.ControlSource = "=Round(Nz(Dlookup(" & Chr(34) & "Déb_Mont_Prin_m1" & Chr(34) & "," & Chr(34) & "DEBITEUR" & Chr(34) & "," & Chr(34) & "Déb_Code='" & Chr(34) & " & [TIrr_Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)-Round(Nz(DSum(" & Chr(34) & "Enc_Prin_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [TIrr_Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

       Case cEURO

           TxtMontant_m.ControlSource = "=Round(Nz(Dlookup(" & Chr(34) & "Déb_Mont_Prin_m2" & Chr(34) & "," & Chr(34) & "DEBITEUR" & Chr(34) & "," & Chr(34) & "Déb_Code='" & Chr(34) & " & [TIrr_Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)-Round(Nz(DSum(" & Chr(34) & "Enc_Prin_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [TIrr_Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

   End Select

   CmbMonnaie = IntMonnaie

   Me.Filter = "Déb_MonnaieEnCours = " & IntMonnaie

   Me.FilterOn = True

   iMonnaieVisu = IntMonnaie

Exit_Inverseur:

    Exit Sub

Err_Inverseur:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Inverseur"

    Resume Exit_Inverseur
End Sub

Public Function ToutRien(ByVal BolSelection As Boolean)
   Dim StrSQL As String

   On Error GoTo Err_ToutRien

   Screen.MousePointer = 11

   StrSQL = "UPDATE TMP_IRRECOUVRABLE SET TIrr_Editer = " & IIf(BolSelection, "True", "False") & " WHERE TIrr_Déb_Code In (SELECT TMP_IRRECOUVRABLE.TIrr_Déb_Code " & _
                 "FROM DEBITEUR INNER JOIN TMP_IRRECOUVRABLE ON DEBITEUR.Déb_Code = TMP_IRRECOUVRABLE.TIrr_Déb_Code " & _
                 "WHERE DEBITEUR.Déb_MonnaieEnCours=" & iMonnaieVisu

   Select Case IsNull(TxtDateDebut)
      Case True

      Case False

         StrSQL = StrSQL & " AND Déb_Date_der_rel>=" & ConvDate(TxtDateDebut)

   End Select

   Select Case IsNull(TxtDateFin)
      Case True

      Case False

         StrSQL = StrSQL & " AND Déb_Date_der_rel<=" & ConvDate(TxtDateFin)

   End Select

   StrSQL = StrSQL & ";);"

   CurrentDb.Execute StrSQL, dbFailOnError

   Me.Requery

   Screen.MousePointer = 0

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

Private Sub CmbMonnaie_AfterUpdate()

   On Error GoTo Err_CmbMonnaie_AfterUpdate

   Inverseur CmbMonnaie

Exit_CmbMonnaie_AfterUpdate:

   Exit Sub

Err_CmbMonnaie_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbMonnaie_AfterUpdate"

   Resume Exit_CmbMonnaie_AfterUpdate
End Sub

Private Sub ChkTIrr_Editer_AfterUpdate()

   On Error GoTo Err_ChkTIrr_Editer_AfterUpdate

   DoCmd.RunCommand acCmdSaveRecord

Exit_ChkTIrr_Editer_AfterUpdate:

   Exit Sub

Err_ChkTIrr_Editer_AfterUpdate:

   MsgBox Err.description & " " & Err.Number, vbCritical, stSource & " : ChkTIrr_Editer_AfterUpdate"

   Resume Exit_ChkTIrr_Editer_AfterUpdate
End Sub

Private Sub CmdOK_Click()

   On Error GoTo Err_CmdOK_Click

   CurrentDb.Execute "UPDATE TMP_IRRECOUVRABLE SET TIrr_Editer = False", dbFailOnError

   Me.Requery

Exit_CmdOK_Click:

   Exit Sub

Err_CmdOK_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOK_Click"

   Resume Exit_CmdOK_Click
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   CurrentDb.Execute "DELETE TMP_IRRECOUVRABLE.* FROM TMP_IRRECOUVRABLE", dbFailOnError

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdImprimer_Click()

   On Error GoTo Err_CmdImprimer_Click

   DoCmd.OpenForm "FrmDebiteurIrrecouvrableSelectionFiltreEdition", acNormal, , , , acDialog, "Déb_Nom;TIrr_Déb_Code;Déb_Réf1_Client"

Exit_CmdImprimer_Click:

   Exit Sub

Err_CmdImprimer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdImprimer_Click"

   Resume Exit_CmdImprimer_Click
End Sub
