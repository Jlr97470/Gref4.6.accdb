Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Function ImpVisu(ByVal IntVisu As Integer)
   Dim IntRepMsg As Integer
   Dim IntRaportEtat As Integer

   On Error GoTo Err_ImpVisu

   DoCmd.OpenReport "EtaAnalyseSituationDebiteur", acDesign

   Reports("EtaAnalyseSituationDebiteur").RecordSource = SFrmDebiteurAnalyseSituation.Form.RecordSource

   DoCmd.Close acReport, "EtaAnalyseSituationDebiteur", acSavePrompt

   Select Case IntVisu
      Case 0, 2

         On Error Resume Next

         DoCmd.OpenReport "EtaAnalyseSituationDebiteur", IntVisu, , , acWindowNormal

         Select Case Err.Number
            Case 0

            Case 2501

               Err.Clear

            Case Else

               GoTo Err_ImpVisu

         End Select

         On Error GoTo Err_ImpVisu

         Do While DonneEtatRaport("EtaAnalyseSituationDebiteur", IntRaportEtat) = True

            DoEvents

         Loop

      Case 5

         IntRepMsg = MsgBox("Souhaitez-vous exporter les données au format Excel ?" & vbCrLf & "Indiquez OUI pour extraire les données au format excel et NON pour le format texte.", vbYesNo + vbQuestion, "Format d'exportation")

         Select Case IntRepMsg
            Case vbYes

               DoCmd.OutputTo acOutputReport, "EtaAnalyseSituationDebiteur", acFormatXLS, , False

            Case vbNo

               DoCmd.OutputTo acOutputReport, "EtaAnalyseSituationDebiteur", acFormatRTF, , False

         End Select

   End Select

Exit_ImpVisu:

   Exit Function

Err_ImpVisu:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ImpVisu"

   Resume Exit_ImpVisu
End Function

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

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

Private Sub CmdAnalyse_Click()
   Dim StrSQL As String
   Dim StrSQLWhere As String
   Dim BolWhere As Boolean

   On Error GoTo Err_CmdAnalyse_Click

   Screen.MousePointer = 11

   EtiNbDebiteurEcheance.Visible = False

   TxtNbDebiteurEcheance.Visible = False

   SFrmDebiteurAnalyseSituation.Visible = False

   SFrmDebiteurAnalyseSituation1.Visible = False

   StrSQLWhere = "((Déb_Date_proch_rel Is Null) Or Déb_Date_proch_rel=Date())"

   Select Case IsNull(TxtDateDebut)
      Case True

      Case False

         StrSQLWhere = "(Déb_Date_proch_rel Is Null) OR (Déb_Date_proch_rel>=" & ConvDate(TxtDateDebut)

         BolWhere = True

   End Select

   Select Case IsNull(TxtDateFin)
      Case True

         Select Case BolWhere
            Case True

               StrSQLWhere = StrSQLWhere & ")"

            Case False

         End Select

      Case False

         Select Case BolWhere
            Case True

               StrSQLWhere = StrSQLWhere & " AND Déb_Date_proch_rel<=" & ConvDate(TxtDateFin) & ")"

            Case False

               StrSQLWhere = "(Déb_Date_proch_rel Is Null) OR (Déb_Date_proch_rel<=" & ConvDate(TxtDateFin) & ")"

         End Select

   End Select

   StrSQL = "SELECT RELANCE.Rel_Code_Action, RELANCE.Rel_N°, RELANCE.Rel_Nom, nz(Sum(iif(" & StrSQLWhere & ",[SelRelanceNbDebiteur].[NbDebiteur],0)),0) AS NbDebiteur " & _
                 "FROM RELANCE LEFT JOIN SelRelanceNbDebiteur ON (RELANCE.Rel_Code_Action = SelRelanceNbDebiteur.CG_New_Code_Action) AND (RELANCE.Rel_N° = SelRelanceNbDebiteur.Déb_New_niv_relance) " & _
                 "WHERE RELANCE.Rel_Code_Action In (1,2,3) " & _
                 "GROUP BY RELANCE.Rel_Code_Action, RELANCE.Rel_N°, RELANCE.Rel_Nom;"

   SFrmDebiteurAnalyseSituation.Form.RecordSource = StrSQL

   SFrmDebiteurAnalyseSituation.Form.Filter = "Rel_Code_Action<2"

   SFrmDebiteurAnalyseSituation.Form.FilterOn = True

   SFrmDebiteurAnalyseSituation1.Form.RecordSource = StrSQL

   SFrmDebiteurAnalyseSituation1.Form.Filter = "Rel_Code_Action>=2"

   SFrmDebiteurAnalyseSituation1.Form.FilterOn = True

   EtiNbDebiteurEcheance.Visible = True

   TxtNbDebiteurEcheance.Visible = True

   SFrmDebiteurAnalyseSituation.Visible = True

   SFrmDebiteurAnalyseSituation1.Visible = True

   Screen.MousePointer = 0

Exit_CmdAnalyse_Click:

   Exit Sub

Err_CmdAnalyse_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdAnalyse_Click"

   Resume Exit_CmdAnalyse_Click
End Sub
