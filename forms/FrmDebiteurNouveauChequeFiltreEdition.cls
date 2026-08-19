Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Function ImpVisu(ByVal IntVisu As Integer)
   Dim IntRepMsg As Integer
   Dim IntRaportEtat As Integer
   Dim StrSQL As String

   On Error GoTo Err_ImpVisu

   Select Case IntVisu
      Case 0, 2

         On Error Resume Next

         DoCmd.OpenReport "EtaCourrierNouveauCheque", IntVisu, , , acWindowNormal

         Select Case Err.Number
            Case 0

            Case 2501

               Err.Clear

            Case Else

               GoTo Err_ImpVisu

         End Select

         On Error GoTo Err_ImpVisu

         Do While DonneEtatRaport("EtaCourrierNouveauCheque", IntRaportEtat) = True

            DoEvents

         Loop

         Select Case IntVisu
            Case 0

               StrSQL = "UPDATE CODE_GESTION INNER JOIN DEBITEUR ON CODE_GESTION.CG_Code = DEBITEUR.Déb_Situation SET DEBITEUR.Déb_Date_der_rel = Date() WHERE (((DEBITEUR.Déb_Date_der_rel) Is Null) AND ((CODE_GESTION.CG_Code_Action)=5));"

               CurrentDb.Execute StrSQL, dbFailOnError

            Case 2

         End Select

      Case 5

         IntRepMsg = MsgBox("Souhaitez-vous exporter les données au format Excel ?" & vbCrLf & "Indiquez OUI pour extraire les données au format excel et NON pour le format texte.", vbYesNo + vbQuestion, "Format d'exportation")

         Select Case IntRepMsg
            Case vbYes

               DoCmd.OutputTo acOutputReport, "EtaCourrierNouveauCheque", acFormatXLS, , False

            Case vbNo

               DoCmd.OutputTo acOutputReport, "EtaCourrierNouveauCheque", acFormatRTF, , False

         End Select

   End Select

   Me.Requery

Exit_ImpVisu:

   Exit Function

Err_ImpVisu:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ImpVisu"

   Resume Exit_ImpVisu
End Function

Private Sub Form_Current()

   On Error GoTo Err_Form_Current

   Select Case TxtQuantite
      Case 0

         CmdImprimer.Enabled = False

         CmdVisualiser.Enabled = False

         CmdExporter.Enabled = False

      Case Else

   End Select

Exit_Form_Current:

   Exit Sub

Err_Form_Current:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Current"

   Resume Exit_Form_Current
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
