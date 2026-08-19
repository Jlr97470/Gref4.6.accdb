Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private StrChpAlpha As String
Private StrChpCode As String
Private StrChpRef As String

Private Function ImpVisu(ByVal IntVisu As Integer)
   Dim StrSQL As String
   Dim StrNomEtat As String
   Dim IntGroupLevel As Integer
   Dim IntRepMsg As Integer
   Dim IntRaportEtat As Integer

   On Error GoTo Err_ImpVisu

   Select Case CadEtat
      Case 1

         StrNomEtat = "EtaLettrePlainte"

         IntGroupLevel = 0

      Case 2, 3

         StrNomEtat = "EtaCourrierAttestationIrrecouvrabilite"

         IntGroupLevel = 2

   End Select

   Select Case IntGroupLevel
      Case -1

      Case Else

         DoCmd.OpenReport StrNomEtat, acDesign

         Select Case CadEtat
            Case 1

            Case 2

               Reports(StrNomEtat).GroupLevel(0).ControlSource = "Cli_Code"

            Case 3

               Reports(StrNomEtat).GroupLevel(0).ControlSource = "TIrr_Déb_Code"

         End Select

         Select Case CadChoix
            Case 1

               Reports(StrNomEtat).GroupLevel(IntGroupLevel).ControlSource = StrChpAlpha

            Case 2

               Reports(StrNomEtat).GroupLevel(IntGroupLevel).ControlSource = StrChpCode

            Case 3

               Reports(StrNomEtat).GroupLevel(IntGroupLevel).ControlSource = StrChpRef

         End Select

         DoCmd.Close acReport, StrNomEtat, acSavePrompt

   End Select

   Select Case IntVisu
      Case 0, 2

         Select Case CadEtat
            Case 1

               On Error Resume Next

               DoCmd.OpenReport StrNomEtat, IntVisu

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_ImpVisu

               End Select

               On Error GoTo Err_ImpVisu

               Do While DonneEtatRaport(StrNomEtat, IntRaportEtat) = True

                  DoEvents

               Loop

            Case 2, 3

               On Error Resume Next

               DoCmd.OpenReport StrNomEtat, IntVisu, , "Déb_MonnaieEnCours=" & cFRANC, , DLookup("Site", "PARAMETRE_GREF") & ";" & cFRANC

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_ImpVisu

               End Select

               On Error GoTo Err_ImpVisu

               Do While DonneEtatRaport(StrNomEtat, IntRaportEtat) = True

                  DoEvents

               Loop

               On Error Resume Next

               DoCmd.OpenReport StrNomEtat, IntVisu, , "Déb_MonnaieEnCours=" & cEURO, , DLookup("Site", "PARAMETRE_GREF") & ";" & cEURO

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_ImpVisu

               End Select

               On Error GoTo Err_ImpVisu

               Do While DonneEtatRaport(StrNomEtat, IntRaportEtat) = True

                  DoEvents

               Loop

               Select Case IntVisu
                  Case 0

                     StrSQL = "UPDATE DEBITEUR INNER JOIN [TMP_IRRECOUVRABLE] ON DEBITEUR.Déb_Code = [TMP_IRRECOUVRABLE].TIrr_Déb_Code SET DEBITEUR.Déb_Date_der_rel = Date(), [TMP_IRRECOUVRABLE].TIrr_Deja_Editer = True WHERE [TMP_IRRECOUVRABLE].TIrr_Editer = True AND IsNull(DEBITEUR.Déb_Date_der_rel);"

                     CurrentDb.Execute StrSQL, dbFailOnError

                  Case 2

               End Select

         End Select

      Case 5

         IntRepMsg = MsgBox("Souhaitez-vous exporter les données au format Excel ?" & vbCrLf & "Indiquez OUI pour extraire les données au format excel et NON pour le format texte.", vbYesNo + vbQuestion, "Format d'exportation")

         Select Case IntRepMsg
            Case vbYes

               DoCmd.OutputTo acOutputReport, StrNomEtat, acFormatXLS, , False

            Case vbNo

               DoCmd.OutputTo acOutputReport, StrNomEtat, acFormatRTF, , False

         End Select

   End Select

Exit_ImpVisu:

   Exit Function

Err_ImpVisu:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ImpVisu"

   Resume Exit_ImpVisu
End Function

Private Sub Form_Open(Cancel As Integer)
   Dim StrParam() As String

   On Error GoTo Err_Form_Open

   StrParam = Split(Nz(Me.OpenArgs), ";")

   Select Case UBound(StrParam)
      Case Is < 1

         Cancel = True

         GoTo Exit_Form_Open

      Case Else

   End Select

   StrChpAlpha = StrParam(0)

   StrChpCode = StrParam(1)

   StrChpRef = StrParam(2)

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
