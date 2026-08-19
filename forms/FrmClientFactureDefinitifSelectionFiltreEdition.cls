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

         StrNomEtat = "EtaCourrierFacture"

         IntGroupLevel = 3

      Case 2

         StrNomEtat = "EtaCourrierFacture"

         IntGroupLevel = 3

      Case 3

         StrNomEtat = "EtaSyntheseFactureGroupe"

         IntGroupLevel = -1

      Case 4

         StrNomEtat = "EtaSyntheseFactureSousGroupe"

         IntGroupLevel = -1

   End Select

   Select Case IntGroupLevel
      Case -1

      Case Else

         DoCmd.OpenReport StrNomEtat, acDesign

         Select Case CadEtat
            Case 1

               Reports(StrNomEtat).GroupLevel(0).ControlSource = "Fac_Date"

            Case 2

               Reports(StrNomEtat).GroupLevel(0).ControlSource = "Cli_Groupe"

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

         On Error Resume Next

         DoCmd.OpenReport StrNomEtat, IntVisu, , "Fac_Montant_mo=" & cFRANC, acWindowNormal, cFRANC & ";"

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

         DoCmd.OpenReport StrNomEtat, IntVisu, , "Fac_Montant_mo=" & cEURO, , cEURO & ";"

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

               Select Case CadEtat
                  Case 1, 2

                     StrSQL = "UPDATE FACTURE SET Fac_Deja_Editer = True, Fac_Editer = False WHERE Fac_Editer=True;"

                     CurrentDb.Execute StrSQL, dbFailOnError

                  Case 3, 4

               End Select

            Case 2

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
