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

         StrNomEtat = "EtaCourrierReglementForfait"

         IntGroupLevel = 2

      Case 2

         StrNomEtat = "EtaCourrierReglementPourcentage"

         IntGroupLevel = 3

   End Select

   Select Case IntGroupLevel
      Case -1

      Case Else

         DoCmd.OpenReport StrNomEtat, acDesign

         Select Case IntVisu
            Case 0, 2

               Reports(StrNomEtat).TxtDéb_Date_Chèque.Visible = False

            Case 5

               Reports(StrNomEtat).TxtDéb_Date_Chèque.Visible = True

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

         DoCmd.OpenReport StrNomEtat, IntVisu, , "Reg_Montant_mo=" & cFRANC, , cFRANC & ";"

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

         DoCmd.OpenReport StrNomEtat, IntVisu, , "Reg_Montant_mo=" & cEURO, , cEURO & ";"

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
                  Case 1

                     StrSQL = "UPDATE REGLEMENT SET Reg_Deja_Editer = True, Reg_Editer = False WHERE Reg_Cli_Pourcentage = 0 AND Reg_Editer=True;"

                  Case 2

                     StrSQL = "UPDATE REGLEMENT SET Reg_Deja_Editer = True, Reg_Editer = False WHERE Reg_Cli_Pourcentage > 0 AND Reg_Editer=True;"

               End Select

                CurrentDb.Execute StrSQL, dbFailOnError

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
