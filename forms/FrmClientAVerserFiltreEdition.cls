Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Function ImpVisu(ByVal IntVisu As Integer)
   Dim StrNomEtat As String
   Dim StrSQLWhere As String
   Dim BolWhere As Boolean
   Dim StrDateDebut As String
   Dim StrDateFin As String
   Dim IntRaportEtat As Integer
   Dim IntRepMsg As Integer

   On Error GoTo Err_ImpVisu

   StrDateDebut = Nz(TxtDateDebut)

   StrDateFin = Nz(TxtDateFin)

   Select Case StrDateDebut
      Case ""

      Case Else

         StrSQLWhere = "Enc_Date>=" & ConvDate(StrDateDebut)

         BolWhere = True

   End Select

   Select Case StrDateFin
      Case ""

      Case Else

         Select Case BolWhere
            Case True

               StrSQLWhere = StrSQLWhere & " AND Enc_Date<=" & ConvDate(StrDateFin)

            Case False

               StrSQLWhere = "Enc_Date<=" & ConvDate(StrDateFin)

         End Select

         BolWhere = True

   End Select

   Select Case CadEtat
      Case 1

         StrNomEtat = "EtaListeClientAVerserDetail"

      Case 2

         StrNomEtat = "EtaListeClientAVerserSynthese"

   End Select

   DoCmd.OpenReport StrNomEtat, acDesign

   Reports(StrNomEtat).Filter = StrSQLWhere

   Reports(StrNomEtat).FilterOn = True

   DoCmd.Close acReport, StrNomEtat, acSavePrompt

   Select Case IntVisu
      Case 0, 2

         Select Case CadEtat
            Case 1

               On Error Resume Next

               DoCmd.OpenReport StrNomEtat, IntVisu, , "Enc_Montant_mo=" & cFRANC & IIf(StrSQLWhere = "", "", " AND ") & StrSQLWhere, acWindowNormal, cFRANC & ";"

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

               DoCmd.OpenReport StrNomEtat, IntVisu, , "Enc_Montant_mo=" & cEURO & IIf(StrSQLWhere = "", "", " AND ") & StrSQLWhere, acWindowNormal, cEURO & ";"

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

            Case 2

               On Error Resume Next

               DoCmd.OpenReport StrNomEtat, IntVisu, , , acWindowNormal

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

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub
