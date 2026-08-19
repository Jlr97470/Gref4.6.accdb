Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Function ImpVisu(ByVal IntVisu As Integer)
   Dim StrCli_Code As String
   Dim StrDéb_Date_Cré As String
   Dim StrSQLWhere As String
   Dim IntRepMsg As Integer
   Dim IntRaportEtat As Integer
   Dim BolWhere As Boolean

   On Error GoTo Err_ImpVisu

   BolWhere = False

   StrCli_Code = Nz(TxtCli_Code)

   StrDéb_Date_Cré = Nz(TxtDéb_Date_Cré)

   Select Case StrCli_Code
      Case ""

      Case Else

         StrSQLWhere = StrSQLWhere & IIf(BolWhere, " AND ", "") & "Cli_Code='" & StrCli_Code & "'"

         BolWhere = True

   End Select

   Select Case StrDéb_Date_Cré
      Case ""

      Case Else

         StrSQLWhere = StrSQLWhere & IIf(BolWhere, " AND ", "") & "Déb_Date_Cré=" & ConvDate(StrDéb_Date_Cré)

         BolWhere = True

   End Select

   DoCmd.OpenReport "EtaCompteRenduSaisie", acDesign

   Reports("EtaCompteRenduSaisie").Filter = StrSQLWhere

   Reports("EtaCompteRenduSaisie").FilterOn = True

   DoCmd.Close acReport, "EtaCompteRenduSaisie", acSavePrompt

   Select Case IntVisu
      Case 0, 2

         On Error Resume Next

         DoCmd.OpenReport "EtaCompteRenduSaisie", IntVisu, , "Déb_MonnaieEnCours=" & cFRANC & IIf(StrSQLWhere = "", "", " AND ") & StrSQLWhere, acWindowNormal, cFRANC & ";"

         Select Case Err.Number
            Case 0

            Case 2501

               Err.Clear

            Case Else

               GoTo Err_ImpVisu

         End Select

         On Error GoTo Err_ImpVisu

         Do While DonneEtatRaport("EtaCompteRenduSaisie", IntRaportEtat) = True

            DoEvents

         Loop

         On Error Resume Next

         DoCmd.OpenReport "EtaCompteRenduSaisie", IntVisu, , "Déb_MonnaieEnCours=" & cEURO & IIf(StrSQLWhere = "", "", " AND ") & StrSQLWhere, acWindowNormal, cEURO & ";"

         Select Case Err.Number
            Case 0

            Case 2501

               Err.Clear

            Case Else

               GoTo Err_ImpVisu

         End Select

         On Error GoTo Err_ImpVisu

         Do While DonneEtatRaport("EtaCompteRenduSaisie", IntRaportEtat) = True

            DoEvents

         Loop

      Case 5

         IntRepMsg = MsgBox("Souhaitez-vous exporter les données au format Excel ?" & vbCrLf & "Indiquez OUI pour extraire les données au format excel et NON pour le format texte.", vbYesNo + vbQuestion, "Format d'exportation")

         Select Case IntRepMsg
            Case vbYes

               DoCmd.OutputTo acOutputReport, "EtaCompteRenduSaisie", acFormatXLS, , False

            Case vbNo

               DoCmd.OutputTo acOutputReport, "EtaCompteRenduSaisie", acFormatRTF, , False

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
