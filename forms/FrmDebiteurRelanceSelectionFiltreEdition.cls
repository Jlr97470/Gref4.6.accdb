Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Function ImpVisu(ByVal IntVisu As Integer)
   Dim StrNomEtat As String
   Dim StrSQL As String
   Dim StrSQLWhere As String
   Dim IntRaportEtat As Integer

   On Error GoTo Err_ImpVisu

   Select Case CadChoix
      Case 1

         StrSQLWhere = "TRel_Editer = True AND TRel_Niv_relance=1"

         StrNomEtat = "EtaAvisMiseRecouvrement"

      Case 2

         StrSQLWhere = "TRel_Editer = True AND TRel_Niv_relance In (2,4,6)"

         StrNomEtat = "EtaPreavisAssignation"

      Case 3

         StrSQLWhere = "TRel_Editer = True AND TRel_Niv_relance In (3,5)"

         StrNomEtat = "EtaAvisAssignation"

      Case 4

         StrSQLWhere = "De_Date_éch<=Date()"

         StrNomEtat = "EtaListeDebiteurEcheance"

   End Select

   DoCmd.OpenReport StrNomEtat, acDesign

   Reports(StrNomEtat).Filter = StrSQLWhere

   Reports(StrNomEtat).FilterOn = True

   DoCmd.Close acReport, StrNomEtat, acSavePrompt

   Select Case CadChoix
      Case 1, 2, 3

         On Error Resume Next

         DoCmd.OpenReport StrNomEtat, IntVisu, , "Déb_MonnaieEnCours=" & cFRANC & IIf(StrSQLWhere = "", "", " AND ") & StrSQLWhere, , DLookup("Site", "PARAMETRE_GREF") & ";" & cFRANC

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

         DoCmd.OpenReport StrNomEtat, IntVisu, , "Déb_MonnaieEnCours=" & cEURO & IIf(StrSQLWhere = "", "", " AND ") & StrSQLWhere, , DLookup("Site", "PARAMETRE_GREF") & ";" & cEURO

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

               StrSQL = "UPDATE TMP_RELANCE SET TRel_Deja_Editer = True  WHERE " & StrSQLWhere & ";"

               CurrentDb.Execute StrSQL, dbFailOnError

            Case 2

         End Select

      Case 4

         On Error Resume Next

         DoCmd.OpenReport StrNomEtat, IntVisu, , StrSQLWhere, acDialog

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
