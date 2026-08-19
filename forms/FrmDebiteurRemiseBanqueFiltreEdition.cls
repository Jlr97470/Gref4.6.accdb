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
   Dim BolWhere As Boolean
   Dim StrDateDebut As String
   Dim StrDateFin As String
   Dim IntRaportEtat As Integer
   Dim IntRepMsg As Integer
   Dim IntNum As Integer
   Dim VarLigne As Variant
   Dim QdfRequete As DAO.QueryDef

   On Error GoTo Err_ImpVisu

   StrDateDebut = Nz(TxtDateDebut)

   StrDateFin = Nz(TxtDateFin)

   Select Case StrDateDebut
      Case ""

      Case Else

         Select Case CadEtat
            Case 1

               StrSQLWhere = "Déb_Date_Représentation>=" & ConvDate(StrDateDebut)

            Case 2

               StrSQLWhere = "Enc_Date>=" & ConvDate(StrDateDebut)

         End Select

         BolWhere = True

   End Select

   Select Case StrDateFin
      Case ""

      Case Else

         Select Case BolWhere
            Case True

               Select Case CadEtat
                  Case 1

                     StrSQLWhere = StrSQLWhere & " AND Déb_Date_Représentation<=" & ConvDate(StrDateFin)

                  Case 2

                     StrSQLWhere = StrSQLWhere & " AND Enc_Date<=" & ConvDate(StrDateFin)

               End Select

            Case False

               Select Case CadEtat
                  Case 1

                     StrSQLWhere = "Déb_Date_Représentation<=" & ConvDate(StrDateFin)

                  Case 2

                     StrSQLWhere = "Enc_Date<=" & ConvDate(StrDateFin)

               End Select

         End Select

         BolWhere = True

   End Select

   Select Case BolWhere
      Case True

         Select Case LstTypeEnc.ItemsSelected.Count
            Case Is > 0

               Select Case CadEtat
                  Case 1

                     StrSQLWhere = StrSQLWhere & " AND Déb_Repré_Type In("

                  Case 2

                     StrSQLWhere = StrSQLWhere & " AND Enc_Type In("

               End Select

            Case Else

         End Select

      Case False

         Select Case LstTypeEnc.ItemsSelected.Count
            Case Is > 0

               Select Case CadEtat
                  Case 1

                     StrSQLWhere = "Déb_Repré_Type In("

                  Case 2

                     StrSQLWhere = "Enc_Type In("

               End Select

            Case Else

         End Select

   End Select

   For Each VarLigne In LstTypeEnc.ItemsSelected

      StrSQLWhere = StrSQLWhere & "'" & LstTypeEnc.ItemData(VarLigne) & "',"

   Next

   Select Case LstTypeEnc.ItemsSelected.Count
      Case Is > 0

         StrSQLWhere = Left(StrSQLWhere, Len(StrSQLWhere) - 1) & ")"

      Case Else

   End Select

   Select Case CadEtat
      Case 1

         StrNomEtat = "EtaListeRepresentation"

         For Each QdfRequete In CurrentDb.QueryDefs

            Select Case QdfRequete.Name
               Case "SelListeRepresentation"

                  CurrentDb.QueryDefs.Delete QdfRequete.Name

               Case Else

            End Select

         Next

         StrSQL = "SELECT DEBITEUR.Déb_Code, DEBITEUR.Déb_Nom, DEBITEUR.Déb_Date_Représentation, DEBITEUR.Déb_Repré_Type, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb_Mont_Prin_m2, DEBITEUR.Déb_Mont_Prin_mo, ENCAISSEMENT_TYPE.ETy_Libellé " & _
                       "FROM DEBITEUR INNER JOIN ENCAISSEMENT_TYPE ON DEBITEUR.Déb_Repré_Type=ENCAISSEMENT_TYPE.ETY_Code " & IIf(StrSQLWhere = "", "", " WHERE ") & StrSQLWhere & ";"

         Set QdfRequete = CurrentDb.CreateQueryDef("SelListeRepresentation", StrSQL)

      Case 2

         StrNomEtat = "EtaListeEncaissement"

         For Each QdfRequete In CurrentDb.QueryDefs

            Select Case QdfRequete.Name
               Case "SelListeEncaissement"

                  CurrentDb.QueryDefs.Delete QdfRequete.Name

               Case Else

            End Select

         Next

         StrSQL = "SELECT ENCAISSEMENT.Enc_Id, ENCAISSEMENT.Enc_Type, ENCAISSEMENT.Enc_Date, ENCAISSEMENT.Enc_Montant_m1, ENCAISSEMENT.Enc_Montant_m2, ENCAISSEMENT.Enc_Montant_mo, DEBITEUR.Déb_Code, DEBITEUR.Déb_Nom, ENCAISSEMENT_TYPE.ETy_Libellé " & _
                       "FROM DEBITEUR INNER JOIN (ENCAISSEMENT INNER JOIN ENCAISSEMENT_TYPE ON ENCAISSEMENT.Enc_Type=ENCAISSEMENT_TYPE.ETy_Code) ON DEBITEUR.Déb_Code=ENCAISSEMENT.Enc_Code_Débiteur " & IIf(StrSQLWhere = "", "", " WHERE ") & StrSQLWhere & ";"

         Set QdfRequete = CurrentDb.CreateQueryDef("SelListeEncaissement", StrSQL)

   End Select

   Select Case IntVisu
      Case 0, 2

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

   MsgBox Err.Number & " " & Err.description

   Resume Exit_CmdFermer_Click
End Sub
