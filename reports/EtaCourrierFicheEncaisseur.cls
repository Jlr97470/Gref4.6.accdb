Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private iMonnaieVisu As Integer

Private Function Inverseur(ByVal IntMonnaie As Integer)

   On Error GoTo Err_Inverseur

   Select Case IntMonnaie
      Case cFRANC

         TxtDue.ControlSource = "=Round(Déb_Mont_Prin_m1,2)-Round(nz(DSum(" & Chr(34) & "Enc_Prin_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtDue.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtFraisTransactionnel.ControlSource = "=DLookUp(" & Chr(34) & "Frt_Frais_m1" & Chr(34) & "," & Chr(34) & "FRAIS_TRANS" & Chr(34) & "," & Chr(34) & "Frt_Grille=" & Chr(34) & " & Cli_Grille_FTrans & " & Chr(34) & " AND ((Frt_Début_m1<=CCur('" & Chr(34) & " & TxtDue & " & Chr(34) & "') AND Frt_Fin_m1>=CCur('" & Chr(34) & " & TxtDue & " & Chr(34) & "')) OR (Frt_Fin_m1=0 AND Frt_Début_m1<=CCur('" & Chr(34) & " & TxtDue & " & Chr(34) & "')))" & Chr(34) & ")"

         TxtFraisTransactionnel.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtFraisDeplacement.ControlSource = "=DLookUp('Frais_Déplacement_m1','PARAMETRE_GREF')"

         TxtFraisDeplacement.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotal.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtDéb_Mont_Prin_m.ControlSource = "=iif(IsNull([Enc_Code_Débiteur]),Null,[Déb_Mont_Prin_m1])"

         TxtDéb_Mont_Prin_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtEnc_Montant_m.ControlSource = "Enc_Montant_m1"

         TxtEnc_Montant_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtDueMonnaie.ControlSource = "=TxtDue * DonnerTauxConvertion(1)"

         TxtDueMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtFraisTransactionnelMonnaie.ControlSource = "=TxtFraisTransactionnel * DonnerTauxConvertion(1)"

         TxtFraisTransactionnelMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtFraisDeplacementMonnaie.ControlSource = "=TxtFraisDeplacement * DonnerTauxConvertion(1)"

         TxtFraisDeplacementMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalMonnaie.ControlSource = "=TxtTotal * DonnerTauxConvertion(1)"

         TxtTotalMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

      Case cEURO

         TxtDue.ControlSource = "=Round(Déb_Mont_Prin_m2,2)-Round(nz(DSum(" & Chr(34) & "Enc_Prin_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtDue.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtFraisTransactionnel.ControlSource = "=DLookUp(" & Chr(34) & "Frt_Frais_m2" & Chr(34) & "," & Chr(34) & "FRAIS_TRANS" & Chr(34) & "," & Chr(34) & "Frt_Grille=" & Chr(34) & " & Cli_Grille_FTrans & " & Chr(34) & " AND ((Frt_Début_m2<=CCur('" & Chr(34) & " & TxtDue & " & Chr(34) & "') AND Frt_Fin_m2>=CCur('" & Chr(34) & " & TxtDue & " & Chr(34) & "')) OR (Frt_Fin_m2=0 AND Frt_Début_m2<=CCur('" & Chr(34) & " & TxtDue & " & Chr(34) & "')))" & Chr(34) & ")"

         TxtFraisTransactionnel.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtFraisDeplacement.ControlSource = "=DLookUp('Frais_Déplacement_m2','PARAMETRE_GREF')"

         TxtFraisDeplacement.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotal.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtDéb_Mont_Prin_m.ControlSource = "=iif(IsNull([Enc_Code_Débiteur]),Null,[Déb_Mont_Prin_m2])"

         TxtDéb_Mont_Prin_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtEnc_Montant_m.ControlSource = "Enc_Montant_m2"

         TxtEnc_Montant_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtDueMonnaie.ControlSource = "=TxtDue * DonnerTauxConvertion(2)"

         TxtDueMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtFraisTransactionnelMonnaie.ControlSource = "=TxtFraisTransactionnel * DonnerTauxConvertion(2)"

         TxtFraisTransactionnelMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtFraisDeplacementMonnaie.ControlSource = "=TxtFraisDeplacement * DonnerTauxConvertion(2)"

         TxtFraisDeplacementMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalMonnaie.ControlSource = "=TxtTotal * DonnerTauxConvertion(2)"

         TxtTotalMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtDueMonnaie.Visible = False

         TxtFraisTransactionnelMonnaie.Visible = False

         TxtFraisDeplacementMonnaie.Visible = False

         TxtTotalMonnaie.Visible = False

   End Select

   iMonnaieVisu = IntMonnaie

Exit_Inverseur:

   Exit Function

Err_Inverseur:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Inverseur"

   Resume Exit_Inverseur
End Function

Private Sub Report_Open(Cancel As Integer)
   Dim StrParam() As String

   On Error GoTo Err_Report_Open

   StrParam = Split(Nz(Me.OpenArgs), ";")

   Select Case UBound(StrParam)
      Case -1

         Inverseur cEURO

      Case Else

         Inverseur Nz(StrParam(0), cEURO)

   End Select

Exit_Report_Open:

   Exit Sub

Err_Report_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Report_Open"

   Resume Exit_Report_Open
End Sub

Private Sub Report_NoData(Cancel As Integer)

   On Error GoTo Err_Report_NoData

   Cancel = True

 '  MsgBox "Aucune fiche encaisseur en " & DLookup("Mon_Nom", "MONNAIE", "Mon_Id=" & iMonnaieVisu) & " selectionnée !!", vbCritical

Exit_Report_NoData:

   Exit Sub

Err_Report_NoData:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Report_NoData"

   Resume Exit_Report_NoData
End Sub
