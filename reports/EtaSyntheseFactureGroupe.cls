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

         TxtFac_Montant_mTTC.ControlSource = "=Round(Round(nz([Fac_Montant_Direct_m1],0),2)*[Fac_Honoraire]/100*(1+[Fac_TVA]),2)+Round(Round(nz([Fac_Montant_Encaisser_m1],0),2)*[Fac_Honoraire]/100*(1+[Fac_TVA]),2)"

         TxtFac_Montant_mTTC.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtCG_NomFac_Montant_mTTC.ControlSource = "=Sum(Round(Round(nz([Fac_Montant_Direct_m1],0),2)*[Fac_Honoraire]/100*(1+[Fac_TVA]),2))+Sum(Round(Round(nz([Fac_Montant_Encaisser_m1],0),2)*[Fac_Honoraire]/100*(1+[Fac_TVA]),2))"

         TxtCG_NomFac_Montant_mTTC.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtCG_Nom1Fac_Montant_mTTC.ControlSource = "=Sum(Round(Round(nz([Fac_Montant_Direct_m1],0),2)*[Fac_Honoraire]/100*(1+[Fac_TVA]),2))+Sum(Round(Round(nz([Fac_Montant_Encaisser_m1],0),2)*[Fac_Honoraire]/100*(1+[Fac_TVA]),2))"

         TxtCG_Nom1Fac_Montant_mTTC.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtFac_Montant_mTVA.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtFac_Montant_mHT.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """" & ""

         TxtCG_NomFac_Montant_mTVA.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """" & ""

         TxtCG_NomFac_Montant_mHT.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """" & ""

         TxtCG_Nom1Fac_Montant_mTVA.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """" & ""

         TxtCG_Nom1Fac_Montant_mHT.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """" & ""

      Case cEURO

         TxtFac_Montant_mTTC.ControlSource = "=Round(Round(nz([Fac_Montant_Direct_m2],0),2)*[Fac_Honoraire]/100*(1+[Fac_TVA]),2)+Round(Round(nz([Fac_Montant_Encaisser_m2],0),2)*[Fac_Honoraire]/100*(1+[Fac_TVA]),2)"

         TxtFac_Montant_mTTC.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtCG_NomFac_Montant_mTTC.ControlSource = "=Sum(Round(Round(nz([Fac_Montant_Direct_m2],0),2)*[Fac_Honoraire]/100*(1+[Fac_TVA]),2))+Sum(Round(Round(nz([Fac_Montant_Encaisser_m2],0),2)*[Fac_Honoraire]/100*(1+[Fac_TVA]),2))"

         TxtCG_NomFac_Montant_mTTC.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtCG_Nom1Fac_Montant_mTTC.ControlSource = "=Sum(Round(Round(nz([Fac_Montant_Direct_m2],0),2)*[Fac_Honoraire]/100*(1+[Fac_TVA]),2))+Sum(Round(Round(nz([Fac_Montant_Encaisser_m2],0),2)*[Fac_Honoraire]/100*(1+[Fac_TVA]),2))"

         TxtCG_Nom1Fac_Montant_mTTC.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtFac_Montant_mTVA.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtFac_Montant_mHT.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """" & ""

         TxtCG_NomFac_Montant_mTVA.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """" & ""

         TxtCG_NomFac_Montant_mHT.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """" & ""

         TxtCG_Nom1Fac_Montant_mTVA.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """" & ""

         TxtCG_Nom1Fac_Montant_mHT.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """" & ""

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

'   MsgBox "Aucune facture en " & DLookup("Mon_Nom", "MONNAIE", "Mon_Id=" & iMonnaieVisu) & " selectionnée !!", vbCritical

Exit_Report_NoData:

   Exit Sub

Err_Report_NoData:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Report_NoData"

   Resume Exit_Report_NoData
End Sub
