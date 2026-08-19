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

         TxtFacD_Montant_m.ControlSource = "FacD_Montant_m1"

         TxtFacD_Montant_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtFacD_Montant_ho.ControlSource = "=Round([FacD_Montant_m1],2)*[Fac_Honoraire]/100"

         TxtFacD_Montant_ho.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalFacD_Montant_m.ControlSource = "=Nz(Sum(Round([FacD_Montant_m1],2)),0)"

         TxtTotalFacD_Montant_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalFacD_Montant_ho.ControlSource = "=Nz(Sum(Round([FacD_Montant_m1],2)*[Fac_Honoraire]/100),0)"

         TxtTotalFacD_Montant_ho.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalFacD_Montant_mEncaisser.ControlSource = "=Nz(Sum(iif([FacD_ETy_Mode]=1,Round([FacD_Montant_m1],2),0)),0)"

         TxtTotalFacD_Montant_mEncaisser.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalFacD_Montant_mDirect.ControlSource = "=Nz(Sum(iif([FacD_ETy_Mode]=2,Round([FacD_Montant_m1],2),0)),0)"

         TxtTotalFacD_Montant_mDirect.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalFacD_Montant_mEncaisserTTC.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalFacD_Montant_mDirectTTC.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalFacD_Montant_mTTC.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalFacD_Montant_mTVA.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtMontantReverser.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalFacD_Montant_mEncaisserMonnaie.ControlSource = "=[TxtTotalFacD_Montant_mEncaisser]*DonnerTauxConvertion(1)"

         TxtTotalFacD_Montant_mEncaisserMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalFacD_Montant_mDirectMonnaie.ControlSource = "=[TxtTotalFacD_Montant_mDirect]*DonnerTauxConvertion(1)"

         TxtTotalFacD_Montant_mDirectMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtMontantReverserMonnaie.ControlSource = "=[TxtMontantReverser]*DonnerTauxConvertion(1)"

         TxtMontantReverserMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalFacD_Montant_mEncaisserTTCMonnaie.ControlSource = "=[TxtTotalFacD_Montant_mEncaisserTTC]*DonnerTauxConvertion(1)"

         TxtTotalFacD_Montant_mEncaisserTTCMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalFacD_Montant_mDirectTTCMonnaie.ControlSource = "=[TxtTotalFacD_Montant_mDirectTTC]*DonnerTauxConvertion(1)"

         TxtTotalFacD_Montant_mDirectTTCMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalFacD_Montant_mTTCMonnaie.ControlSource = "=[TxtTotalFacD_Montant_mTTC]*DonnerTauxConvertion(1)"

         TxtTotalFacD_Montant_mTTCMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalFacD_Montant_mTVAMonnaie.ControlSource = "=[TxtTotalFacD_Montant_mTVA]*DonnerTauxConvertion(1)"

         TxtTotalFacD_Montant_mTVAMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

      Case cEURO

         TxtFacD_Montant_m.ControlSource = "FacD_Montant_m2"

         TxtFacD_Montant_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtFacD_Montant_ho.ControlSource = "=Round([FacD_Montant_m2],2)*[Fac_Honoraire]/100"

         TxtFacD_Montant_ho.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalFacD_Montant_m.ControlSource = "=Nz(Sum(Round([FacD_Montant_m2],2)),0)"

         TxtTotalFacD_Montant_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalFacD_Montant_ho.ControlSource = "=Nz(Sum(Round([FacD_Montant_m2],2)*[Fac_Honoraire]/100),0)"

         TxtTotalFacD_Montant_ho.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalFacD_Montant_mEncaisser.ControlSource = "=Nz(Sum(iif([FacD_ETy_Mode]=1,Round([FacD_Montant_m2],2),0)),0)"

         TxtTotalFacD_Montant_mEncaisser.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalFacD_Montant_mDirect.ControlSource = "=Nz(Sum(iif([FacD_ETy_Mode]=2,Round([FacD_Montant_m2],2),0)),0)"

         TxtTotalFacD_Montant_mDirect.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalFacD_Montant_mEncaisserTTC.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalFacD_Montant_mDirectTTC.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalFacD_Montant_mTTC.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalFacD_Montant_mTVA.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtMontantReverser.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalFacD_Montant_mEncaisserMonnaie.ControlSource = "=[TxtTotalFacD_Montant_mEncaisser]*DonnerTauxConvertion(2)"

         TxtTotalFacD_Montant_mEncaisserMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalFacD_Montant_mDirectMonnaie.ControlSource = "=[TxtTotalFacD_Montant_mDirect]*DonnerTauxConvertion(2)"

         TxtTotalFacD_Montant_mDirectMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtMontantReverserMonnaie.ControlSource = "=[TxtMontantReverser]*DonnerTauxConvertion(2)"

         TxtMontantReverserMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalFacD_Montant_mEncaisserTTCMonnaie.ControlSource = "=[TxtTotalFacD_Montant_mEncaisserTTC]*DonnerTauxConvertion(2)"

         TxtTotalFacD_Montant_mEncaisserTTCMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalFacD_Montant_mDirectTTCMonnaie.ControlSource = "=[TxtTotalFacD_Montant_mDirectTTC]*DonnerTauxConvertion(2)"

         TxtTotalFacD_Montant_mDirectTTCMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalFacD_Montant_mTTCMonnaie.ControlSource = "=[TxtTotalFacD_Montant_mTTC]*DonnerTauxConvertion(2)"

         TxtTotalFacD_Montant_mTTCMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalFacD_Montant_mTVAMonnaie.ControlSource = "=[TxtTotalFacD_Montant_mTVA]*DonnerTauxConvertion(2)"

         TxtTotalFacD_Montant_mTVAMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalFacD_Montant_mEncaisserMonnaie.Visible = False

         TxtTotalFacD_Montant_mDirectMonnaie.Visible = False

         TxtMontantReverserMonnaie.Visible = False

         TxtTotalFacD_Montant_mEncaisserTTCMonnaie.Visible = False

         TxtTotalFacD_Montant_mDirectTTCMonnaie.Visible = False

         TxtTotalFacD_Montant_mTTCMonnaie.Visible = False

         TxtTotalFacD_Montant_mTVAMonnaie.Visible = False

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

  ' MsgBox "Aucune facture en " & DLookup("Mon_Nom", "MONNAIE", "Mon_Id=" & iMonnaieVisu) & " selectionnée !!", vbCritical

Exit_Report_NoData:

   Exit Sub

Err_Report_NoData:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Report_NoData"

   Resume Exit_Report_NoData
End Sub
