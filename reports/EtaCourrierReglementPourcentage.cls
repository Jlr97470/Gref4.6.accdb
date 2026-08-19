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

         TxtRegD_Montant_m.ControlSource = "RegD_Montant_m1"

         TxtRegD_Montant_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtResteAVerser.ControlSource = "=IIf([RegD_Ety_Mode]=1,Round([Déb_Mont_Prin_m1],2)-Round(Nz(DSum(" & Chr(34) & "Enc_Prin_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_RegD_Id<=" & Chr(34) & " & [RegD_Id] & " & Chr(34) & " AND Enc_Code_Débiteur='" & Chr(34) & " & [RegD_Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2),Null)"

         TxtResteAVerser.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalRegD_Montant_m.ControlSource = "=Nz(Sum(Round([RegD_Montant_m1],2)),0)"

         TxtTotalRegD_Montant_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalRegD_Montant_mEncaisser.ControlSource = "=Nz(Sum(iif([RegD_ETy_Mode]=1,Round([RegD_Montant_m1],2),0)),0)"

         TxtTotalRegD_Montant_mEncaisser.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalRegD_Montant_mDirect.ControlSource = "=Nz(Sum(iif([RegD_ETy_Mode]=2,Round([RegD_Montant_m1],2),0)),0)"

         TxtTotalRegD_Montant_mDirect.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalRegD_Montant_mEncaisserTTC.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalRegD_Montant_mDirectTTC.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalRegD_Montant_mTTC.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalRegD_Montant_mTVA.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtMontantReverser.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtRegD_Montant_mMonnaie.ControlSource = "=[TxtRegD_Montant_m]*DonnerTauxConvertion(1)"

         TxtRegD_Montant_mMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtResteAVerserMonnaie.ControlSource = "=iif([RegD_ETy_Mode]=1,[TxtResteAVerser]*DonnerTauxConvertion(1),Null)"

         TxtResteAVerserMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalRegD_Montant_mMonnaie.ControlSource = "=[TxtTotalRegD_Montant_m]*DonnerTauxConvertion(1)"

         TxtTotalRegD_Montant_mMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalRegD_Montant_mEncaisserMonnaie.ControlSource = "=[TxtTotalRegD_Montant_mEncaisser]*DonnerTauxConvertion(1)"

         TxtTotalRegD_Montant_mEncaisserMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalRegD_Montant_mDirectMonnaie.ControlSource = "=[TxtTotalRegD_Montant_mDirect]*DonnerTauxConvertion(1)"

         TxtTotalRegD_Montant_mDirectMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtMontantReverserMonnaie.ControlSource = "=[TxtMontantReverser]*DonnerTauxConvertion(1)"

         TxtMontantReverserMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalRegD_Montant_mEncaisserTTCMonnaie.ControlSource = "=[TxtTotalRegD_Montant_mEncaisserTTC]*DonnerTauxConvertion(1)"

         TxtTotalRegD_Montant_mEncaisserTTCMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalRegD_Montant_mDirectTTCMonnaie.ControlSource = "=[TxtTotalRegD_Montant_mDirectTTC]*DonnerTauxConvertion(1)"

         TxtTotalRegD_Montant_mDirectTTCMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalRegD_Montant_mTTCMonnaie.ControlSource = "=[TxtTotalRegD_Montant_mTTC]*DonnerTauxConvertion(1)"

         TxtTotalRegD_Montant_mTTCMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalRegD_Montant_mTVAMonnaie.ControlSource = "=[TxtTotalRegD_Montant_mTVA]*DonnerTauxConvertion(1)"

         TxtTotalRegD_Montant_mTVAMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

      Case cEURO

         TxtRegD_Montant_m.ControlSource = "RegD_Montant_m2"

         TxtRegD_Montant_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtResteAVerser.ControlSource = "=IIf([RegD_Ety_Mode]=1,Round([Déb_Mont_Prin_m2],2)-Round(Nz(DSum(" & Chr(34) & "Enc_Prin_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_RegD_Id<=" & Chr(34) & " & [RegD_Id] & " & Chr(34) & " AND Enc_Code_Débiteur='" & Chr(34) & " & [RegD_Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2),Null)"

         TxtResteAVerser.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalRegD_Montant_m.ControlSource = "=Nz(Sum(Round([RegD_Montant_m2],2)),0)"

         TxtTotalRegD_Montant_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalRegD_Montant_mEncaisser.ControlSource = "=Nz(Sum(iif([RegD_ETy_Mode]=1,Round([RegD_Montant_m2],2),0)),0)"

         TxtTotalRegD_Montant_mEncaisser.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalRegD_Montant_mDirect.ControlSource = "=Nz(Sum(iif([RegD_ETy_Mode]=2,Round([RegD_Montant_m2],2),0)),0)"

         TxtTotalRegD_Montant_mDirect.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalRegD_Montant_mEncaisserTTC.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalRegD_Montant_mDirectTTC.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalRegD_Montant_mTTC.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalRegD_Montant_mTVA.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtMontantReverser.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtRegD_Montant_mMonnaie.ControlSource = "=[TxtRegD_Montant_m]*DonnerTauxConvertion(2)"

         TxtRegD_Montant_mMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtResteAVerserMonnaie.ControlSource = "=iif([RegD_ETy_Mode]=1,[TxtResteAVerser]*DonnerTauxConvertion(2),Null)"

         TxtResteAVerserMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalRegD_Montant_mMonnaie.ControlSource = "=[TxtTotalRegD_Montant_m]*DonnerTauxConvertion(2)"

         TxtTotalRegD_Montant_mMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalRegD_Montant_mEncaisserMonnaie.ControlSource = "=[TxtTotalRegD_Montant_mEncaisser]*DonnerTauxConvertion(2)"

         TxtTotalRegD_Montant_mEncaisserMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalRegD_Montant_mDirectMonnaie.ControlSource = "=[TxtTotalRegD_Montant_mDirect]*DonnerTauxConvertion(2)"

         TxtTotalRegD_Montant_mDirectMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtMontantReverserMonnaie.ControlSource = "=[TxtMontantReverser]*DonnerTauxConvertion(2)"

         TxtMontantReverserMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalRegD_Montant_mEncaisserTTCMonnaie.ControlSource = "=[TxtTotalRegD_Montant_mEncaisserTTC]*DonnerTauxConvertion(2)"

         TxtTotalRegD_Montant_mEncaisserTTCMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalRegD_Montant_mDirectTTCMonnaie.ControlSource = "=[TxtTotalRegD_Montant_mDirectTTC]*DonnerTauxConvertion(2)"

         TxtTotalRegD_Montant_mDirectTTCMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalRegD_Montant_mTTCMonnaie.ControlSource = "=[TxtTotalRegD_Montant_mTTC]*DonnerTauxConvertion(2)"

         TxtTotalRegD_Montant_mTTCMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalRegD_Montant_mTVAMonnaie.ControlSource = "=[TxtTotalRegD_Montant_mTVA]*DonnerTauxConvertion(2)"

         TxtTotalRegD_Montant_mTVAMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtRegD_Montant_mMonnaie.Visible = False

         TxtResteAVerserMonnaie.Visible = False

         TxtTotalRegD_Montant_mMonnaie.Visible = False

         TxtTotalRegD_Montant_mEncaisserMonnaie.Visible = False

         TxtTotalRegD_Montant_mDirectMonnaie.Visible = False

         TxtMontantReverserMonnaie.Visible = False

         TxtTotalRegD_Montant_mEncaisserTTCMonnaie.Visible = False

         TxtTotalRegD_Montant_mDirectTTCMonnaie.Visible = False

         TxtTotalRegD_Montant_mTTCMonnaie.Visible = False

         TxtTotalRegD_Montant_mTVAMonnaie.Visible = False

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

'   MsgBox "Aucun réglement au pourcentage en " & DLookup("Mon_Nom", "MONNAIE", "Mon_Id=" & iMonnaieVisu) & " selectionnée !!", vbCritical

Exit_Report_NoData:

   Exit Sub

Err_Report_NoData:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Report_NoData"

   Resume Exit_Report_NoData
End Sub
