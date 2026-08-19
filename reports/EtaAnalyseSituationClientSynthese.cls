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

         TxtSolde.ControlSource = "=[Déb_Mont_Prin_m1]-[Déb_Prin_Enc_m1]"

         TxtSolde.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtSoldeGroupe.ControlSource = "=Sum([Déb_Mont_Prin_m1]-[Déb_Prin_Enc_m1])"

         TxtSoldeGroupe.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalSolde.ControlSource = "=Sum([Déb_Mont_Prin_m1]-[Déb_Prin_Enc_m1])"

         TxtTotalDéb_Mont_Reversé_m.ControlSource = "=Sum([Déb_Mont_Reversé_m1])+Round(Sum([Déb_Mont_Reversé_m2])*DonnerTauxConvertion(2),2)"

         TxtTotalDéb_Prin_Enc_m.ControlSource = "=Sum([Déb_Prin_Enc_m1])"

         TxtTotalReversé.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalResteAVerser.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtSoldeGroupeMonnaie.ControlSource = "=[TxtSoldeGroupe]*DonnerTauxConvertion(1)"

         TxtSoldeGroupeMonnaie.Format = "(#,##0.00""" & " €)" & """;(-#,##0.00""" & " €)" & """"

         TxtTotalReverséMonnaie.ControlSource = "=[TxtTotalDéb_Mont_Reversé_m]*DonnerTauxConvertion(1)"

         TxtTotalReverséMonnaie.Format = "(#,##0.00""" & " €)" & """;(-#,##0.00""" & " €)" & """"

         TxtTotalResteAVerserMonnaie.ControlSource = "=([TxtTotalDéb_Prin_Enc_m] - [TxtTotalDéb_Mont_Reversé_m]) * DonnerTauxConvertion(1)"

         TxtTotalResteAVerserMonnaie.Format = "(#,##0.00""" & " €)" & """;(-#,##0.00""" & " €)" & """"

      Case cEURO

         TxtSolde.ControlSource = "=[Déb_Mont_Prin_m2]-[Déb_Prin_Enc_m2]"

         TxtSolde.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtSoldeGroupe.ControlSource = "=Sum([Déb_Mont_Prin_m2]-[Déb_Prin_Enc_m2])"

         TxtSoldeGroupe.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalSolde.ControlSource = "=Sum([Déb_Mont_Prin_m2]-[Déb_Prin_Enc_m2])"

         TxtTotalDéb_Mont_Reversé_m.ControlSource = "=Sum([Déb_Mont_Reversé_m2])+Round(Sum([Déb_Mont_Reversé_m1])*DonnerTauxConvertion(1),2)"

         TxtTotalDéb_Prin_Enc_m.ControlSource = "=Sum([Déb_Prin_Enc_m2])"

         TxtTotalReversé.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalResteAVerser.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtSoldeGroupeMonnaie.ControlSource = "=[TxtSoldeGroupe]*DonnerTauxConvertion(2)"

         TxtSoldeGroupeMonnaie.Format = "(#,##0.00""" & " F)" & """;(-#,##0.00""" & " F)" & """"

         TxtTotalReverséMonnaie.ControlSource = "=[TxtTotalDéb_Mont_Reversé_m]*DonnerTauxConvertion(2)"

         TxtTotalReverséMonnaie.Format = "(#,##0.00""" & " F)" & """;(-#,##0.00""" & " F)" & """"

         TxtTotalResteAVerserMonnaie.ControlSource = "=([TxtTotalDéb_Prin_Enc_m] - [TxtTotalDéb_Mont_Reversé_m]) * DonnerTauxConvertion(2)"

         TxtTotalResteAVerserMonnaie.Format = "(#,##0.00""" & " F)" & """;(-#,##0.00""" & " F)" & """"

         TxtSoldeGroupeMonnaie.Visible = False

         TxtTotalReverséMonnaie.Visible = False

         TxtTotalResteAVerserMonnaie.Visible = False

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
