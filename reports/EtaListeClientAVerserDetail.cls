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

         TxtEnc_Prin_m.ControlSource = "=Enc_Prin_m1"

         TxtEnc_Prin_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtEnc_Prin_mMonnaie.ControlSource = "=Enc_Prin_m2"

         TxtEnc_Prin_mMonnaie.Format = "(#,##0.00""" & " €)" & """;-(#,##0.00""" & " €)" & """"

         TxtTotalEnc_Prin_m.ControlSource = "=Sum(Enc_Prin_m1)"

         TxtTotalEnc_Prin_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalEnc_Prin_mMonnaie.ControlSource = "=[TxtTotalEnc_Prin_m]*DonnerTauxConvertion(1)"

         TxtTotalEnc_Prin_mMonnaie.Format = "(#,##0.00""" & " €)" & """;(-#,##0.00""" & " €)" & """"

         TxtTotal.ControlSource = "=Sum([Enc_Prin_m1])"

         TxtTotal.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalMonnaie.ControlSource = "=[TxtTotal]*DonnerTauxConvertion(1)"

         TxtTotalMonnaie.Format = "(#,##0.00""" & " €)" & """;(-#,##0.00""" & " €)" & """"

      Case cEURO

         TxtEnc_Prin_m.ControlSource = "=Enc_Prin_m2"

         TxtEnc_Prin_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtEnc_Prin_mMonnaie.ControlSource = "=Enc_Prin_m1"

         TxtEnc_Prin_mMonnaie.Format = "(#,##0.00""" & " F)" & """;-(#,##0.00""" & " F)" & """"

         TxtTotalEnc_Prin_m.ControlSource = "=Sum(Enc_Prin_m2)"

         TxtTotalEnc_Prin_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalEnc_Prin_mMonnaie.ControlSource = "=[TxtTotalEnc_Prin_m]*DonnerTauxConvertion(2)"

         TxtTotalEnc_Prin_mMonnaie.Format = "(#,##0.00""" & " F)" & """;(-#,##0.00""" & " F)" & """"

         TxtTotal.ControlSource = "=Sum([Enc_Prin_m2])"

         TxtTotal.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalMonnaie.ControlSource = "=[TxtTotal]*DonnerTauxConvertion(2)"

         TxtTotalMonnaie.Format = "(#,##0.00""" & " F)" & """;(-#,##0.00""" & " F)" & """"

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

'   MsgBox "Aucun débiteur selectionnée en " & DLookup("Mon_Nom", "MONNAIE", "Mon_Id=" & iMonnaieVisu) & "!!", vbCritical

Exit_Report_NoData:

   Exit Sub

Err_Report_NoData:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Report_NoData"

   Resume Exit_Report_NoData
End Sub
