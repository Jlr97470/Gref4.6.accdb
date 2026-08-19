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

         TxtMontant_m.ControlSource = "=Round([Déb_Mont_Prin_m1],2)-Round(nz(DSum(" & Chr(34) & "Enc_Prin_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [TIrr_Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtMontant_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalMontant_m.ControlSource = "=Sum(Round([Déb_Mont_Prin_m1],2)-Round(nz(DSum(" & Chr(34) & "Enc_Prin_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [TIrr_Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2))"

         TxtTotalMontant_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtMontant_mMonnaie.ControlSource = "=TxtMontant_m*DonnerTauxConvertion(1)"

         TxtMontant_mMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalMontant_mMonnaie.ControlSource = "=TxtTotalMontant_m*DonnerTauxConvertion(1)"

         TxtTotalMontant_mMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

      Case cEURO

         TxtMontant_m.ControlSource = "=Round([Déb_Mont_Prin_m2],2)-Round(nz(DSum(" & Chr(34) & "Enc_Prin_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [TIrr_Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtMontant_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalMontant_m.ControlSource = "=Sum(Round([Déb_Mont_Prin_m2],2)-Round(nz(DSum(" & Chr(34) & "Enc_Prin_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [TIrr_Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2))"

         TxtTotalMontant_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtMontant_mMonnaie.ControlSource = "=TxtMontant_m*DonnerTauxConvertion(2)"

         TxtMontant_mMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalMontant_mMonnaie.ControlSource = "=TxtTotalMontant_m*DonnerTauxConvertion(2)"

         TxtTotalMontant_mMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtMontant_mMonnaie.Visible = False

         TxtTotalMontant_mMonnaie.Visible = False

         EtiTxtTotalMontant_mMonnaie.Visible = False

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

         Select Case StrParam(0)
            Case "Caen"

               EtiEnteteCaen.Visible = True

            Case "Paris"

               EtiEnteteCaen.Visible = False

         End Select

         Inverseur Nz(StrParam(1), cEURO)

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

  ' MsgBox "Aucun débiteur irrécouvrable en " & DLookup("Mon_Nom", "MONNAIE", "Mon_Id=" & iMonnaieVisu) & " selectionner !!", vbCritical

Exit_Report_NoData:

   Exit Sub

Err_Report_NoData:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Report_NoData"

   Resume Exit_Report_NoData
End Sub
