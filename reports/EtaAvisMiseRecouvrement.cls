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

         TxtDéb_Mont_Prin_m.ControlSource = "=Round([Déb_Mont_Prin_m1],2)-Round(nz(DSum(" & Chr(34) & "Enc_Prin_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtDéb_Mont_Prin_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtDéb_Frais_m.ControlSource = "=Round([Déb_Frais_m1],2)-Round(nz(DSum(" & Chr(34) & "Enc_Frais_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtDéb_Frais_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotal.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtDéb_Mont_Prin_mMonnaie.ControlSource = "=TxtDéb_Mont_Prin_m*DonnerTauxConvertion(1)"

         TxtDéb_Mont_Prin_mMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtDéb_Frais_mMonnaie.ControlSource = "=TxtDéb_Frais_m*DonnerTauxConvertion(1)"

         TxtDéb_Frais_mMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalMonnaie.ControlSource = "=TxtTotal*DonnerTauxConvertion(1)"

         TxtTotalMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

      Case cEURO

         TxtDéb_Mont_Prin_m.ControlSource = "=Round([Déb_Mont_Prin_m2],2)-Round(nz(DSum(" & Chr(34) & "Enc_Prin_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtDéb_Mont_Prin_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtDéb_Frais_m.ControlSource = "=Round([Déb_Frais_m2],2)-Round(nz(DSum(" & Chr(34) & "Enc_Frais_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtDéb_Frais_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotal.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtDéb_Mont_Prin_mMonnaie.ControlSource = "=TxtDéb_Mont_Prin_m*DonnerTauxConvertion(2)"

         TxtDéb_Mont_Prin_mMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtDéb_Frais_mMonnaie.ControlSource = "=TxtDéb_Frais_m*DonnerTauxConvertion(2)"

         TxtDéb_Frais_mMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalMonnaie.ControlSource = "=TxtTotal*DonnerTauxConvertion(2)"

         TxtTotalMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtDéb_Mont_Prin_mMonnaie.Visible = False

         TxtDéb_Frais_mMonnaie.Visible = False

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

   'MsgBox "Aucun débiteur avec des avis de mise en recouvrement en " & DLookup("Mon_Nom", "MONNAIE", "Mon_Id=" & iMonnaieVisu) & "!!", vbCritical

Exit_Report_NoData:

   Exit Sub

Err_Report_NoData:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Report_NoData"

   Resume Exit_Report_NoData
End Sub
