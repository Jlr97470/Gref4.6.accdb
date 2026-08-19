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

         TxtPrincipal.ControlSource = "=iif([CG_Code_Action] In (6,7),0,[Déb_Mont_Prin_m1]-[Déb_Mont_reversé_m1]-[Déb_Mont_reversé_m2]*DonnerTauxConvertion(2))"

         TxtPrincipal.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtRèglement.ControlSource = "=iif([CG_Code_Action]=7,0,[Déb_Mont_Reversé_m1]+[Déb_Mont_Reversé_m2]*DonnerTauxConvertion(2))"

         TxtRèglement.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtSolde.ControlSource = "=iif([CG_Code_Action] In (6,7),0,[Déb_Mont_Prin_m1]-[Déb_Mont_Reversé_m1]-[Déb_Mont_Reversé_m2]*DonnerTauxConvertion(2))"

         TxtSolde.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotal.ControlSource = "=Sum(iif([CG_Code_Action] In (1,2,3,4,5,9,10),[Déb_Mont_Prin_m1]-[Déb_Enc_Prin_m1],0))"

         TxtTotal.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalMonnaie.ControlSource = "=[TxtTotal]*DonnerTauxConvertion(1)"

         TxtTotalMonnaie.Format = "(#,##0.00""" & " €)" & """;(-#,##0.00""" & " €)" & """"

      Case cEURO

         TxtPrincipal.ControlSource = "=iif([CG_Code_Action] In (6,7),0,[Déb_Mont_Prin_m2]-[Déb_Mont_reversé_m2]-[Déb_Mont_reversé_m1]*DonnerTauxConvertion(1))"

         TxtPrincipal.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtRèglement.ControlSource = "=iif([CG_Code_Action]=7,0,[Déb_Mont_Reversé_m2]+[Déb_Mont_Reversé_m1]*DonnerTauxConvertion(1))"

         TxtRèglement.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtSolde.ControlSource = "=iif([CG_Code_Action] In (6,7),0,[Déb_Mont_Prin_m2]-[Déb_Mont_Reversé_m2]-[Déb_Mont_Reversé_m1]*DonnerTauxConvertion(1))"

         TxtSolde.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotal.ControlSource = "=Sum(iif([CG_Code_Action] In (1,2,3,4,5,9,10),[Déb_Mont_Prin_m2]-[Déb_Enc_Prin_m2],0))"

         TxtTotal.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalMonnaie.ControlSource = "=[TxtTotal]*DonnerTauxConvertion(2)"

         TxtTotalMonnaie.Format = "(#,##0.00""" & " F)" & """;(-#,##0.00""" & " F)" & """"

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

   'MsgBox "Aucun débiteur selectionnée en " & DLookup("Mon_Nom", "MONNAIE", "Mon_Id=" & iMonnaieVisu) & "!!", vbCritical

Exit_Report_NoData:

   Exit Sub

Err_Report_NoData:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Report_NoData"

   Resume Exit_Report_NoData
End Sub
