Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Open(Cancel As Integer)
   Dim IntMonnaie As Integer
   Dim IntTra As Integer
   Dim IntRel As Integer

   Dim StrParam() As String

   On Error GoTo Err_Report_Open

   StrParam = Split(Nz(Me.OpenArgs), ";")

   Select Case UBound(StrParam)
      Case -1

         IntMonnaie = cEURO

      Case Else

         IntMonnaie = Nz(StrParam(0), cEURO)

   End Select

   Select Case IntMonnaie
      Case cEURO

        For IntTra = 1 To 18

            Me("Fra_Tranche" & IntTra).ControlSource = "Fra_Tranche" & IntTra & "_m2"

            For IntRel = 1 To 7

                Me(CStr(IntTra) & CStr(IntRel)).Format = "0.00 E"

            Next IntRel

        Next IntTra

      Case cFRANC

        For IntTra = 1 To 18

            Me("Fra_Tranche" & IntTra).ControlSource = "Fra_Tranche" & IntTra & "_m1"

            For IntRel = 1 To 7

                Me(CStr(IntTra) & CStr(IntRel)).Format = "0.00 F"

            Next IntRel

        Next IntTra

    End Select

Exit_Report_Open:

   Exit Sub

Err_Report_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Report_Open"

   Resume Exit_Report_Open
End Sub
