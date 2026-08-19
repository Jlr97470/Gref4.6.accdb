Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private StrForm As String

Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   StrForm = Me.OpenArgs

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Load"

   Resume Exit_Form_Load
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

    Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

    Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdOK_Click()
   Dim BolWhere As Boolean
   Dim StrDéb_Code As String
   Dim StrDéb_Nom As String
   Dim StrDéb_Mont_Prin_m As String
   Dim StrDéb_Code_Enc As String

   On Error GoTo Err_CmdOK_Click

   StrDéb_Code = Nz(TxtDéb_Code)

   StrDéb_Nom = Nz(TxtDéb_Nom)

   StrDéb_Mont_Prin_m = Nz(TxtDéb_Mont_Prin_m)

   StrDéb_Code_Enc = Nz(TxtDéb_Code_Enc)

   BolWhere = False

   Forms(StrForm).SQLWhere = ""

   Select Case StrDéb_Code
      Case ""

      Case Else

            Select Case InStr(StrDéb_Code, "*")
               Case Is > 0

               Case Else

                 StrDéb_Code = Nz(StrDéb_Code) & "*"

            End Select

            Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Code] LIKE """ & StrDéb_Code & """"

            BolWhere = True

   End Select

   Select Case StrDéb_Nom
      Case ""

      Case Else

         Select Case InStr(StrDéb_Nom, "*")
            Case Is > 0

            Case Else

               StrDéb_Nom = Nz(StrDéb_Nom) & "*"

         End Select

         Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Nom] LIKE """ & StrDéb_Nom & """"

         BolWhere = True

   End Select

   Select Case StrDéb_Mont_Prin_m
      Case ""

      Case Else

         Select Case InStr(StrDéb_Mont_Prin_m, "*")
            Case Is > 0

            Case Else

               StrDéb_Mont_Prin_m = Nz(StrDéb_Mont_Prin_m) & "*"

         End Select

         Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & " (Round([Déb_Mont_Prin_m1],2) LIKE """ & StrDéb_Mont_Prin_m & """" & " OR Round([Déb_Mont_Prin_m2],2) LIKE """ & StrDéb_Mont_Prin_m & """" & ")"

         BolWhere = True

   End Select

   Select Case StrDéb_Code_Enc
      Case ""

      Case Else

            Select Case InStr(StrDéb_Code_Enc, "*")
               Case Is > 0

               Case Else

                 StrDéb_Code_Enc = Nz(StrDéb_Code_Enc) & "*"

            End Select

            Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Code_Enc] LIKE """ & StrDéb_Code_Enc & """"

            BolWhere = True

   End Select

   Select Case CadClass
      Case 0

        Forms(StrForm).SQLOrder = "Déb_Code"

      Case 1

        Forms(StrForm).SQLOrder = "Déb_Nom"

      Case 2

        Forms(StrForm).SQLOrder = "Déb_Mont_Prin_m1"

      Case 3

        Forms(StrForm).SQLOrder = "Déb_Code_Enc"

    End Select

    DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdOK_Click:

    Exit Sub

Err_CmdOK_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

    Resume Exit_CmdOK_Click
End Sub
