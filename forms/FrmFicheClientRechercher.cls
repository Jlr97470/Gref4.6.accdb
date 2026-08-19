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

Private Sub CmbCG_Parent_Id_AfterUpdate()

   On Error GoTo Err_CmbCG_Parent_Id_AfterUpdate

   CmbCli_Groupe.Requery

   CmbCli_Groupe = Null

Exit_CmbCG_Parent_Id_AfterUpdate:

   Exit Sub

Err_CmbCG_Parent_Id_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbCG_Parent_Id_AfterUpdate:"

   Resume Exit_CmbCG_Parent_Id_AfterUpdate
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
   Dim StrCli_Code As String
   Dim StrCli_Nom As String
   Dim LngCG_Parent_Id As Long
   Dim LngCli_Groupe As Long
   Dim StrAC_CodePostal As String
   Dim StrAC_Ville As String

   On Error GoTo Err_CmdOK_Click

   StrCli_Code = Nz(TxtCli_Code)

   StrCli_Nom = Nz(TxtCli_Nom)

   LngCG_Parent_Id = Nz(CmbCG_Parent_Id)

   LngCli_Groupe = Nz(CmbCli_Groupe)

   StrAC_CodePostal = Nz(TxtAC_CodePostal)

   StrAC_Ville = Nz(TxtAC_Ville)

   Forms(StrForm).SQLWhere = ""

   BolWhere = False

   Select Case StrCli_Code
      Case ""

      Case Else

         Select Case InStr(StrCli_Code, "*")
            Case Is > 0

            Case Else

              StrCli_Code = StrCli_Code & "*"

         End Select

         Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & "[Cli_Code] LIKE """ & StrCli_Code & """"

         BolWhere = True

   End Select

   Select Case StrCli_Nom
      Case ""

      Case Else

         Select Case InStr(StrCli_Nom, "*")
           Case Is > 0

           Case Else

              StrCli_Nom = Nz(StrCli_Nom) & "*"

         End Select

         Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & "[Cli_Nom] LIKE """ & StrCli_Nom & """"

         BolWhere = True

   End Select

   Select Case LngCG_Parent_Id <> 0 And LngCli_Groupe = 0
      Case True

         Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & "[Cli_Groupe] In (SELECT CG_Id FROM CLIENT_GROUPE WHERE CG_Parent_Id=" & LngCG_Parent_Id & ";)"

         BolWhere = True

      Case False

   End Select

   Select Case LngCli_Groupe
      Case 0

      Case Else

         Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & "[Cli_Groupe]=" & LngCli_Groupe

         BolWhere = True

   End Select

   Select Case StrAC_CodePostal
      Case ""

      Case Else

         Select Case InStr(StrAC_CodePostal, "*")
            Case Is > 0

            Case Else

              StrAC_CodePostal = StrAC_CodePostal & "*"

         End Select

         Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & "[Cli_Code] In (SELECT AC_Code_Client FROM ADRESSE_CLIENT WHERE AC_CodePostal Like """ & StrAC_CodePostal & """;)"

         BolWhere = True

   End Select

   Select Case StrAC_Ville
      Case ""

      Case Else

         Select Case InStr(StrAC_Ville, "*")
            Case Is > 0

            Case Else

              StrAC_Ville = StrAC_Ville & "*"

         End Select

         Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & "[Cli_Code] In (SELECT AC_Code_Client FROM ADRESSE_CLIENT WHERE AC_Ville Like """ & StrAC_Ville & """;)"

         BolWhere = True

   End Select

   Select Case CadClass
      Case 0

        Forms(StrForm).SQLOrder = "Cli_Code"

      Case 1

        Forms(StrForm).SQLOrder = "Cli_Nom"

      Case 2

        Forms(StrForm).SQLOrder = "Cli_Groupe"

      Case 3

        Forms(StrForm).SQLOrder = "Cli_Groupe"

    End Select

    DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdOK_Click:

    Exit Sub

Err_CmdOK_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

    Resume Exit_CmdOK_Click
End Sub
