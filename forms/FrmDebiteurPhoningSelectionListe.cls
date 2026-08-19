Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private SQLBase As String
Private SQLWhereGlobal As String
Private SQLWhere As String
Public SQLOrder As String

Private Function ImpVisu(ByVal IntVisu As Integer)

End Function

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   SQLBase = "SELECT DEBITEUR.*, ADRESSE_DEBITEUR.AD_CodePostal, ADRESSE_DEBITEUR.AD_Ville FROM DEBITEUR INNER JOIN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Code = ADRESSE_DEBITEUR.AD_Code_Débiteur"

   SQLWhereGlobal = ""

   SQLWhere = "Déb_Situation='31'"

   SQLOrder = ""

   Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
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
   Dim StrDéb_Code As String
   Dim StrDéb_Nom As String
   Dim LngCli_Type As Long
   Dim LngCG_Parent_Id As Long
   Dim LngCli_Groupe As Long
   Dim StrDéb_Date_ChèqueDebut As String
   Dim StrDéb_Date_ChèqueFin As String
   Dim StrDéb_Situation As String
   Dim IntDéb_niv_relance As Long
   Dim StrAD_CodePostal As String
   Dim StrAD_Ville As String

   On Error GoTo Err_CmdOK_Click

   StrDéb_Code = Nz(TxtDéb_Code)

   StrDéb_Nom = Nz(TxtDéb_Nom)

   LngCli_Type = Nz(CmbCli_Type)

   LngCG_Parent_Id = Nz(CmbCG_Parent_Id)

   LngCli_Groupe = Nz(CmbCli_Groupe)

   StrDéb_Date_ChèqueDebut = Nz(TxtDéb_Date_ChèqueDebut)

   StrDéb_Date_ChèqueFin = Nz(TxtDéb_Date_ChèqueFin)

   StrDéb_Situation = Nz(TxtDéb_Situation)

   IntDéb_niv_relance = Nz(TxtDéb_niv_relance)

   StrAD_CodePostal = Nz(TxtAD_CodePostal)

   StrAD_Ville = Nz(TxtAD_Ville)

   BolWhere = False

   SQLWhere = ""

   Select Case StrDéb_Code
      Case ""

      Case Else

            Select Case InStr(StrDéb_Code, "*")
               Case Is > 0

               Case Else

                 StrDéb_Code = StrDéb_Code & "*"

            End Select

            SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Code] LIKE """ & StrDéb_Code & """"

            BolWhere = True

   End Select

   Select Case StrDéb_Nom
      Case ""

      Case Else

         Select Case InStr(StrDéb_Nom, "*")
            Case Is > 0

            Case Else

               StrDéb_Nom = StrDéb_Nom & "*"

         End Select

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Nom] LIKE """ & StrDéb_Nom & """"

         BolWhere = True

   End Select

   Select Case LngCli_Type
      Case 0

      Case Else

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Code_Client] In (SELECT Cli_Code FROM CLIENT WHERE Cli_Type = " & LngCli_Type & ";)"

         BolWhere = True

   End Select

   Select Case LngCG_Parent_Id <> 0 And LngCli_Groupe = 0
      Case True

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Code_Client] In (SELECT Cli_Code FROM CLIENT INNER JOIN CLIENT_GROUPE ON CLIENT.Cli_Groupe = CLIENT_GROUPE.CG_Id WHERE CG_Parent_Id = " & LngCG_Parent_Id & ";)"

         BolWhere = True

      Case False

   End Select

   Select Case LngCli_Groupe
      Case 0

      Case Else

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Code_Client] In (SELECT Cli_Code FROM CLIENT WHERE Cli_Groupe = " & LngCli_Groupe & ";)"

         BolWhere = True

   End Select

   Select Case StrDéb_Date_ChèqueDebut
      Case ""

      Case Else

         StrDéb_Date_ChèqueDebut = ConvDate(StrDéb_Date_ChèqueDebut)

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Date_Chèque]>=" & StrDéb_Date_ChèqueDebut

         BolWhere = True

   End Select

   Select Case StrDéb_Date_ChèqueFin
      Case ""

      Case Else

         StrDéb_Date_ChèqueDebut = ConvDate(StrDéb_Date_ChèqueFin)

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Date_Chèquefin]<=" & StrDéb_Date_ChèqueFin

         BolWhere = True

   End Select

   Select Case StrDéb_Situation
      Case ""

      Case Else

         Select Case InStr(StrDéb_Situation, "*")
            Case Is > 0

            Case Else

               StrDéb_Situation = StrDéb_Situation & "*"

         End Select

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Situation] LIKE """ & StrDéb_Situation & """"

         BolWhere = True

   End Select

   Select Case IntDéb_niv_relance
      Case 0

      Case Else

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_niv_relance] = " & IntDéb_niv_relance

         BolWhere = True

   End Select

   Select Case StrAD_CodePostal
      Case ""

      Case Else

         Select Case InStr(StrAD_CodePostal, "*")
            Case Is > 0

            Case Else

               StrAD_CodePostal = StrAD_CodePostal & "*"

         End Select

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [AD_CodePostal] LIKE """ & StrAD_CodePostal & """"

         BolWhere = True

   End Select

   Select Case StrAD_Ville
      Case ""

      Case Else

         Select Case InStr(StrAD_Ville, "*")
            Case Is > 0

            Case Else

               StrAD_Ville = StrAD_Ville & "*"

         End Select

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [AD_Ville] LIKE """ & StrAD_Ville & """"

         BolWhere = True

   End Select

   Select Case CadClass
      Case 1

         Select Case SQLWhere
            Case ""

            Case Else

               SQLOrder = "Déb_Code"

         End Select

      Case 2

         SQLOrder = "Déb_Nom"

      Case 3

         SQLOrder = "Déb_Date_Chèque"

      Case 4

         SQLOrder = "Déb_Situation"

      Case 5

         SQLOrder = "Déb_niv_relance"

      Case 6

         SQLOrder = "AD_CodePostal"

      Case 7

         SQLOrder = "AD_Ville"

   End Select

   Select Case SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"
      Case Me.RecordSource

      Case Else

         Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

   End Select

Exit_CmdOK_Click:

   Exit Sub

Err_CmdOK_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

   Resume Exit_CmdOK_Click
End Sub

Private Sub CmdDetail_Click()

   On Error GoTo Err_CmdDetail_Click

   DoCmd.Minimize

   DoCmd.OpenForm "FrmDebiteurPhoningListe", acNormal, , "Déb_Nom='" & Déb_Nom & "' AND AD_CodePostal='" & AD_CodePostal & "'", , acWindowNormal, Déb_Code & ";"

Exit_CmdDetail_Click:

   Exit Sub

Err_CmdDetail_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdDetail_Click"

   Resume Exit_CmdDetail_Click
End Sub
