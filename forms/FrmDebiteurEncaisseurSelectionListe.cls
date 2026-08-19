Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private SQLBase As String
Private SQLWhereGlobal As String
Private SQLWhere As String
Private SQLOrder As String

Private Function ImpVisu(ByVal IntVisu As Integer)
   Dim StrSQL As String
   Dim IntRaportEtat As Integer

   On Error GoTo Err_ImpVisu

   On Error Resume Next

   DoCmd.OpenReport "EtaCourrierFicheEncaisseur", IntVisu, , IIf(SQLWhere <> "", SQLWhere & " AND ", SQLWhere) & "Déb_MonnaieEnCours=" & cFRANC, acWindowNormal, cFRANC & ";"

   Select Case Err.Number
      Case 0

      Case 2501

         Err.Clear

      Case Else

         GoTo Err_ImpVisu

   End Select

   On Error GoTo Err_ImpVisu

    Do While DonneEtatRaport("EtaCourrierFicheEncaisseur", IntRaportEtat) = True

       DoEvents

    Loop

    On Error Resume Next

   DoCmd.OpenReport "EtaCourrierFicheEncaisseur", IntVisu, , IIf(SQLWhere <> "", SQLWhere & " AND ", SQLWhere) & "Déb_MonnaieEnCours=" & cEURO, acWindowNormal, cEURO & ";"

   Select Case Err.Number
      Case 0

      Case 2501

         Err.Clear

      Case Else

         GoTo Err_ImpVisu

   End Select

   On Error GoTo Err_ImpVisu

    Do While DonneEtatRaport("EtaCourrierFicheEncaisseur", IntRaportEtat) = True

       DoEvents

    Loop

   Select Case IntVisu
      Case 0

         StrSQL = "UPDATE CODE_GESTION INNER JOIN (DEBITEUR INNER JOIN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Code = ADRESSE_DEBITEUR.AD_Code_Débiteur) ON CODE_GESTION.CG_Code = DEBITEUR.Déb_Situation SET DEBITEUR.Déb_Code_Enc = 'S' & [Déb_Code_Enc], DEBITEUR.Déb_Date_proch_rel = Date()"

         StrSQL = StrSQL & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & ";"

         CurrentDb.Execute StrSQL, dbFailOnError

      Case 2

   End Select

   Me.Requery

Exit_ImpVisu:

   Exit Function

Err_ImpVisu:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ImpVisu"

   Resume Exit_ImpVisu
End Function

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   SQLBase = "SELECT DEBITEUR.*, ADRESSE_DEBITEUR.AD_CodePostal, ADRESSE_DEBITEUR.AD_Ville FROM (CODE_GESTION INNER JOIN DEBITEUR ON CODE_GESTION.CG_Code = DEBITEUR.Déb_Situation) INNER JOIN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Code = ADRESSE_DEBITEUR.AD_Code_Débiteur"

   SQLWhereGlobal = "(((CODE_GESTION.CG_Code_Action)=3) AND ((Len([Déb_Code_Enc]))=2))"

   SQLWhere = ""

   SQLOrder = "Déb_Code"

   Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
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
   Dim StrDéb_Date_der_relDebut As String
   Dim StrDéb_Date_der_relFin As String
   Dim StrDéb_Code_Enc As String
   Dim StrAD_CodePostal As String
   Dim StrAD_Ville As String

   On Error GoTo Err_CmdOK_Click

   Screen.MousePointer = 11

   StrDéb_Code = Nz(TxtDéb_Code)

   StrDéb_Nom = Nz(TxtDéb_Nom)

   StrDéb_Date_der_relDebut = Nz(TxtDéb_Date_der_relDebut)

   StrDéb_Date_der_relFin = Nz(TxtDéb_Date_der_relFin)

   StrDéb_Code_Enc = Nz(TxtDéb_Code_Enc)

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

   Select Case StrDéb_Date_der_relDebut
      Case ""

      Case Else

         StrDéb_Date_der_relDebut = ConvDate(StrDéb_Date_der_relDebut)

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Date_der_rel]>=" & StrDéb_Date_der_relDebut

         BolWhere = True

   End Select

   Select Case StrDéb_Date_der_relFin
      Case ""

      Case Else

         StrDéb_Date_der_relFin = ConvDate(StrDéb_Date_der_relFin)

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Date_der_rel]<=" & StrDéb_Date_der_relFin

         BolWhere = True

   End Select

   Select Case StrDéb_Code_Enc
      Case ""

      Case Else

         Select Case InStr(StrDéb_Code_Enc, "*")
            Case Is > 0

            Case Else

               StrDéb_Code_Enc = StrDéb_Code_Enc & "*"

         End Select

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & " [Déb_Code_Enc] LIKE """ & StrDéb_Code_Enc & """"

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

         SQLOrder = "Déb_Code"

      Case 2

         SQLOrder = "Déb_Nom"

      Case 3

         SQLOrder = "Déb_Date_der_rel"

      Case 4

         SQLOrder = "Déb_Code_Enc"

      Case 5

         SQLOrder = "AD_CodePostal"

      Case 6

         SQLOrder = "AD_Ville"

   End Select

   Select Case SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"
      Case Me.RecordSource

      Case Else

         Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

   End Select

Exit_CmdOK_Click:

   Screen.MousePointer = 0

   Exit Sub

Err_CmdOK_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

   Resume Exit_CmdOK_Click
End Sub
