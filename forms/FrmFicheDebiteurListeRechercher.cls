Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private StrForm As String
Private SQLBase As String
Private SQLWhere As String
Private SQLOrder As String

Private Function ImpVisu(ByVal IntVisu As Integer)
   Dim IntRaportEtat As Integer
   Dim IntRepMsg As Integer

   On Error GoTo Err_ImpVisu

   DoCmd.OpenReport "EtaListeDebiteurDetail", acDesign

   Reports("EtaListeDebiteurDetail").Filter = SQLWhere

   Reports("EtaListeDebiteurDetail").FilterOn = True

   DoCmd.Close acReport, "EtaListeDebiteurDetail", acSavePrompt

   Select Case IntVisu
      Case 0, 2

         On Error Resume Next

         DoCmd.OpenReport "EtaListeDebiteurDetail", IntVisu

         Select Case Err.Number
            Case 0

            Case 2501

               Err.Clear

            Case Else

               GoTo Err_ImpVisu

         End Select

         On Error GoTo Err_ImpVisu

         Do While DonneEtatRaport("EtaListeDebiteurDetail", IntRaportEtat) = True

            DoEvents

         Loop

      Case 5

         IntRepMsg = MsgBox("Souhaitez-vous exporter les données au format Excel ?" & vbCrLf & "Indiquez OUI pour extraire les données au format excel et NON pour le format texte.", vbYesNo + vbQuestion, "Format d'exportation")

         Select Case IntRepMsg
            Case vbYes

               DoCmd.OutputTo acOutputReport, "EtaListeDebiteurDetail", acFormatXLS, , False

            Case vbNo

               DoCmd.OutputTo acOutputReport, "EtaListeDebiteurDetail", acFormatRTF, , False

         End Select

   End Select

Exit_ImpVisu:

   Exit Function

Err_ImpVisu:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Err_ImpVisu"

   Resume Exit_ImpVisu
End Function

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   StrForm = Nz(Me.OpenArgs)

   SQLBase = "SELECT DEBITEUR.*, ADRESSE_DEBITEUR.AD_CodePostal, ADRESSE_DEBITEUR.AD_Ville FROM DEBITEUR INNER JOIN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Code=ADRESSE_DEBITEUR.AD_Code_Débiteur"

   SQLWhere = "[Déb_MonnaieEnCours]=2"

   SQLOrder = ""

   Me.RecordSource = SQLBase & IIf((SQLWhere) = "", " ", " WHERE ") & SQLWhere & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub TxtDéb_Situation1_AfterUpdate()
   Dim IntCode_Action As Integer
   Dim IntDelai As Integer
   Dim IntGrilleTauxNum As Integer
   Dim IntRep As Integer

   On Error GoTo Err_TxtDéb_Situation1_AfterUpdate

   IntCode_Action = Nz(DLookup("CG_Code_Action", "CODE_GESTION", "CG_Code='" & Déb_Situation & "'"), 0)

   Select Case IntCode_Action
      Case 1

         TxtDéb_niv_relance = DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=" & IntCode_Action)

         Déb_Date_proch_rel = Null

      Case 2

         TxtDéb_niv_relance = DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=" & IntCode_Action)

         IntGrilleTauxNum = DLookup("Cli_GrilleTaux", "CLIENT", "Cli_Code = '" & Déb_Code_Client & "'") '*** Numéro de la grille de taux du client

         IntDelai = DLookup("Tx_Délai", "TAUX", "Tx_N° = 10 AND Tx_Grille = " & IntGrilleTauxNum & " AND Tx_Code_Action = 2")

         IntRep = MsgBox("Une ""Fiche encaisseur"" sera éditée le " & CDate(Date + 1) & " pour le débiteur " & Déb_Code & ".", 68)

         Select Case IntRep
            Case 6

               Déb_Date_der_rel = Nz(Déb_Date_der_rel, Date)

               Déb_Date_proch_rel = Date + IntDelai

            Case Else

               Déb_Date_der_rel = Null

               Déb_Date_proch_rel = Null

         End Select

      Case 3, 5, 6

         Déb_niv_relance = DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=" & IntCode_Action)

         Déb_Date_der_rel = Null

         Déb_Date_proch_rel = Null

      Case 4, 7, 8, 9

         Déb_niv_relance = Nz(DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=" & IntCode_Action & " AND Rel_N°=" & Nz(Déb_niv_relance, 0)), 1)

   End Select

   TxtDéb_niv_relance_AfterUpdate

Exit_TxtDéb_Situation1_AfterUpdate:

   Exit Sub

Err_TxtDéb_Situation1_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDéb_Situation1_AfterUpdate"

   Resume Exit_TxtDéb_Situation1_AfterUpdate
End Sub

Private Sub TxtDéb_niv_relance_AfterUpdate()
   Dim IntRep As Integer
   Dim IntMaxRelanceCli As Integer

   On Error GoTo Err_TxtDéb_niv_relance_AfterUpdate

   Select Case DLookup("CG_Code_Action", "CODE_GESTION", "CG_Code = '" & Déb_Situation & "'")
      Case 1

         IntMaxRelanceCli = DLookup("Cli_Nb_Relance", "CLIENT", "Cli_Code = '" & Déb_Code_Client & "'") '*** Nb maxi de relances accepté par le client pour ses débiteurs

         Select Case (Déb_niv_relance >= IntMaxRelanceCli)
            Case True

               TxtDéb_niv_relance = IntMaxRelanceCli

               IntRep = MsgBox("Une ""Fiche encaisseur"" sera éditée le " & CDate(Date + 1) & " pour le débiteur " & Déb_Code & ".", 68)

               Select Case IntRep
                  Case 6

                     Déb_Date_der_rel = Nz(Déb_Date_der_rel, Date)

                     Déb_Date_proch_rel = CDate(Date + 1)

                  Case Else

               End Select

            Case False

        End Select

      Case Else

   End Select

Exit_TxtDéb_niv_relance_AfterUpdate:

   Exit Sub

Err_TxtDéb_niv_relance_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDéb_niv_relance_AfterUpdate"

   Resume Exit_TxtDéb_niv_relance_AfterUpdate
End Sub

Private Sub ChkDéb_Npai1_AfterUpdate()
   Dim IntMaxRelanceCli As Integer
   Dim IntRep As Integer

   On Error GoTo Err_ChkDéb_Npai1_AfterUpdate

   Select Case ChkDéb_Npai1
      Case True

         IntMaxRelanceCli = DLookup("Cli_Nb_Relance", "CLIENT", "Cli_Code = '" & Déb_Code_Client & "'") '*** Nb maxi de relances accepté par le client pour ses débiteurs

         IntRep = MsgBox("Une ""Fiche encaisseur"" sera éditée le " & CDate(Date + 1) & " pour le débiteur " & Déb_Code & ".", 68)

         Select Case IntRep
            Case 6

               Déb_Situation = 10

               Déb_niv_relance = IntMaxRelanceCli

               Déb_Date_der_rel = Nz(Déb_Date_der_rel, Date)

               Déb_Date_proch_rel = CDate(Date + 1)

            Case Else

         End Select

      Case False

   End Select

Exit_ChkDéb_Npai1_AfterUpdate:

   Exit Sub

Err_ChkDéb_Npai1_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ChkDéb_Npai1_AfterUpdate"

   Resume Exit_ChkDéb_Npai1_AfterUpdate
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   StrFicheDebiteurListeDéb_Code = ""

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
   Dim IntDéb_MonnaieEnCours As Integer
   Dim StrDéb_Réf_Client As String
   Dim StrDéb_Code_Enc As String
   Dim StrDéb_Situation As String
   Dim StrAD_CodePostal As String
   Dim StrAD_Ville As String
   Dim StrDéb_Npai As String

   On Error GoTo Err_CmdOK_Click

   Screen.MousePointer = 11

   StrDéb_Code = Nz(TxtDéb_Code)

   StrDéb_Nom = Nz(TxtDéb_Nom)

   StrDéb_Mont_Prin_m = Nz(TxtDéb_Mont_Prin_m)

   IntDéb_MonnaieEnCours = Nz(CmbDéb_MonnaieEnCours)

   StrDéb_Réf_Client = Nz(TxtDéb_Réf_Client)

   StrDéb_Code_Enc = Nz(TxtDéb_Code_Enc)

   StrDéb_Situation = Nz(TxtDéb_Situation)

   StrAD_CodePostal = Nz(TxtAD_CodePostal)

   StrAD_Ville = Nz(TxtAD_Ville)

   StrDéb_Npai = Nz(ChkDéb_Npai)

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

            SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & "[Déb_Code] LIKE """ & StrDéb_Code & """"

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

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & "[Déb_Nom] LIKE """ & StrDéb_Nom & """"

         BolWhere = True

   End Select

   Select Case StrDéb_Mont_Prin_m
      Case ""

      Case Else

         Select Case InStr(StrDéb_Mont_Prin_m, "*")
            Case Is > 0

            Case Else

               StrDéb_Mont_Prin_m = StrDéb_Mont_Prin_m & "*"

         End Select

         Select Case IntDéb_MonnaieEnCours
            Case cFRANC

               SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & "Round([Déb_Mont_Prin_m1],2) LIKE """ & StrDéb_Mont_Prin_m & """"

            Case Else

               SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & "Round([Déb_Mont_Prin_m2],2) LIKE """ & StrDéb_Mont_Prin_m & """"

               IntDéb_MonnaieEnCours = cEURO

               CmbDéb_MonnaieEnCours = cEURO

         End Select

         BolWhere = True

   End Select

   Select Case StrDéb_Réf_Client
      Case ""

      Case Else

         Select Case InStr(StrDéb_Réf_Client, "*")
            Case Is > 0

            Case Else

               StrDéb_Réf_Client = StrDéb_Réf_Client & "*"

         End Select

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & "[Déb_Réf1_Client] LIKE """ & StrDéb_Réf_Client & """" & " OR [Déb_Réf2_Client] LIKE """ & StrDéb_Réf_Client & """"

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

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & "[Déb_Code_Enc] LIKE """ & StrDéb_Code_Enc & """"

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

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & "[Déb_Situation] LIKE """ & StrDéb_Situation & """"

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

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & "[AD_CodePostal] LIKE """ & StrAD_CodePostal & """"

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

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & "[AD_Ville] LIKE """ & StrAD_Ville & """"

         BolWhere = True

   End Select

   Select Case StrDéb_Npai
      Case ""

      Case Else

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & "[Déb_Npai] =" & StrDéb_Npai

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

         SQLOrder = "Déb_Mont_Prin_m1"

      Case 4

         SQLOrder = "Déb_Mont_Prin_mo"

      Case 5

         SQLOrder = "Déb_Code_Enc"

      Case 6

         SQLOrder = "Déb_Situation"

      Case 7

         SQLOrder = "AD_CodePostal"

      Case 8

         SQLOrder = "AD_Ville"

   End Select

   Select Case IntDéb_MonnaieEnCours
      Case 0

         TxtMon_Nom.Visible = True

         TxtDéb_Mont_Prin_m1.Visible = False

      Case Else

         TxtMon_Nom.Visible = False

         TxtDéb_Mont_Prin_m1.Visible = True

         SQLWhere = SQLWhere & IIf(BolWhere, " AND ", "") & "[Déb_MonnaieEnCours]=" & IntDéb_MonnaieEnCours

         BolWhere = True

         Select Case IntDéb_MonnaieEnCours
             Case cFRANC

                 TxtDéb_Mont_Prin_m1.ControlSource = "Déb_Mont_Prin_m1"

             Case cEURO

                 TxtDéb_Mont_Prin_m1.ControlSource = "Déb_Mont_Prin_m2"

         End Select

   End Select

   Select Case SQLBase & IIf((SQLWhere) = "", " ", " WHERE ") & SQLWhere & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"
      Case Me.RecordSource

      Case Else

         Me.RecordSource = SQLBase & IIf((SQLWhere) = "", " ", " WHERE ") & SQLWhere & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

   End Select

Exit_CmdOK_Click:

   Screen.MousePointer = 0

   Exit Sub

Err_CmdOK_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

   Resume Exit_CmdOK_Click
End Sub

Private Sub CmdDetail_Click()
   Dim IntEtat As Integer

   On Error GoTo Err_CmdDetail_Click

   DoCmd.RunCommand acCmdSaveRecord

   Select Case StrForm
      Case ""

         StrFicheDebiteurListeDéb_Code = Déb_Code

         DoCmd.Close acForm, Me.Name, acSavePrompt

      Case Else

         Select Case DonneEtatFormulaire(StrForm, IntEtat)
            Case True

               DoCmd.Minimize

               Forms(StrForm).Filter = "Déb_Code='" & Déb_Code & "'"

               Forms(StrForm).FilterOn = True

            Case False

         End Select

   End Select

Exit_CmdDetail_Click:

   Exit Sub

Err_CmdDetail_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdDetail_Click"

   Resume Exit_CmdDetail_Click
End Sub
