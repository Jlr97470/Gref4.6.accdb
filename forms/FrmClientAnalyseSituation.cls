Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private iMonnaieVisu As Integer
Private StrSQLWhere As String

Public Sub Inverseur(ByVal IntMonnaie As Integer)
    'Cette fonction permet de visualiser une fiche dans une monnaie particulière (argMonnaie)
    On Error GoTo Err_Inverseur

   Select Case IntMonnaie
       Case cFRANC

           TxtTotalDéb_Mont_Reversé_m1.ControlSource = "=SFrmClientAnalyseSituation.Form!TxtTotalDéb_Mont_Reversé_m"

           TxtTotalDéb_Mont_Reversé_m2.ControlSource = "=SFrmClientAnalyseSituation.Form!TxtTotalDéb_Mont_Reversé_m*DonnerTauxConvertion(1)"

           TxtTotalDéb_Mont_ResteAVerser_m1.ControlSource = "=SFrmClientAnalyseSituation.Form!TxtTotalDéb_Prin_Enc_m-SFrmClientAnalyseSituation.Form!TxtTotalDéb_Mont_Reversé_m"

           TxtTotalDéb_Mont_ResteAVerser_m2.ControlSource = "=(SFrmClientAnalyseSituation.Form!TxtTotalDéb_Prin_Enc_m-SFrmClientAnalyseSituation.Form!TxtTotalDéb_Mont_Reversé_m)*DonnerTauxConvertion(1)"

       Case cEURO

           TxtTotalDéb_Mont_Reversé_m1.ControlSource = "=SFrmClientAnalyseSituation.Form!TxtTotalDéb_Mont_Reversé_m*DonnerTauxConvertion(2)"

           TxtTotalDéb_Mont_Reversé_m2.ControlSource = "=SFrmClientAnalyseSituation.Form!TxtTotalDéb_Mont_Reversé_m"

           TxtTotalDéb_Mont_ResteAVerser_m1.ControlSource = "=(SFrmClientAnalyseSituation.Form!TxtTotalDéb_Prin_Enc_m-SFrmClientAnalyseSituation.Form!TxtTotalDéb_Mont_Reversé_m)*DonnerTauxConvertion(2)"

           TxtTotalDéb_Mont_ResteAVerser_m2.ControlSource = "=SFrmClientAnalyseSituation.Form!TxtTotalDéb_Prin_Enc_m-SFrmClientAnalyseSituation.Form!TxtTotalDéb_Mont_Reversé_m"

   End Select

   iMonnaieVisu = IntMonnaie

   CmbMonnaie = IntMonnaie

   SFrmClientAnalyseSituation.Form.Inverseur (IntMonnaie)

Exit_Inverseur:

    Exit Sub

Err_Inverseur:

    MsgBox Err.Number & " : " & Err.description, vbCritical, stSource & " : Inverseur"

    Resume Exit_Inverseur
End Sub

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   Inverseur (cEURO)

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

    Resume Exit_Form_Open
End Sub

Private Sub TxtDéb_Code_Client_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_TxtDéb_Code_Client_BeforeUpdate

   Select Case IsNull(DLookup("Cli_Code", "CLIENT", "Cli_Code = '" & Nz(TxtDéb_Code_Client) & "'")) And Not IsNull(TxtDéb_Code_Client)
      Case True

         Cancel = True

         MsgBox "Ce Client n'existe pas !" & vbCrLf & "Pour le créer, sélectionnez -Fiche Client- dans le menu principal.", vbInformation, stSource

      Case False

   End Select

Exit_TxtDéb_Code_Client_BeforeUpdate:

   Exit Sub

Err_TxtDéb_Code_Client_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDéb_Code_Client_BeforeUpdate"

   Resume Exit_TxtDéb_Code_Client_BeforeUpdate
End Sub

Private Sub TxtDéb_Code_Client_AfterUpdate()

   On Error GoTo Err_TxtDéb_Code_Client_AfterUpdate

   CmdAnalyse.Enabled = True

   CmdImprimer.Enabled = False

   TxtTotalDéb_Mont_Reversé_m1.Visible = False

   TxtTotalDéb_Mont_Reversé_m2.Visible = False

   TxtTotalDéb_Mont_ResteAVerser_m1.Visible = False

   TxtTotalDéb_Mont_ResteAVerser_m2.Visible = False

   SFrmClientAnalyseSituation.Visible = False

   SFrmClientAnalyseSituation.Form.RecordSource = ""

Exit_TxtDéb_Code_Client_AfterUpdate:

   Exit Sub

Err_TxtDéb_Code_Client_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDéb_Code_Client_AfterUpdate"

   Resume Exit_TxtDéb_Code_Client_AfterUpdate
End Sub

Private Sub TxtDéb_Date_CréDebut_AfterUpdate()

   On Error GoTo Err_TxtDéb_Date_CréDebut_AfterUpdate

   CmdImprimer.Enabled = False

   TxtTotalDéb_Mont_Reversé_m1.Visible = False

   TxtTotalDéb_Mont_Reversé_m2.Visible = False

   TxtTotalDéb_Mont_ResteAVerser_m1.Visible = False

   TxtTotalDéb_Mont_ResteAVerser_m2.Visible = False

   SFrmClientAnalyseSituation.Visible = False

   SFrmClientAnalyseSituation.Form.RecordSource = ""

Exit_TxtDéb_Date_CréDebut_AfterUpdate:

   Exit Sub

Err_TxtDéb_Date_CréDebut_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDéb_Date_CréDebut_AfterUpdate"

   Resume Exit_TxtDéb_Date_CréDebut_AfterUpdate
End Sub

Private Sub TxtDéb_Date_CréFin_AfterUpdate()

   On Error GoTo Err_TxtDéb_Date_CréFin_AfterUpdate

   CmdImprimer.Enabled = False

   TxtTotalDéb_Mont_Reversé_m1.Visible = False

   TxtTotalDéb_Mont_Reversé_m2.Visible = False

   TxtTotalDéb_Mont_ResteAVerser_m1.Visible = False

   TxtTotalDéb_Mont_ResteAVerser_m2.Visible = False

   SFrmClientAnalyseSituation.Visible = False

   SFrmClientAnalyseSituation.Form.RecordSource = ""

Exit_TxtDéb_Date_CréFin_AfterUpdate:

   Exit Sub

Err_TxtDéb_Date_CréFin_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDéb_Date_CréFin_AfterUpdate"

   Resume Exit_TxtDéb_Date_CréFin_AfterUpdate
End Sub

Private Sub CmbMonnaie_AfterUpdate()

   On Error GoTo Err_CmbMonnaie_AfterUpdate

   Inverseur CmbMonnaie

   CmdImprimer.Enabled = False

   TxtTotalDéb_Mont_Reversé_m1.Visible = False

   TxtTotalDéb_Mont_Reversé_m2.Visible = False

   TxtTotalDéb_Mont_ResteAVerser_m1.Visible = False

   TxtTotalDéb_Mont_ResteAVerser_m2.Visible = False

   SFrmClientAnalyseSituation.Visible = False

   SFrmClientAnalyseSituation.Form.RecordSource = ""

Exit_CmbMonnaie_AfterUpdate:

   Exit Sub

Err_CmbMonnaie_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbMonnaie_AfterUpdate"

   Resume Exit_CmbMonnaie_AfterUpdate
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

Private Sub CmdImprimer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.OpenForm "FrmClientAnalyseSituationFiltreEdition", acNormal, , , , acDialog, iMonnaieVisu & ";" & StrSQLWhere & ";Déb_Nom;Déb_Code;Déb_Réf1_Client"

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdClientListe_Click()

   On Error GoTo Err_CmdClientListe_Click

   DoCmd.OpenForm "FrmFicheClientListe", acNormal, , , acFormReadOnly, acDialog

   Select Case DLookup("Cli_Bloquer", "CLIENT", "Cli_Code =" & "'" & StrFicheClientListeCli_Code & "'")
      Case True

            MsgBox "Client en situation de blocage !"

      Case Else

   End Select

   Select Case StrFicheClientListeCli_Code
      Case ""

      Case Else

         TxtDéb_Code_Client = StrFicheClientListeCli_Code

         CmdAnalyse.Enabled = True

         CmdImprimer.Enabled = False

         TxtTotalDéb_Mont_Reversé_m1.Visible = False

         TxtTotalDéb_Mont_Reversé_m2.Visible = False

         TxtTotalDéb_Mont_ResteAVerser_m1.Visible = False

         TxtTotalDéb_Mont_ResteAVerser_m2.Visible = False

         SFrmClientAnalyseSituation.Visible = False

         SFrmClientAnalyseSituation.Form.RecordSource = ""

         TxtDéb_Date_CréDebut.SetFocus

   End Select

Exit_CmdClientListe_Click:

   Exit Sub

Err_CmdClientListe_Click:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_CmdClientListe_Click
End Sub

Private Sub CmdAnalyse_Click()
   Dim QdfRequete As DAO.QueryDef
   Dim StrSQL As String
   Dim StrDéb_Code_Client As String
   Dim StrDéb_Date_CréDebut As String
   Dim StrDéb_Date_CréFin As String
   Dim BolWhere As Boolean

   BolWhere = False

   StrDéb_Code_Client = Nz(TxtDéb_Code_Client)

   StrDéb_Date_CréDebut = Nz(TxtDéb_Date_CréDebut)

   StrDéb_Date_CréFin = Nz(TxtDéb_Date_CréFin)

   Select Case StrDéb_Code_Client
      Case ""

      Case Else

         StrSQLWhere = "[Déb_Code_Client]='" & StrDéb_Code_Client & "'"

         BolWhere = True

   End Select

   Select Case StrDéb_Date_CréDebut
      Case ""

      Case Else

         Select Case BolWhere
            Case True

               StrSQLWhere = StrSQLWhere & " AND [Déb_Date_Cré]>=" & ConvDate(StrDéb_Date_CréDebut)

            Case False

               StrSQLWhere = "[Déb_Date_Cré]>=" & ConvDate(StrDéb_Date_CréDebut)

               BolWhere = True

         End Select

   End Select

   Select Case StrDéb_Date_CréFin
      Case ""

      Case Else

         Select Case BolWhere
            Case True

               StrSQLWhere = StrSQLWhere & " AND [Déb_Date_Cré]<=" & ConvDate(StrDéb_Date_CréFin)

            Case False

               StrSQLWhere = "[Déb_Date_Cré]<=" & ConvDate(StrDéb_Date_CréFin)

               BolWhere = True

         End Select

   End Select

   Select Case BolWhere
      Case True

         StrSQLWhere = StrSQLWhere & " AND [Déb_MonnaieEnCours]=" & iMonnaieVisu

      Case False

         StrSQLWhere = "[Déb_MonnaieEnCours]=" & iMonnaieVisu

         BolWhere = True

   End Select

   SFrmClientAnalyseSituation.Visible = False

   SFrmClientAnalyseSituation.Form.RecordSource = ""

   For Each QdfRequete In CurrentDb.QueryDefs

      Select Case QdfRequete.Name
         Case "SelClientAnalyseSituation"

            CurrentDb.QueryDefs.Delete QdfRequete.Name

         Case Else

      End Select

   Next

   StrSQL = "SELECT DEBITEUR.Déb_Situation, DEBITEUR.Déb_Code, Round([DEBITEUR].[Déb_Mont_Prin_m1],2) AS Déb_Mont_Prin_m1, Round([DEBITEUR].[Déb_Mont_Prin_m2],2) AS Déb_Mont_Prin_m2, Round(Sum(IIf([CG_Code_Action]<>7 And [Enc_Situation] Is Not Null And [Enc_Montant_mo]=1,[Enc_Prin_m1],0)),2) AS Déb_Mont_Reversé_m1, Round(Sum(IIf([CG_CODE_Action]<>7 And [Enc_Situation] Is Not Null And [Enc_Montant_mo]=2,[Enc_Prin_m2],0)),2) AS Déb_Mont_Reversé_m2, Round(Sum(IIf([CG_Code_Action]<>7,[Enc_Prin_m1],0)),2) AS Déb_Prin_Enc_m1, Round(Sum(IIf([CG_Code_Action]<>7,[Enc_Prin_m2],0)),2) AS Déb_Prin_Enc_m2 " & _
                 "FROM CODE_GESTION INNER JOIN (DEBITEUR LEFT JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur) ON CODE_GESTION.CG_Code = DEBITEUR.Déb_Situation " & _
                 "WHERE " & StrSQLWhere & " " & _
                 "GROUP BY DEBITEUR.Déb_Situation, DEBITEUR.Déb_Code, Round([DEBITEUR].[Déb_Mont_Prin_m1],2), Round([DEBITEUR].[Déb_Mont_Prin_m2],2);"

   Set QdfRequete = CurrentDb.CreateQueryDef("SelClientAnalyseSituation", StrSQL)

   StrSQL = "SELECT CODE_GESTION.CG_Code, CODE_GESTION.CG_Code_Groupe, CODE_GESTION.CG_Libellé, Nz(Count([Déb_Code]),0) AS NbDebiteur, Nz(Sum(SelClientAnalyseSituation.Déb_Mont_Prin_m1),0) AS Déb_Mont_Prin_m1, Nz(Sum(SelClientAnalyseSituation.Déb_Mont_Prin_m2),0) AS Déb_Mont_Prin_m2, Nz(Sum(SelClientAnalyseSituation.Déb_Mont_Reversé_m1),0) AS Déb_Mont_Reversé_m1, Nz(Sum(SelClientAnalyseSituation.Déb_Mont_Reversé_m2),0) AS Déb_Mont_Reversé_m2, Nz(Sum(SelClientAnalyseSituation.Déb_Prin_Enc_m1),0) AS Déb_Prin_Enc_m1, Nz(Sum(SelClientAnalyseSituation.Déb_Prin_Enc_m2),0) AS Déb_Prin_Enc_m2 FROM CODE_GESTION LEFT JOIN SelClientAnalyseSituation ON CODE_GESTION.CG_Code=SelClientAnalyseSituation.Déb_Situation GROUP BY CODE_GESTION.CG_Code, CG_Code_Groupe, CODE_GESTION.CG_Libellé;"

   Screen.MousePointer = 11

   SFrmClientAnalyseSituation.Form.RecordSource = StrSQL

   CmdImprimer.Enabled = True

   TxtTotalDéb_Mont_Reversé_m1.Visible = True

   TxtTotalDéb_Mont_Reversé_m2.Visible = True

   TxtTotalDéb_Mont_ResteAVerser_m1.Visible = True

   TxtTotalDéb_Mont_ResteAVerser_m2.Visible = True

   SFrmClientAnalyseSituation.Visible = True

   Screen.MousePointer = 0

Exit_CmdAnalyse_Click:

   Exit Sub

Err_CmdAnalyse_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdAnalyse_Click"

   Resume Exit_CmdAnalyse_Click
End Sub
