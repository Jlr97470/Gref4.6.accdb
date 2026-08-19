Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private iMonnaieVisu As Integer 'Monnaie dans laquelle est visualisée la fiche en cours

Public Sub Inverseur(ByVal IntMonnaie As Integer)
    'Cette fonction permet de visualiser une fiche dans une monnaie particulière
   On Error GoTo Err_Inverseur

   Select Case IntMonnaie
      Case cFRANC

         Select Case Cli_MonnaieEnCours <> IntMonnaie
            Case True

               EtiMonnaieVue1.Caption = "Vue en FRANC"

               EtiMonnaieVue2.Caption = "Vue en FRANC"

            Case False

               EtiMonnaieVue1.Caption = ""

               EtiMonnaieVue2.Caption = ""

           End Select

           TxtCli_Mont_Hon_m.ControlSource = "Cli_Mont_Hon_m1"

      Case cEURO

         Select Case Cli_MonnaieEnCours <> IntMonnaie
            Case True

               EtiMonnaieVue1.Caption = "Vue en EURO"

               EtiMonnaieVue2.Caption = "Vue en EURO"

            Case False

               EtiMonnaieVue1.Caption = ""

               EtiMonnaieVue2.Caption = ""

           End Select

           TxtCli_Mont_Hon_m.ControlSource = "Cli_Mont_Hon_m2"

   End Select

   iMonnaieVisu = IntMonnaie

Exit_Inverseur:

   Exit Sub

Err_Inverseur:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Inverseur"

   Resume Exit_Inverseur
End Sub

Private Sub ChangerMonnaie(IntMonnaie As Integer)
   Dim StrMotDePasse  As String
   Dim StrReponse As String
   'Cette procédure bascule une fiche d'une monnaie vers l'autre
   On Error GoTo Err_ChangerMonnaie

   StrMotDePasse = DLookup("Mot_de_Passe", "PARAMETRE_GREF")

   StrReponse = InputBox("Pour basculer ce dossier en " & DLookup("Mon_nom", "MONNAIE", "Mon_Id=" & IntMonnaie) & ", veuillez inscrire le mot de passe :", "Mot de passe")

   Select Case StrReponse
      Case StrMotDePasse

      Case Else

        MsgBox "Mot de passe incorrect !", 48

        GoTo Exit_ChangerMonnaie

   End Select

   Cli_MonnaieEnCours = IntMonnaie

   Select Case IntMonnaie
       Case cFRANC

           CmdBasculer.Enabled = True

       Case cEURO

           TxtCli_Nom.SetFocus 'permet de désélectionner le bouton "BasculerEuro avant de le désactiver"

           CmdBasculer.Enabled = False

   End Select

   Inverseur IntMonnaie

Exit_ChangerMonnaie:

   Exit Sub

Err_ChangerMonnaie:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ChangerMonnaie"

   Resume Exit_ChangerMonnaie
End Sub

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   Inverseur cEURO

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub Form_Load()

   On Error GoTo Err_Form_Load

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Load"

   Resume Exit_Form_Load
End Sub

Private Sub Form_Current()

   On Error GoTo Err_Form_Current

   Inverseur Cli_MonnaieEnCours

   TxtCode_Client = Cli_Code

   CmbCG_Parent_Id = DLookup("CG_Parent_Id", "CLIENT_GROUPE", "CG_Id=" & Nz(Cli_Groupe, 0))

   CmbCli_Groupe.Requery

   CmdBasculer.Enabled = (Cli_MonnaieEnCours = cFRANC)

   Select Case Cli_MonnaieEnCours
      Case cFRANC

         CmdBasculer.Caption = "Basculer En " & DLookup("Mon_nom", "MONNAIE", "Mon_Id=" & cEURO)

      Case cEURO

         CmdBasculer.Caption = "Basculer En " & DLookup("Mon_nom", "MONNAIE", "Mon_Id=" & cFRANC)

   End Select

   CmdBasculer.Requery

   CmdInverseur.Enabled = True

   CmbCli_Grille_Frais.Enabled = (Not Cli_NewCalculFrais)

   CmdDebiteurNouveau.Enabled = (Not Cli_Bloquer)

   CmdDebiteurListe.Enabled = True

   TxtCli_Mont_Hon_m.Enabled = True

   CmdContratNouveau.Enabled = True

Exit_Form_Current:

   Exit Sub

Err_Form_Current:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Current"

   Resume Exit_Form_Current
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)

   On Error GoTo Err_Form_AfterDelConfirm

   TxtCode_Client = Null

   CmdDebiteurNouveau.Enabled = False

   CmdDebiteurListe.Enabled = False

Exit_Form_AfterDelConfirm:

   Exit Sub

Err_Form_AfterDelConfirm:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_AfterDelConfirm"

   Resume Exit_Form_AfterDelConfirm
End Sub

Private Sub TxtCode_Client_AfterUpdate()

   On Error GoTo Err_TxtCode_Client_AfterUpdate

   Me.Filter = "Cli_Code='" & TxtCode_Client & "'"

   Me.FilterOn = True

Exit_TxtCode_Client_AfterUpdate:

   Exit Sub

Err_TxtCode_Client_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtCode_Client_AfterUpdate"

   Resume Exit_TxtCode_Client_AfterUpdate
End Sub

Private Sub TxtCode_Client_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_TxtCode_Client_BeforeUpdate

   Select Case DLookup("Cli_Code", "CLIENT", "Cli_Code='" & TxtCode_Client & "'")
      Case TxtCode_Client

      Case Else

         Cancel = True

         MsgBox "Identifiant incorrect !", vbExclamation, stSource

   End Select

Exit_TxtCode_Client_BeforeUpdate:

   Exit Sub

Err_TxtCode_Client_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtCode_Client_BeforeUpdate"

   Resume Exit_TxtCode_Client_BeforeUpdate
End Sub

Private Sub ChkCli_Bloquer_AfterUpdate()

   On Error GoTo Err_ChkCli_Bloquer_AfterUpdate

   CmdDebiteurNouveau.Enabled = Not Cli_Bloquer

Exit_ChkCli_Bloquer_AfterUpdate:

   Exit Sub

Err_ChkCli_Bloquer_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ChkCli_Bloquer_AfterUpdate"

   Resume Exit_ChkCli_Bloquer_AfterUpdate
End Sub

Private Sub GrpCli_NewCalculFrais_AfterUpdate()

   On Error GoTo Err_GrpCli_NewCalculFrais_AfterUpdate

   CmbCli_Grille_Frais.Enabled = Not Cli_NewCalculFrais

Exit_GrpCli_NewCalculFrais_AfterUpdate:

   Exit Sub

Err_GrpCli_NewCalculFrais_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : GrpCli_NewCalculFrais_AfterUpdate"

   Resume Exit_GrpCli_NewCalculFrais_AfterUpdate
End Sub

Private Sub TxtCli_Mont_Hon_m_AfterUpdate()

   On Error GoTo Err_TxtCli_Mont_Hon_m_AfterUpdate
   'Mise à jour des valeurs de l'autre monnaie
   Select Case Cli_MonnaieEnCours
      Case cFRANC

         Cli_Mont_Hon_m2 = Cli_Mont_Hon_m1 * DonnerTauxConvertion(iMonnaieVisu)

      Case cEURO

         Cli_Mont_Hon_m1 = Cli_Mont_Hon_m2 * DonnerTauxConvertion(iMonnaieVisu)

   End Select

   Cli_Mont_Hon_mo = Cli_MonnaieEnCours

   Select Case True
      Case TxtCli_Mont_Hon_m > 0 And Cli_Pourcentage = 0

      Case TxtCli_Mont_Hon_m = 0 And Cli_Pourcentage > 0

      Case TxtCli_Mont_Hon_m = 0 And Cli_Pourcentage = 0

        MsgBox "Le type de contrat est obligatoire.", vbInformation, stSource

      Case Else

        MsgBox "Ne choisissez qu'un seul type de Contrat.", vbInformation, stSource

   End Select

Exit_TxtCli_Mont_Hon_m_AfterUpdate:

   Exit Sub

Err_TxtCli_Mont_Hon_m_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtCli_Mont_Hon_m_AfterUpdate"

   Resume Exit_TxtCli_Mont_Hon_m_AfterUpdate
End Sub

Private Sub TxtCli_Pourcentage_AfterUpdate()

   On Error GoTo Err_TxtCli_Pourcentage_AfterUpdate

   Select Case True
      Case TxtCli_Mont_Hon_m > 0 And Cli_Pourcentage = 0

      Case TxtCli_Mont_Hon_m = 0 And Cli_Pourcentage > 0

      Case TxtCli_Mont_Hon_m = 0 And Cli_Pourcentage = 0

        MsgBox "Le type de contrat est obligatoire.", vbInformation, stSource

      Case Else

        MsgBox "Ne choisissez qu'un seul type de Contrat.", vbInformation, stSource

   End Select

Exit_TxtCli_Pourcentage_AfterUpdate:

   Exit Sub

Err_TxtCli_Pourcentage_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtCli_Pourcentage_AfterUpdate"

   Resume Exit_TxtCli_Pourcentage_AfterUpdate
End Sub

Private Sub TxtCli_Mont_Hon_m_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_TxtCli_Mont_Hon_m_BeforeUpdate

   Select Case IsNull(TxtCli_Mont_Hon_m)
      Case True

         Cancel = True

         MsgBox "Le montant des honoraires est obligatoire"

      Case False

   End Select

Exit_TxtCli_Mont_Hon_m_BeforeUpdate:

   Exit Sub

Err_TxtCli_Mont_Hon_m_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtCli_Mont_Hon_m_BeforeUpdate"

   Resume Exit_TxtCli_Mont_Hon_m_BeforeUpdate
End Sub

Private Sub TxtCli_Pourcentage_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_TxtCli_Pourcentage_BeforeUpdate

   Select Case IsNull(TxtCli_Pourcentage)
      Case True

         Cancel = True

         MsgBox "Le pourcentage est obligatoire"

      Case False

   End Select

Exit_TxtCli_Pourcentage_BeforeUpdate:

   Exit Sub

Err_TxtCli_Pourcentage_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtCli_Pourcentage_BeforeUpdate"

   Resume Exit_TxtCli_Pourcentage_BeforeUpdate
End Sub

Private Sub TxtCli_Code_Etab_AfterUpdate()

   On Error GoTo Err_TxtCli_Code_Etab_AfterUpdate

   Select Case IsNull(Cli_Code_Etab)
      Case True

         Cli_Ttt_Auto = False

      Case False

         Cli_Ttt_Auto = True

    End Select

Exit_TxtCli_Code_Etab_AfterUpdate:

   Exit Sub

Err_TxtCli_Code_Etab_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtCli_Code_Etab_AfterUpdate"

   Resume Exit_TxtCli_Code_Etab_AfterUpdate
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo Err_Form_KeyDown

   Select Case KeyCode
      Case 107

         CmdSuivant_Click

         CmdSuivant.SetFocus

      Case 109

         CmdPrecedent_Click

         CmdPrecedent.SetFocus

   End Select

Exit_Form_KeyDown:

   Exit Sub

Err_Form_KeyDown:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " :"

   Resume Exit_Form_KeyDown
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdFermer_Click

   End Select

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdPrecedent_Click()
   Dim StrCli_Code As String

   On Error GoTo Err_CmdPrecedent_Click

   StrCli_Code = Nz(TxtCode_Client)

   Me.Filter = ""

   Me.FilterOn = True

   Me.Recordset.FindLast "Cli_Code<'" & StrCli_Code & "'"

   Select Case StrCli_Code
      Case ""

      Case Else

         Me.Filter = "Cli_Code='" & TxtCode_Client & "'"

         Me.FilterOn = True

         TxtCli_Nom.SetFocus

   End Select

Exit_CmdPrecedent_Click:

   Exit Sub

Err_CmdPrecedent_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdPrecedent_Click"

   Resume Exit_CmdPrecedent_Click
End Sub

Private Sub CmdSuivant_Click()
   Dim StrCli_Code As String

   On Error GoTo Err_CmdSuivant_Click

   'StrCli_Code = Nz(DMin("Cli_Code", "CLIENT", "Cli_Code>'" & TxtCode_Client & "'"))

   StrCli_Code = Nz(TxtCode_Client)

   Me.Filter = ""

   Me.FilterOn = True

   Me.Recordset.FindFirst "Cli_Code>'" & StrCli_Code & "'"

   Select Case StrCli_Code
      Case ""

      Case Else

         Me.Filter = "Cli_Code='" & TxtCode_Client & "'"

         Me.FilterOn = True

         TxtCli_Nom.SetFocus

   End Select

Exit_CmdSuivant_Click:

   Exit Sub

Err_CmdSuivant_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdSuivant_Click"

   Resume Exit_CmdSuivant_Click
End Sub

Private Sub CmdInverseur_Click()

   On Error GoTo Err_CmdInverseur_Click

   Select Case iMonnaieVisu
      Case cEURO

        Inverseur cFRANC 'inverser la monnaie de visualisation

      Case cFRANC

        Inverseur cEURO 'inverser la monnaie de visualisation

    End Select

   TxtCli_Mont_Hon_m.Enabled = (Cli_MonnaieEnCours = iMonnaieVisu)

   CmdContratNouveau.Enabled = (Cli_MonnaieEnCours = iMonnaieVisu)

Exit_CmdInverseur_Click:

   Exit Sub

Err_CmdInverseur_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdInverseur_Click"

   Resume Exit_CmdInverseur_Click
End Sub

Private Sub CmdBasculer_Click()

   On Error GoTo Err_CmdBasculer_Click

   Select Case Cli_MonnaieEnCours
      Case cFRANC

         ChangerMonnaie cEURO

      Case cEURO

         ChangerMonnaie cEURO

   End Select

Exit_CmdBasculer_Click:

   Exit Sub

Err_CmdBasculer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdBasculer_Click"

   Resume Exit_CmdBasculer_Click
End Sub

Private Sub CmdTypeAdresse_Click()

   On Error GoTo Err_CmdTypeAdresse_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdTypeAdresse_Click

   End Select

   On Error GoTo Err_CmdTypeAdresse_Click

   DoCmd.OpenForm "FrmParam_TypeAdresse", acNormal, , , , acDialog

   SFrmFicheClientAdresse!CmbAC_TypeAdresse_ID.Requery

Exit_CmdTypeAdresse_Click:

   Exit Sub

Err_CmdTypeAdresse_Click:

   MsgBox Err.Number & "  " & Err.description, vbCritical, stSource & " : CmdTypeAdresse_Click"

   Resume Exit_CmdTypeAdresse_Click
End Sub

Private Sub CmdContratNouveau_Click()

   On Error GoTo Err_CmdContratNouveau_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdContratNouveau_Click

   End Select

   On Error GoTo Err_CmdContratNouveau_Click

   Select Case IsNull(Cli_Date_contrat)
      Case True

         MsgBox "Aucune date de contrat !", vbCritical

      Case False

         DoCmd.OpenForm "FrmFicheClientContratNouveau", acNormal, , , , acDialog, Cli_Code & ";" & iMonnaieVisu

         Me.Requery

   End Select

Exit_CmdContratNouveau_Click:

   Exit Sub

Err_CmdContratNouveau_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdContratNouveau_Click"

   Resume Exit_CmdContratNouveau_Click
End Sub

Private Sub CmdContratHistorique_Click()

   On Error GoTo Err_CmdContratHistorique_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdContratHistorique_Click

   End Select

   On Error GoTo Err_CmdContratHistorique_Click

   DoCmd.OpenForm "FrmFicheClientContratHistorique", acNormal, , , , acDialog, Cli_Code & ";" & iMonnaieVisu

Exit_CmdContratHistorique_Click:

   Exit Sub

Err_CmdContratHistorique_Click:

   MsgBox Err.Number & "  " & Err.description, vbCritical, stSource & " : CmdContratHistorique_Click"

   Resume Exit_CmdContratHistorique_Click
End Sub

Private Sub CmdTypeClient_Click()

   On Error GoTo Err_CmdTypeClient_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdTypeClient_Click

   End Select

   On Error GoTo Err_CmdTypeClient_Click

   DoCmd.OpenForm "FrmParam_TypeClient", acNormal, , , , acDialog

   CmbCli_Type.Requery

Exit_CmdTypeClient_Click:

   Exit Sub

Err_CmdTypeClient_Click:

   MsgBox Err.Number & "  " & Err.description, vbCritical, stSource & " : CmdTypeClient_Click"

   Resume Exit_CmdTypeClient_Click
End Sub

Private Sub CmdCG_Parent_Id_Click()

   On Error GoTo Err_CmdCG_Parent_Id_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdCG_Parent_Id_Click

   End Select

   On Error GoTo Err_CmdCG_Parent_Id_Click

   DoCmd.OpenForm "FrmParam_GroupeClient", acNormal, , , , acDialog

   CmbCG_Parent_Id.Requery

Exit_CmdCG_Parent_Id_Click:

   Exit Sub

Err_CmdCG_Parent_Id_Click:

   MsgBox Err.Number & "  " & Err.description, vbCritical, stSource & " : CmdCG_Parent_Id_Click"

   Resume Exit_CmdCG_Parent_Id_Click
End Sub

Private Sub CmdCli_Groupe_Click()

   On Error GoTo Err_CmdGroupe_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdGroupe_Click

   End Select

   On Error GoTo Err_CmdGroupe_Click

   DoCmd.OpenForm "FrmParam_SousGroupeClient", acNormal, , , , acDialog

   CmbCli_Groupe.Requery

Exit_CmdGroupe_Click:

   Exit Sub

Err_CmdGroupe_Click:

   MsgBox Err.Number & "  " & Err.description, vbCritical, stSource & " : CmdGroupe_Click"

   Resume Exit_CmdGroupe_Click
End Sub

Private Sub CmdClientListe_Click()

   On Error GoTo Err_CmdClientListe_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdClientListe_Click

   End Select

   On Error GoTo Err_CmdClientListe_Click

   DoCmd.OpenForm "FrmFicheClientListe", acNormal, , , , acDialog

   Select Case StrFicheClientListeCli_Code
      Case ""

      Case Else

         Me.Filter = "Cli_Code='" & StrFicheClientListeCli_Code & "'"

         Me.FilterOn = True

   End Select

Exit_CmdClientListe_Click:

   Exit Sub

Err_CmdClientListe_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdClientListe_Click"

   Resume Exit_CmdClientListe_Click
End Sub

Private Sub CmdClientNouveau_Click()

   On Error GoTo Err_CmdClientNouveau_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdClientNouveau_Click

   End Select

   On Error GoTo Err_CmdClientNouveau_Click

   DoCmd.OpenForm "FrmFicheClientNouveau", acNormal, , , , acWindowNormal

Exit_CmdClientNouveau_Click:

   Exit Sub

Err_CmdClientNouveau_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdClientNouveau_Click"

   Resume Exit_CmdClientNouveau_Click
End Sub

Private Sub CmdDebiteurListe_Click()

   On Error GoTo Err_CmdDebiteurListe_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdDebiteurListe_Click

   End Select

   On Error GoTo Err_CmdDebiteurListe_Click

   DoCmd.OpenForm "FrmFicheDebiteurListe", acNormal, , , , acDialog, "Déb_Code_Client='" & Cli_Code & "';" & iMonnaieVisu

   Select Case StrFicheDebiteurListeDéb_Code
      Case ""

      Case Else

         DoCmd.Close acForm, Me.Name, acSavePrompt

         DoCmd.OpenForm "FrmFicheDebiteur", acNormal, , "Déb_Code='" & StrFicheDebiteurListeDéb_Code & "'", , acWindowNormal

   End Select

Exit_CmdDebiteurListe_Click:

   Exit Sub

Err_CmdDebiteurListe_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdDebiteurListe_Click"

   Resume Exit_CmdDebiteurListe_Click
End Sub

Private Sub CmdDebiteurNouveau_Click()

   On Error GoTo Err_CmdDebiteurNouveau_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdDebiteurNouveau_Click

   End Select

   On Error GoTo Err_CmdDebiteurNouveau_Click

   DoCmd.OpenForm "FrmFicheDebiteurNouveau", acNormal, , , , acWindowNormal, Cli_Code & ";"

Exit_CmdDebiteurNouveau_Click:

   Exit Sub

Err_CmdDebiteurNouveau_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdDebiteurNouveau_Click"

   Resume Exit_CmdDebiteurNouveau_Click
End Sub
