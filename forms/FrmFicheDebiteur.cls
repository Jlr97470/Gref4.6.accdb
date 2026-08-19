Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes

Private iMonnaieVisu As Integer 'Monnaie dans laquelle est visualisée la fiche en cours

Public Sub Inverseur(ByVal IntMonnaie As Integer)
    'Cette fonction permet de visualiser une fiche dans une monnaie particulière (argMonnaie)
   On Error GoTo Err_Inverseur

   Select Case IntMonnaie
      Case cFRANC

         Select Case Déb_MonnaieEnCours <> IntMonnaie  'Ce n'est pas la monnaie du débiteur
            Case True

               EtiMonnaieVue1.Caption = "Vue en FRANC"

               EtiMonnaieVue2.Caption = "Vue en FRANC"

            Case False

               EtiMonnaieVue1.Caption = ""

               EtiMonnaieVue2.Caption = ""

         End Select

         TxtDéb_Mont_Prin_m.ControlSource = "Déb_Mont_Prin_m1"

         TxtDéb_Mont_Prin.ControlSource = "Déb_Mont_Prin_m1"

         TxtDéb_Frais_m.ControlSource = "Déb_Frais_m1"

         TxtTotalDu.ControlSource = "=Round(Déb_Mont_Prin_m1,2) + Round(Déb_Frais_m1,2)"

         TxtDéb_Prin_enc_m.ControlSource = "=Round(NZ(DSum(" & Chr(34) & "Enc_Prin_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtDéb_Frais_enc_m.ControlSource = "=Round(NZ(DSum(" & Chr(34) & "Enc_Frais_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtTotalEncaisse.ControlSource = "=Round(NZ(DSum(" & Chr(34) & "Enc_Prin_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2) + Round(NZ(DSum(" & Chr(34) & "Enc_Frais_m1" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtSoldeFrais.ControlSource = "=Round([Déb_Frais_m1],2)-[TxtDéb_Frais_enc_m]"

         TxtSoldePrincipal.ControlSource = "=Round([Déb_Mont_Prin_m1],2)-[TxtDéb_Prin_enc_m]"

      Case cEURO

         Select Case Déb_MonnaieEnCours <> IntMonnaie  'Ce n'est pas la monnaie du débiteur
            Case True

               EtiMonnaieVue1.Caption = "Vue en EURO"

               EtiMonnaieVue2.Caption = "Vue en EURO"

            Case False

               EtiMonnaieVue1.Caption = ""

               EtiMonnaieVue2.Caption = ""

         End Select

         TxtDéb_Mont_Prin_m.ControlSource = "Déb_Mont_Prin_m2"

         TxtDéb_Mont_Prin.ControlSource = "Déb_Mont_Prin_m2"

         TxtDéb_Frais_m.ControlSource = "Déb_Frais_m2"

         TxtTotalDu.ControlSource = "=Round(Déb_Mont_Prin_m2,2) + Round(Déb_Frais_m2,2)"

         TxtDéb_Prin_enc_m.ControlSource = "=Round(NZ(DSum(" & Chr(34) & "Enc_Prin_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtDéb_Frais_enc_m.ControlSource = "=Round(NZ(DSum(" & Chr(34) & "Enc_Frais_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtTotalEncaisse.ControlSource = "=Round(NZ(DSum(" & Chr(34) & "Enc_Prin_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2) + Round(NZ(DSum(" & Chr(34) & "Enc_Frais_m2" & Chr(34) & "," & Chr(34) & "ENCAISSEMENT" & Chr(34) & "," & Chr(34) & "Enc_Code_Débiteur='" & Chr(34) & " & [Déb_Code] & " & Chr(34) & "'" & Chr(34) & "),0),2)"

         TxtSoldeFrais.ControlSource = "=Round([Déb_Frais_m2],2)-[TxtDéb_Frais_enc_m]"

         TxtSoldePrincipal.ControlSource = "=Round([Déb_Mont_Prin_m2],2)-[TxtDéb_Prin_enc_m]"

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

   Déb_MonnaieEnCours = IntMonnaie

   Select Case IntMonnaie
       Case cFRANC

           CmdBasculer.Enabled = True

       Case cEURO

           TxtDéb_Nom.SetFocus 'permet de désélectionner le bouton "BasculerEuro avant de le désactiver"

           CmdBasculer.Enabled = False

   End Select

   Inverseur IntMonnaie

   SFrmFicheDebiteurEncaissement.Form.Inverseur IntMonnaie

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

   Inverseur Déb_MonnaieEnCours

   SFrmFicheDebiteurEncaissement.Form.Inverseur Déb_MonnaieEnCours

   TxtCode_Débiteur = Déb_Code

   CmdBasculer.Enabled = (Déb_MonnaieEnCours = cFRANC)

   Select Case Déb_MonnaieEnCours
      Case cFRANC

         CmdBasculer.Caption = "Basculer En " & DLookup("Mon_nom", "MONNAIE", "Mon_Id=" & cEURO)

      Case cEURO

         CmdBasculer.Caption = "Basculer En " & DLookup("Mon_nom", "MONNAIE", "Mon_Id=" & cFRANC)

   End Select

   CmdInverseur.Enabled = True

   CmdEcheance.Enabled = True

   CmdPrincipal.Enabled = True

   CmdEncaissement.Enabled = True

   TxtDéb_Frais_m.Enabled = True

   TxtDéb_Mont_Prin.Enabled = True

   CmbDéb_Repré_Type.Visible = Not IsNull(Déb_Date_Représentation)

   CmbDéb_niv_relance.Requery

Exit_Form_Current:

   Exit Sub

Err_Form_Current:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Current"

   Resume Exit_Form_Current
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)

   On Error GoTo Err_Form_AfterDelConfirm

   TxtCode_Débiteur = Null

   CmdInverseur.Enabled = False

   CmdEcheance.Enabled = False

   CmdPrincipal.Enabled = False

   CmdEncaissement.Enabled = False

Exit_Form_AfterDelConfirm:

   Exit Sub

Err_Form_AfterDelConfirm:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_AfterDelConfirm"

   Resume Exit_Form_AfterDelConfirm
End Sub

Private Sub TxtCode_Débiteur_AfterUpdate()

   On Error GoTo Err_TxtCode_Débiteur_AfterUpdate

   Me.Filter = "Déb_Code='" & TxtCode_Débiteur & "'"

   Me.FilterOn = True

   Select Case SFrmFicheDebiteurEncaissement.Visible
      Case True

         SFrmFicheDebiteurEncaissement.SetFocus

         SFrmFicheDebiteurEncaissement.Form.TxtEnc_Date.SetFocus

      Case False

         TxtDéb_Nom.SetFocus

   End Select

Exit_TxtCode_Débiteur_AfterUpdate:

   Exit Sub

Err_TxtCode_Débiteur_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtCode_Débiteur_AfterUpdate"

   Resume Exit_TxtCode_Débiteur_AfterUpdate
End Sub

Private Sub TxtCode_Débiteur_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_TxtCode_Débiteur_BeforeUpdate

   Select Case TxtCode_Débiteur
      Case DLookup("Déb_Code", "DEBITEUR", "Déb_Code='" & TxtCode_Débiteur & "'")

      Case DLookup("Déb_Code", "ARCHIVE_DEBITEUR", "Déb_Code='" & TxtCode_Débiteur & "'")

         Cancel = True

         MsgBox "Débiteur Archivé !", vbExclamation, stSource

      Case Else

         Cancel = True

         MsgBox "Identifiant incorrect !", vbExclamation, stSource

   End Select

Exit_TxtCode_Débiteur_BeforeUpdate:

   Exit Sub

Err_TxtCode_Débiteur_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtCode_Débiteur_BeforeUpdate"

   Resume Exit_TxtCode_Débiteur_BeforeUpdate
End Sub

Private Sub CmbDéb_Situation_AfterUpdate()
   Dim IntCode_Action As Integer
   Dim IntGrilleTauxNum As Integer
   Dim IntDelai As Integer

   On Error GoTo Err_CmbDéb_Situation_AfterUpdate

   IntCode_Action = Nz(DLookup("CG_Code_Action", "CODE_GESTION", "CG_Code='" & Déb_Situation & "'"), 0)

   Select Case IntCode_Action
      Case 1

         Déb_niv_relance = DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=" & IntCode_Action)

         Déb_Date_proch_rel = Null

      Case 2

         Déb_niv_relance = DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=" & IntCode_Action)

         IntGrilleTauxNum = DLookup("Cli_GrilleTaux", "CLIENT", "Cli_Code = '" & Déb_Code_Client & "'") '*** Numéro de la grille de taux du client

         IntDelai = DLookup("Tx_Délai", "TAUX", "Tx_N° = 9 AND Tx_Grille = " & IntGrilleTauxNum & " AND Tx_Code_Action = 2")

         Déb_Date_der_rel = Nz(Déb_Date_der_rel, Date)

         Déb_Date_proch_rel = Date + IntDelai

      Case 3, 5, 6

         CmbDéb_niv_relance = DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=" & IntCode_Action)

         Déb_Date_der_rel = Null

         Déb_Date_proch_rel = Null

      Case 4, 7, 8, 9

         Déb_niv_relance = Nz(DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=" & IntCode_Action & " AND Rel_N°=" & Nz(Déb_niv_relance, 0)), 1)

   End Select

   CmbDéb_niv_relance.Requery

   TxtDéb_Date_der_rel.Requery

   TxtDéb_Date_proch_rel.Requery

Exit_CmbDéb_Situation_AfterUpdate:

   Exit Sub

Err_CmbDéb_Situation_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbDéb_Situation_AfterUpdate"

   Resume Exit_CmbDéb_Situation_AfterUpdate
End Sub

Private Sub TxtDéb_Frais_m_AfterUpdate()

   On Error GoTo Err_TxtDéb_Frais_m_AfterUpdate
    'Mise à jour des valeurs de l'autre monnaie
    Select Case iMonnaieVisu
      Case cFRANC

         Déb_Frais_m2 = Déb_Frais_m1 * DonnerTauxConvertion(iMonnaieVisu)

      Case cEURO

         Déb_Frais_m1 = Déb_Frais_m2 * DonnerTauxConvertion(iMonnaieVisu)

    End Select

Exit_TxtDéb_Frais_m_AfterUpdate:

   Exit Sub

Err_TxtDéb_Frais_m_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDéb_Frais_m_AfterUpdate"

   Resume Exit_TxtDéb_Frais_m_AfterUpdate
End Sub

Private Sub TxtDéb_Typ_Cré_AfterUpdate()
   Dim StrChamp As String
   Dim StrTableau() As String
   Dim StrChaine As String
   Dim IntNumero As Integer

   On Error GoTo Err_TxtDéb_Typ_Cré_AfterUpdate

   Déb_Date_Chèque = Null

   Select Case IsNull(TxtDéb_Typ_Cré)
      Case True

      Case False

         StrChamp = TxtDéb_Typ_Cré

         StrTableau = Split(StrChamp, "/")

         Select Case UBound(StrTableau)
            Case Is >= 2

               Select Case IsDate(Right(StrTableau(0), 2) & "/" & Left(StrTableau(1), 2) & "/" & Left(StrTableau(2), 4))
                  Case True

                     Déb_Date_Chèque = Right(StrTableau(0), 2) & "/" & Left(StrTableau(1), 2) & "/" & Left(StrTableau(2), 4)

                     StrChamp = Replace(StrChamp, Right(StrTableau(0), 2) & "/" & Left(StrTableau(1), 2) & "/" & Left(StrTableau(2), 4), "")

                  Case False

                     Select Case IsDate(Right(StrTableau(0), 2) & "/" & Left(StrTableau(1), 2) & "/" & Left(StrTableau(2), 2))
                        Case True

                           Déb_Date_Chèque = Right(StrTableau(0), 2) & "/" & Left(StrTableau(1), 2) & "/" & Left(StrTableau(2), 2)

                           StrChamp = Replace(StrChamp, Right(StrTableau(0), 2) & "/" & Left(StrTableau(1), 2) & "/" & Left(StrTableau(2), 2), "")

                        Case False

                     End Select

               Case Else

            End Select

         End Select

         Déb_Nom_Banque = Null

         StrTableau = Split(StrChamp, " ")

         For IntNumero = 0 To UBound(StrTableau)

            StrChaine = StrTableau(IntNumero)

            Select Case IsNumeric(StrChaine)
               Case True

                  Déb_Num_Chèque = StrChaine

               Case False

                  Select Case StrChaine
                     Case "CHEQUE", "CHEQUES", "HOSPITALISATION", "FACTURE", "FACTURES", "DU", "CSP", "CHQ", "HOSP", "POUR", "SOLDE", "CONSOMMATION", "IMPAYEE", "IMPAYEES", "RADIO", "RADIOS", "URGENCE", "N°"

                     Case Else

                        Déb_Nom_Banque = IIf(Déb_Nom_Banque & IIf(StrChaine = "", "", " ") & StrChaine = "", Null, Déb_Nom_Banque & IIf(StrChaine = "", "", " ") & StrChaine)

                  End Select

            End Select

         Next

      End Select

Exit_TxtDéb_Typ_Cré_AfterUpdate:

   Exit Sub

Err_TxtDéb_Typ_Cré_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDéb_Typ_Cré_AfterUpdate"

   Resume Exit_TxtDéb_Typ_Cré_AfterUpdate
End Sub

Private Sub TxtDéb_Date_Représentation_AfterUpdate()
   Dim IntCode_Représentation As Integer
   Dim IntRepMsg As Integer

   Select Case IsNull(Déb_Date_Représentation)
      Case True

      Case False

         IntCode_Représentation = Nz(DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action=10"))

         Select Case IntCode_Représentation
            Case 0

               MsgBox "ATTENTION !" & Chr(10) & "Le parametres code représentation est abscent dans les parametres GREF.", vbExclamation, stSource

               GoTo Exit_TxtDéb_Date_Représentation_AfterUpdate

            Case Else

         End Select

         Select Case Déb_Situation
            Case IntCode_Représentation

            Case Else

               IntRepMsg = MsgBox("Voulez-vous mettre le dossier en situation " & IntCode_Représentation & "?", 36)

               Select Case IntRepMsg
                  Case 6

                     CmbDéb_Situation = IntCode_Représentation

                  Case Else

               End Select

         End Select

    End Select

   CmbDéb_Repré_Type.Visible = Not IsNull(Déb_Date_Représentation)

Exit_TxtDéb_Date_Représentation_AfterUpdate:

   Exit Sub

Err_TxtDéb_Date_Représentation_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDéb_Date_Représentation_AfterUpdate"

   Resume Exit_TxtDéb_Date_Représentation_AfterUpdate
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

Private Sub CmdVisualiser_Click()

   On Error GoTo Err_CmdVisualiser_Click

   DoCmd.OpenReport "EtaAttestationPaiement", acViewPreview, , "Déb_Code='" & Déb_Code & "'"

Exit_CmdVisualiser_Click:

   Exit Sub

Err_CmdVisualiser_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdVisualiser_Click"

   Resume Exit_CmdVisualiser_Click
End Sub

Private Sub CmdFermer_Click()
   Dim IntEtat As Integer
   Dim StrDéb_Code As String

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

   Select Case DonneEtatFormulaire("FrmFicheDebiteurListeRechercher", IntEtat)
      Case True

         DoCmd.Close acForm, "FrmFicheDebiteurListeRechercher", acSavePrompt

      Case False

   End Select

   Select Case DonneEtatFormulaire("FrmDebiteurPhoningListe", IntEtat)
      Case True

         DoCmd.OpenForm "FrmDebiteurPhoningListe"

         StrDéb_Code = Forms!FrmDebiteurPhoningListe.Recordset("Déb_Code")

         Forms!FrmDebiteurPhoningListe.Form.Requery

         Forms!FrmDebiteurPhoningListe.Recordset.FindFirst "Déb_Code>='" & StrDéb_Code & "'"

      Case False

   End Select

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdPrecedent_Click()
   Dim StrDéb_Code As String

   On Error GoTo Err_CmdPrecedent_Click

   StrDéb_Code = Nz(TxtCode_Débiteur)

   Me.Filter = ""

   Me.FilterOn = True

   Me.Recordset.FindLast "Déb_Code<'" & StrDéb_Code & "'"

   Select Case StrDéb_Code
      Case ""

      Case Else

         Me.Filter = "Déb_Code='" & TxtCode_Débiteur & "'"

         Me.FilterOn = True

         Select Case SFrmFicheDebiteurEncaissement.Visible
            Case True

               SFrmFicheDebiteurEncaissement.SetFocus

               SFrmFicheDebiteurEncaissement.Form.TxtEnc_Date.SetFocus

            Case False

               TxtDéb_Nom.SetFocus

         End Select

   End Select

Exit_CmdPrecedent_Click:

   Exit Sub

Err_CmdPrecedent_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdPrecedent_Click"

   Resume Exit_CmdPrecedent_Click
End Sub

Private Sub CmdSuivant_Click()
   Dim StrDéb_Code As String

   On Error GoTo Err_CmdSuivant_Click

'   StrDéb_Code = Nz(DMin("Déb_Code", "DEBITEUR", "Déb_Code>'" & TxtCode_Débiteur & "'"))

   StrDéb_Code = Nz(TxtCode_Débiteur)

   Me.Filter = ""

   Me.FilterOn = True

   Me.Recordset.FindFirst "Déb_Code>'" & StrDéb_Code & "'"

   Select Case StrDéb_Code
      Case ""

      Case Else

         Me.Filter = "Déb_Code='" & TxtCode_Débiteur & "'"

         Me.FilterOn = True

         Select Case SFrmFicheDebiteurEncaissement.Visible
            Case True

               SFrmFicheDebiteurEncaissement.SetFocus

               SFrmFicheDebiteurEncaissement.Form.TxtEnc_Date.SetFocus

            Case False

               TxtDéb_Nom.SetFocus

         End Select

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

         SFrmFicheDebiteurEncaissement.Form.Inverseur cFRANC

      Case cFRANC

        Inverseur cEURO 'inverser la monnaie de visualisation

        SFrmFicheDebiteurEncaissement.Form.Inverseur cEURO

    End Select

   TxtDéb_Frais_m.Enabled = (Déb_MonnaieEnCours = iMonnaieVisu)

   TxtDéb_Mont_Prin.Enabled = (Déb_MonnaieEnCours = iMonnaieVisu)

   CmdEcheance.Enabled = (Déb_MonnaieEnCours = iMonnaieVisu)

Exit_CmdInverseur_Click:

   Exit Sub

Err_CmdInverseur_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdInverseur_Click"

   Resume Exit_CmdInverseur_Click
End Sub

Private Sub CmdBasculer_Click()

   On Error GoTo Err_CmdBasculer_Click

   Select Case Déb_MonnaieEnCours
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

   SFrmFicheDebiteurAdresse!CmbAD_TypeAdresse_ID.Requery

Exit_CmdTypeAdresse_Click:

   Exit Sub

Err_CmdTypeAdresse_Click:

   MsgBox Err.Number & "  " & Err.description, vbCritical, stSource & " : CmdTypeAdresse_Click"

   Resume Exit_CmdTypeAdresse_Click
End Sub

Private Sub CmdFicheClient_Click()

   On Error GoTo Err_CmdFicheClient_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdFicheClient_Click

   End Select

   On Error GoTo Err_CmdFicheClient_Click

   DoCmd.OpenForm "FrmFicheClient", acNormal, , "Cli_Code='" & Déb_Code_Client & "'", acFormReadOnly, acWindowNormal

   Forms!FrmFicheClient.CmdClientListe.Visible = False

   Forms!FrmFicheClient.CmdClientNouveau.Visible = False

   Forms!FrmFicheClient.CmdDebiteurListe.Visible = False

   Forms!FrmFicheClient.CmdDebiteurNouveau.Visible = False

   Forms!FrmFicheClient.RecCode_Client.Visible = False

   Forms!FrmFicheClient.TxtCode_Client.Visible = False

   Forms!FrmFicheClient.CmdSuivant.Visible = False

   Forms!FrmFicheClient.CmdPrecedent.Visible = False

Exit_CmdFicheClient_Click:

    Exit Sub

Err_CmdFicheClient_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFicheClient_Click"

    Resume Exit_CmdFicheClient_Click
End Sub

Private Sub CmdDebiteurDetail_Click()

   On Error GoTo Err_CmdDebiteurDetail_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdDebiteurDetail_Click

   End Select

   On Error GoTo Err_CmdDebiteurDetail_Click

   DoCmd.OpenForm "FrmFicheDebiteurDetail", acNormal, , , , acDialog, Déb_Code & ";" & iMonnaieVisu

Exit_CmdDebiteurDetail_Click:

   Exit Sub

Err_CmdDebiteurDetail_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdDebiteurDetail_Click"

   Resume Exit_CmdDebiteurDetail_Click
End Sub

Private Sub CmdSolde_Click()

   On Error GoTo Err_CmdSolde_Click

   EtiSoldePrincipal.Visible = Not EtiSoldePrincipal.Visible

   EtiSoldeFrais.Visible = Not EtiSoldeFrais.Visible

   TxtSoldePrincipal.Visible = Not TxtSoldePrincipal.Visible

   TxtSoldeFrais.Visible = Not TxtSoldeFrais.Visible

   TxtSolde.Visible = Not TxtSolde.Visible

Exit_CmdSolde_Click:

   Exit Sub

Err_CmdSolde_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdSolde_Click"

   Resume Exit_CmdSolde_Click
End Sub

Private Sub CmdEncaissement_Click()

   On Error GoTo Err_CmdEncaissement_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdEncaissement_Click

   End Select

   On Error GoTo Err_CmdEncaissement_Click

   Select Case TxtSolde
     Case Is <= 0

       MsgBox "Débiteur Soldé !", vbExclamation

     Case Else

   End Select
   'Méssage si des atermoiements sont en cours
   Select Case IsNull(DLookup("De_Code_Débiteur", "DEBITEUR_ECHEANCE", "De_Code_Débiteur = '" & Déb_Code & "'"))
     Case True

     Case False

       MsgBox "Echéance CB en cours...", vbInformation, stSource

   End Select

   EtiSFrmFicheDebiteurEncaissement.Visible = True

   RecSFrmFicheDebiteurEncaissement.Visible = True

   SFrmFicheDebiteurEncaissement.Visible = True

   SFrmFicheDebiteurEncaissement.SetFocus

   SFrmFicheDebiteurEncaissement.Form.TxtEnc_Date.SetFocus

   TxtDéb_Date_Cré.Visible = False

   TxtDéb_Mont_Prin.Visible = False

   TxtDéb_Typ_Cré.Visible = False

   TxtDéb_Date_Représentation.Visible = False

   CmbDéb_Repré_Type.Visible = False

   CmdPrincipal.Visible = True

   CmdEncaissement.Visible = False

Exit_CmdEncaissement_Click:

   Exit Sub

Err_CmdEncaissement_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdEncaissement_Click"

   Resume Exit_CmdEncaissement_Click
End Sub

Private Sub CmdPrincipal_Click()

   On Error GoTo Err_CmdPrincipal_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdPrincipal_Click

   End Select

   On Error GoTo Err_CmdPrincipal_Click

   TxtDéb_Date_Cré.Visible = True

   TxtDéb_Mont_Prin.Visible = True

   TxtDéb_Typ_Cré.Visible = True

   TxtDéb_Date_Représentation.Visible = True

   CmbDéb_Repré_Type.Visible = Not IsNull(Déb_Date_Représentation)

   TxtDéb_Date_Cré.SetFocus

   EtiSFrmFicheDebiteurEncaissement.Visible = False

   RecSFrmFicheDebiteurEncaissement.Visible = False

   SFrmFicheDebiteurEncaissement.Visible = False

   CmdPrincipal.Visible = False

   CmdEncaissement.Visible = True

Exit_CmdPrincipal_Click:

    Exit Sub

Err_CmdPrincipal_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdPrincipal_Click"

    Resume Exit_CmdPrincipal_Click
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

   DoCmd.OpenForm "FrmFicheDebiteurListeRechercher", acNormal, , , , acWindowNormal, Me.Name

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

   DoCmd.OpenForm "FrmFicheDebiteurNouveau", acNormal, , , , acWindowNormal

Exit_CmdDebiteurNouveau_Click:

   Exit Sub

Err_CmdDebiteurNouveau_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdDebiteurNouveau_Click"

   Resume Exit_CmdDebiteurNouveau_Click
End Sub

Private Sub CmdEcheance_Click()
   Dim IntMsg As Integer

   On Error GoTo Err_CmdEcheance_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdEcheance_Click

   End Select

   On Error GoTo Err_CmdEcheance_Click

   DoCmd.OpenForm "FrmFicheDebiteurEcheance", acNormal, , , , acDialog, Déb_Code & ";" & iMonnaieVisu

   Select Case DCount("De_Date_éch", "DEBITEUR_ECHEANCE", "De_Code_Débiteur='" & Déb_Code & "'")
      Case 0

            Select Case CmbDéb_Situation
               Case DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action=" & 4)

                  CmbDéb_Situation = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action=" & 1)

                  CmbDéb_niv_relance = DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=" & 1 & " AND Rel_N°=" & Nz(Déb_niv_relance, 1))

                  CmbDéb_niv_relance.Requery

               Case Else

            End Select

      Case Else

         Select Case CmbDéb_Situation
            Case DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action=" & 4)

            Case Else

                IntMsg = MsgBox("Basculer le débiteur en situation d'écheancier (" & DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action=" & 4) & ") ?", 36)

                Select Case IntMsg
                  Case 6

                     CmbDéb_Situation = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action=" & 4)

                     CmbDéb_niv_relance = DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=" & 4 & " AND Rel_N°=" & Nz(Déb_niv_relance, 1))

                     CmbDéb_niv_relance.Requery

                  Case Else

               End Select

         End Select

   End Select

Exit_CmdEcheance_Click:

   Exit Sub

Err_CmdEcheance_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdEcheance_Click"

   Resume Exit_CmdEcheance_Click
End Sub
