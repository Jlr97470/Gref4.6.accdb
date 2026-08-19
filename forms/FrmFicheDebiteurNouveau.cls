Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private StrCli_Code  As String
Private iMonnaieVisu As Integer

Public Sub Inverseur(ByVal IntMonnaie As Integer)
   'Cette fonction permet de visualiser une fiche dans une monnaie particulière (argMonnaie)
   On Error GoTo Err_Inverseur

   Select Case IntMonnaie
      Case cFRANC

         TxtDéb_Mont_Prin_m.ControlSource = "Déb_Mont_Prin_m1"

      Case cEURO

         TxtDéb_Mont_Prin_m.ControlSource = "Déb_Mont_Prin_m2"

   End Select

   iMonnaieVisu = IntMonnaie

   CmbDéb_Mont_Prin_mo = IntMonnaie

Exit_Inverseur:

   Exit Sub

Err_Inverseur:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Inverseur"

    Resume Exit_Inverseur
End Sub

Public Sub AfficherCodeDebiteur(StrCode_Client As String)
   Dim StrCodeDebiteur As String

   Select Case StrCode_Client
      Case ""

         TxtDéb_Code.Enabled = False

         CmbDéb_Qualité.Enabled = False

         TxtDéb_Nom.Enabled = False

         TxtDéb_Typ_Cré.Enabled = False

         TxtDéb_Mont_Prin_m.Enabled = False

         CmbDéb_Mont_Prin_mo.Enabled = False

         CmdOK.Enabled = False

         DoCmd.RunCommand acCmdUndo

      Case Else

         StrCodeDebiteur = DonnerNouveauCodeDebiteur(StrCode_Client)

         TxtDéb_Code.Enabled = True

         CmbDéb_Qualité.Enabled = True

         TxtDéb_Nom.Enabled = True

         TxtDéb_Typ_Cré.Enabled = True

         TxtDéb_Mont_Prin_m.Enabled = True

         CmbDéb_Mont_Prin_mo.Enabled = True

         CmdOK.Enabled = True

         TxtDéb_Code = StrCodeDebiteur

         TxtDéb_Code_Client = StrCode_Client

         CmbDéb_Qualité.SetFocus

   End Select

Exit_AfficherCodeDebiteur:

   Exit Sub

Err_AfficherCodeDebiteur:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : AfficherCodeDebiteur"

   Resume Exit_AfficherCodeDebiteur
End Sub

Private Sub Form_Open(Cancel As Integer)
   Dim StrParam() As String

   On Error GoTo Err_Form_Open

   StrParam = Split(Nz(Me.OpenArgs), ";")

   Select Case UBound(StrParam)
      Case Is < 1

         GoTo Exit_Form_Open

      Case Else

   End Select

   StrCli_Code = StrParam(0)

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   Inverseur cEURO

   Select Case StrCli_Code
      Case ""

      Case Else

         AfficherCodeDebiteur StrCli_Code

         CmbDéb_Qualité.SetFocus

   End Select

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Load"

   Resume Exit_Form_Load
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_Form_BeforeUpdate

   Select Case IsNull(TxtDéb_Nom)
      Case True

         Cancel = True

         MsgBox "Le nom du débiteur est obligatoire.", vbInformation, stSource

         TxtDéb_Nom.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case False

   End Select

   Select Case IsNull(Me.TxtDéb_Typ_Cré)
      Case True

         Cancel = True

         MsgBox "L'objet de la créance est obligatoire.", vbInformation, stSource

         TxtDéb_Typ_Cré.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case False

   End Select

   Select Case IsNull(TxtDéb_Mont_Prin_m) Or TxtDéb_Mont_Prin_m <= 0
      Case True

         Cancel = True

         MsgBox "Le montant de la créance est obligatoire." & vbCrLf & _
                      "(Choisissez d'abord la monnaie de la créance.)", vbInformation, stSource

         TxtDéb_Mont_Prin_m.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case False

   End Select

   Select Case IsNull(Déb_Mont_Prin_mo)
      Case True

         Cancel = True

         MsgBox "Monnaie incorrecte.", vbInformation, stSource

         CmbDéb_Mont_Prin_mo.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case False

   End Select

   Déb_Date_Cré = Date

   Déb_MonnaieEnCours = Déb_Mont_Prin_mo

   Select Case Déb_MonnaieEnCours
      Case cFRANC

         Déb_Mont_Prin_m2 = Déb_Mont_Prin_m1 * DonnerTauxConvertion(Déb_MonnaieEnCours)

         Déb_Frais_m1 = DonnerMontantFrais(Déb_Code_Client, Déb_Mont_Prin_m1, Déb_MonnaieEnCours)

         Déb_Frais_m2 = Déb_Frais_m1 * DonnerTauxConvertion(Déb_MonnaieEnCours)

      Case cEURO

         Déb_Mont_Prin_m1 = Déb_Mont_Prin_m2 * DonnerTauxConvertion(Déb_MonnaieEnCours)

         Déb_Frais_m2 = DonnerMontantFrais(Déb_Code_Client, Déb_Mont_Prin_m2, Déb_MonnaieEnCours)

         Déb_Frais_m1 = Déb_Frais_m2 * DonnerTauxConvertion(Déb_MonnaieEnCours)

   End Select

   Déb_Situation = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action=1")

   Déb_niv_relance = DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=1")

Exit_Form_BeforeUpdate:

   Exit Sub

Err_Form_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_BeforeUpdate"

   Resume Exit_Form_BeforeUpdate
End Sub

Private Sub Form_AfterUpdate()

   On Error GoTo Err_Form_AfterUpdate

   CurrentDb.Execute "UPDATE CLIENT SET CLIENT.Cli_Code_der_déb = '" & Déb_Code & "', CLIENT.Cli_Date_der_cré = " & ConvDate(Date) & _
                               "WHERE CLIENT.Cli_Code='" & Déb_Code_Client & "';", dbFailOnError

Exit_Form_AfterUpdate:

   Exit Sub

Err_Form_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_AfterUpdate"

   Resume Exit_Form_AfterUpdate
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

   AfficherCodeDebiteur Nz(TxtDéb_Code_Client)

Exit_TxtDéb_Code_Client_AfterUpdate:

   Exit Sub

Err_TxtDéb_Code_Client_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDéb_Code_Client_AfterUpdate"

   Resume Exit_TxtDéb_Code_Client_AfterUpdate
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

Private Sub CmbDéb_Mont_Prin_mo_AfterUpdate()

   On Error GoTo Err_CmbDéb_Mont_Prin_mo_AfterUpdate

   Inverseur CmbDéb_Mont_Prin_mo

Exit_CmbDéb_Mont_Prin_mo_AfterUpdate:

   Exit Sub

Err_CmbDéb_Mont_Prin_mo_AfterUpdate:

   MsgBox Err.Number & " " & Err.description

   Resume Err_CmbDéb_Mont_Prin_mo_AfterUpdate
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdUndo

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

Private Sub CmdClientListe_Click()

   On Error GoTo Err_CmdClientListe_Click

   DoCmd.OpenForm "FrmFicheClientListe", acNormal, , , acFormReadOnly, acDialog

   AfficherCodeDebiteur StrFicheClientListeCli_Code

Exit_CmdClientListe_Click:

   Exit Sub

Err_CmdClientListe_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdClientListe_Click"

   Resume Exit_CmdClientListe_Click
End Sub

Private Sub CmdOK_Click()
   Dim IntEtat As Integer
   Dim IntTypeAdrCourrier As Integer
   Dim CtlControl As Control

   On Error GoTo Err_CmdOK_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 3021, 2046

         Err.Clear

         GoTo Exit_CmdOK_Click

      Case Else

         GoTo Err_CmdOK_Click

   End Select

   Select Case Me.NewRecord
      Case True

         Form_BeforeUpdate False

         GoTo Exit_CmdOK_Click

      Case False

   End Select

   On Error GoTo Err_CmdOK_Click

   Select Case DonneEtatFormulaire("FrmFicheClient", IntEtat)
      Case True

         DoCmd.Close acForm, "FrmFicheClient", acSavePrompt

      Case False

   End Select

   Select Case DonneEtatFormulaire("FrmFicheDebiteur", IntEtat)
      Case True

         Set CtlControl = Forms!FrmFicheDebiteur.Form.ActiveControl

         Forms!FrmFicheDebiteur.Filter = "Déb_Code='" & Déb_Code & "'"

         Forms!FrmFicheDebiteur.FilterOn = True

      Case False

         DoCmd.OpenForm "FrmFicheDebiteur", acNormal, , "Déb_Code='" & Déb_Code & "'", , acWindowNormal

         Set CtlControl = Forms!FrmFicheDebiteur.Form.ActiveControl

   End Select

   CtlControl.SetFocus

   IntTypeAdrCourrier = Nz(DLookup("Adresse_Courrier", "PARAMETRE_GREF"))

   Select Case IntTypeAdrCourrier
      Case 0

         MsgBox "ATTENTION !" & Chr(10) & "Le parametres type d'adresse par defaut est abscent dans les parametres GREF.", vbExclamation, stSource

         GoTo Exit_CmdOK_Click

      Case Else

   End Select

   DoCmd.OpenForm "FrmFicheDebiteurDetailAdresse", acNormal, , , acFormAdd, acWindowNormal, Déb_Code & ";" & IntTypeAdrCourrier

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdOK_Click:

    Exit Sub

Err_CmdOK_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

    Resume Exit_CmdOK_Click
End Sub
