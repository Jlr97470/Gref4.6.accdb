Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private iMonnaieVisu As Integer

Public Sub Inverseur(IntMonnaie As Integer)
    'Cette fonction permet de visualiser le formulaire dans une monnaie particulière (argMonnaie)
   On Error GoTo Err_Inverseur

   Select Case IntMonnaie
       Case cFRANC

           TxtEnc_Montant_m.ControlSource = "Enc_Montant_m1"

           TxtEnc_Prin_m.ControlSource = "Enc_Prin_m1"

           TxtEnc_Frais_m.ControlSource = "Enc_Frais_m1"

       Case cEURO

           TxtEnc_Montant_m.ControlSource = "Enc_Montant_m2"

           TxtEnc_Prin_m.ControlSource = "Enc_Prin_m2"

           TxtEnc_Frais_m.ControlSource = "Enc_Frais_m2"

   End Select

   Me.Filter = "Enc_Montant_mo=" & IntMonnaie

   Me.FilterOn = True

   Me.AllowAdditions = Forms!FrmFicheDebiteur.Déb_MonnaieEnCours = IntMonnaie 'On empèche la création d'encaissements

   Me.AllowEdits = Forms!FrmFicheDebiteur.Déb_MonnaieEnCours = IntMonnaie 'On empèche la modification d'encaissements

   On Error Resume Next

   CmdEncaissementDetail.SetFocus

   Select Case Err.Number
      Case 0

      Case 2105

         Err.Clear

      Case Else

         GoTo Err_Inverseur

   End Select

   On Error GoTo Err_Inverseur

   TxtEnc_Montant_m.Enabled = Forms!FrmFicheDebiteur.Déb_MonnaieEnCours = IntMonnaie

   TxtEnc_Prin_m.Enabled = Forms!FrmFicheDebiteur.Déb_MonnaieEnCours = IntMonnaie

   iMonnaieVisu = IntMonnaie

Exit_Inverseur:

    Exit Sub

Err_Inverseur:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource

   Resume Exit_Inverseur
End Sub

Private Function VerifierSolde()
   Dim StrCode_Solder As String 'Code de Gestion correspondant au code Action 5 (Soldé)
   Dim StrCode_Direct As String
   Dim StrCode_Situation As String  'Code de Gestion correspondant à la situation du débiteur
   Dim CurMontantPrincipal As Currency
   Dim CurFrais As Currency
   Dim CurEncaisserPrincipal As Currency
   Dim IntRepMsg As Integer
   Dim IntModeEnc As Integer 'mode de l'encaissement : CLIENT ou GREF

   On Error GoTo Err_VerifierSolde

   StrCode_Solder = Nz(DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action = 8"))

   Select Case StrCode_Solder
      Case ""

         MsgBox "Attention, aucun des codes de gestion ne correspond au code action Soldé !"

         GoTo Exit_VerifierSolde

      Case Else

   End Select

   StrCode_Direct = Nz(DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action = 7"))

   Select Case StrCode_Direct
      Case ""

         MsgBox "Attention, aucun des codes de gestion ne correspond au code action Encaissement Direct !"

         GoTo Exit_VerifierSolde

      Case Else

   End Select

   StrCode_Situation = Forms!FrmFicheDebiteur.Déb_Situation

   Select Case iMonnaieVisu
     Case cFRANC

        CurMontantPrincipal = Round(Forms!FrmFicheDebiteur.Déb_Mont_Prin_m1, 2)

        CurFrais = Round(Forms!FrmFicheDebiteur.Déb_Frais_m1, 2)

        CurEncaisserPrincipal = Round(DSum("Enc_Prin_m1", "ENCAISSEMENT", "Enc_Code_Débiteur = '" & Enc_Code_Débiteur & "'"), 2)

     Case cEURO

        CurMontantPrincipal = Round(Forms!FrmFicheDebiteur.Déb_Mont_Prin_m2, 2)

        CurFrais = Round(Forms!FrmFicheDebiteur.Déb_Frais_m2, 2)

        CurEncaisserPrincipal = Round(DSum("Enc_Prin_m2", "ENCAISSEMENT", "Enc_Code_Débiteur = '" & Enc_Code_Débiteur & "'"), 2)

   End Select

   Select Case CurMontantPrincipal
      Case Is <= CurEncaisserPrincipal

         IntModeEnc = DLookup("ETy_Mode", "ENCAISSEMENT_TYPE", "ETy_Code='" & Enc_Type & "'")
        'Encaissement chez le client = 2 Si un seul encaissement client   ou Si plusieurs encaissement client et aucun encaissement gref   ou Si encaissement client + encaissement Gref ne payant que les frais
         Select Case IntModeEnc = 2 And (DCount("Enc_Id", "ENCAISSEMENT", "Enc_Code_Débiteur = '" & Enc_Code_Débiteur & "'") = 1 Or DCount("Enc_Id", "SelEncaissementTypeMode", "Enc_Code_Débiteur = '" & Enc_Code_Débiteur & "' AND ETy_Mode = 1 ") = 0 Or DSum("Enc_Prin_m1", "SelEncaissementTypeMode", "Enc_Code_Débiteur = '" & Enc_Code_Débiteur & "' AND ETy_Mode = 1") = 0)
            Case True

               Select Case StrCode_Situation
                  Case StrCode_Direct

                  Case Else

                     IntRepMsg = MsgBox("Cette encaissement direct solde le principal de la créance. Voulez-vous modifier le code situation du dossier en conséquence?", 36)

                     Select Case IntRepMsg
                        Case 6

                           Forms!FrmFicheDebiteur.CmbDéb_Situation = StrCode_Direct

                           Forms!FrmFicheDebiteur.CmbDéb_Situation.Requery

                           Forms!FrmFicheDebiteur.CmbDéb_niv_relance.Requery

                        Case Else

                     End Select

               End Select

            Case False

               Select Case StrCode_Situation
                  Case StrCode_Solder

                  Case Else

                     IntRepMsg = MsgBox("Voulez-vous solder le débiteur ?", 36)

                     Select Case IntRepMsg
                        Case 6

                           Forms!FrmFicheDebiteur.CmbDéb_Situation = StrCode_Solder

                           Forms!FrmFicheDebiteur.CmbDéb_Situation.Requery

                           Forms!FrmFicheDebiteur.CmbDéb_niv_relance.Requery

                        Case Else

                   End Select

               End Select

         End Select

      Case Else

   End Select

Exit_VerifierSolde:

    Exit Function

Err_VerifierSolde:

    MsgBox Err.Number & " " & Err.description

    Resume Exit_VerifierSolde
End Function

Private Sub Form_BeforeUpdate(Cancel As Integer)
   Dim LngEnc_ID As Long

   On Error GoTo Err_Form_BeforeUpdate

   Enc_Montant_mo = iMonnaieVisu

   LngEnc_ID = Nz(DMax("Enc_Id", "ENCAISSEMENT"), 0) + 1

   Select Case LngEnc_ID
      Case Is > Me.Enc_ID

         Me.Enc_ID = LngEnc_ID

      Case Else

   End Select

Exit_Form_BeforeUpdate:

   Exit Sub

Err_Form_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_Form_BeforeUpdate
End Sub

Private Sub Form_AfterUpdate()

   On Error GoTo Err_Form_AfterUpdate

   VerifierSolde

Exit_Form_AfterUpdate:

   Exit Sub

Err_Form_AfterUpdate:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_Form_AfterUpdate
End Sub

Private Sub Form_Delete(Cancel As Integer)

   On Error GoTo Err_Form_Delete

   Select Case Enc_Situation
      Case "R"

        Cancel = True

        MsgBox "Impossible de modifier cet encaissement." & vbCrLf & "Le paiement a déjà été réglé au client.", vbExclamation, stSource

      Case "F"

        Cancel = True

        MsgBox "Impossible de modifier cet encaissement." & vbCrLf & "Le paiement a déjà été facturé au client.", vbExclamation, stSource

      Case Else

   End Select

Exit_Form_Delete:

   Exit Sub

Err_Form_Delete:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_Form_Delete
End Sub

Private Sub TxtEnc_Date_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_TxtEnc_Date_BeforeUpdate

   Select Case Enc_Situation
      Case "R"

        Cancel = True

        MsgBox "Impossible de modifier cet encaissement." & vbCrLf & "Le paiement a déjà été réglé au client.", vbExclamation, stSource

      Case "F"

        Cancel = True

        MsgBox "Impossible de modifier cet encaissement." & vbCrLf & "Le paiement a déjà été facturé au client.", vbExclamation, stSource

      Case Else

   End Select

   Select Case IsNull(Enc_Date)
      Case True

         Cancel = True

         MsgBox "La date d'encaissement est obligatoire"

         TxtEnc_Date.SetFocus

         GoTo Exit_TxtEnc_Date_BeforeUpdate

      Case False

   End Select

Exit_TxtEnc_Date_BeforeUpdate:

   Exit Sub

Err_TxtEnc_Date_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_TxtEnc_Date_BeforeUpdate
End Sub

Private Sub CmbEnc_Type_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_CmbEnc_Type_BeforeUpdate

   Select Case Enc_Situation
      Case "R"

        Cancel = True

        MsgBox "Impossible de modifier cet encaissement." & vbCrLf & "Le paiement a déjà été réglé au client.", vbExclamation, stSource

      Case "F"

        Cancel = True

        MsgBox "Impossible de modifier cet encaissement." & vbCrLf & "Le paiement a déjà été facturé au client.", vbExclamation, stSource

      Case Else

   End Select

   Select Case IsNull(Enc_Type)
      Case True

         Cancel = True

         MsgBox "Le type d'encaissement est obligatoire"

         TxtEnc_Date.SetFocus

         GoTo Exit_CmbEnc_Type_BeforeUpdate

      Case False

   End Select

Exit_CmbEnc_Type_BeforeUpdate:

   Exit Sub

Err_CmbEnc_Type_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_CmbEnc_Type_BeforeUpdate
End Sub

Private Sub TxtEnc_Montant_m_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_TxtEnc_Montant_m_BeforeUpdate

   Select Case Enc_Situation
      Case "R"

        Cancel = True

        MsgBox "Impossible de modifier cet encaissement." & vbCrLf & "Le paiement a déjà été réglé au client.", vbExclamation, stSource

      Case "F"

        Cancel = True

        MsgBox "Impossible de modifier cet encaissement." & vbCrLf & "Le paiement a déjà été facturé au client.", vbExclamation, stSource

      Case Else

   End Select

   Select Case IsNull(TxtEnc_Montant_m)
      Case True

         Cancel = True

         MsgBox "Le montant de l'encaissement est obligatoire"

         TxtEnc_Date.SetFocus

         GoTo Exit_TxtEnc_Montant_m_BeforeUpdate

      Case False

   End Select

Exit_TxtEnc_Montant_m_BeforeUpdate:

   Exit Sub

Err_TxtEnc_Montant_m_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_TxtEnc_Montant_m_BeforeUpdate
End Sub

Private Sub TxtEnc_Prin_m_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_TxtEnc_Prin_m_BeforeUpdate

   Select Case Enc_Situation
      Case "R"

        Cancel = True

        MsgBox "Impossible de modifier cet encaissement." & vbCrLf & "Le paiement a déjà été réglé au client.", vbExclamation, stSource

      Case "F"

        Cancel = True

        MsgBox "Impossible de modifier cet encaissement." & vbCrLf & "Le paiement a déjà été facturé au client.", vbExclamation, stSource

      Case Else

   End Select

   Select Case IsNull(TxtEnc_Prin_m)
      Case True

         Cancel = True

         MsgBox "Le principal de l'encaissement est obligatoire"

         TxtEnc_Date.SetFocus

         GoTo Exit_TxtEnc_Prin_m_BeforeUpdate

      Case False

   End Select

Exit_TxtEnc_Prin_m_BeforeUpdate:

   Exit Sub

Err_TxtEnc_Prin_m_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_TxtEnc_Prin_m_BeforeUpdate
End Sub

Private Sub TxtEnc_Montant_m_AfterUpdate()
   Dim CurMontantPrincipal As Currency
   Dim CurFrais As Currency
   Dim CurEncaisserPrincipal As Currency

   On Error GoTo Err_TxtEnc_Montant_m_AfterUpdate

   Select Case iMonnaieVisu
     Case cFRANC

        Enc_Prin_m1 = 0

        Enc_Frais_m1 = 0

        CurMontantPrincipal = Round(Forms!FrmFicheDebiteur.Déb_Mont_Prin_m1, 2)

        CurFrais = Round(Forms!FrmFicheDebiteur.Déb_Frais_m1, 2)

        CurEncaisserPrincipal = Round(Nz(DSum("Enc_Prin_m1", "ENCAISSEMENT", "Enc_Code_Débiteur = '" & Forms!FrmFicheDebiteur.Déb_Code & "' AND Enc_Id<>" & Enc_ID)), 2)

        Select Case CurMontantPrincipal
           Case Is < (CurEncaisserPrincipal + Enc_Montant_m1)

              Enc_Prin_m1 = CurMontantPrincipal - CurEncaisserPrincipal

              Enc_Frais_m1 = Enc_Montant_m1 - Enc_Prin_m1

           Case Else

              Enc_Prin_m1 = Enc_Montant_m1

              Enc_Frais_m1 = 0

        End Select

       Enc_Montant_m2 = Enc_Montant_m1 * DonnerTauxConvertion(iMonnaieVisu)

       Enc_Prin_m2 = Enc_Prin_m1 * DonnerTauxConvertion(iMonnaieVisu)

       Enc_Frais_m2 = Enc_Frais_m1 * DonnerTauxConvertion(iMonnaieVisu)

     Case cEURO

        Enc_Prin_m2 = 0

        Enc_Frais_m2 = 0

        CurMontantPrincipal = Round(Forms!FrmFicheDebiteur.Déb_Mont_Prin_m2, 2)

        CurFrais = Round(Forms!FrmFicheDebiteur.Déb_Frais_m2, 2)

        CurEncaisserPrincipal = Round(Nz(DSum("Enc_Prin_m2", "ENCAISSEMENT", "Enc_Code_Débiteur = '" & Forms!FrmFicheDebiteur.Déb_Code & "' AND Enc_Id<>" & Enc_ID)), 2)

        Select Case CurMontantPrincipal
           Case Is < (CurEncaisserPrincipal + Enc_Montant_m2)

              Enc_Prin_m2 = CurMontantPrincipal - CurEncaisserPrincipal

              Enc_Frais_m2 = Enc_Montant_m2 - Enc_Prin_m2

           Case Else

              Enc_Prin_m2 = Enc_Montant_m2

              Enc_Frais_m2 = 0

        End Select

       Enc_Montant_m1 = Enc_Montant_m2 * DonnerTauxConvertion(iMonnaieVisu)

       Enc_Prin_m1 = Enc_Prin_m2 * DonnerTauxConvertion(iMonnaieVisu)

       Enc_Frais_m1 = Enc_Frais_m2 * DonnerTauxConvertion(iMonnaieVisu)

   End Select

   TxtEnc_Prin_m.Requery

   TxtEnc_Frais_m.Requery

Exit_TxtEnc_Montant_m_AfterUpdate:

   Exit Sub

Err_TxtEnc_Montant_m_AfterUpdate:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_TxtEnc_Montant_m_AfterUpdate
End Sub

Private Sub TxtEnc_Prin_m_AfterUpdate()

   On Error GoTo Err_TxtEnc_Prin_m_AfterUpdate

   'Mise à jour des valeurs de l'autre monnaie
   Select Case iMonnaieVisu
     Case cFRANC

        Select Case Enc_Prin_m1
           Case Is > Enc_Montant_m1

              Enc_Montant_m1 = Enc_Prin_m1 + Enc_Frais_m1

           Case Else

              Enc_Frais_m1 = Enc_Montant_m1 - Enc_Prin_m1

        End Select

        Enc_Montant_m2 = Enc_Montant_m1 * DonnerTauxConvertion(iMonnaieVisu)

        Enc_Prin_m2 = Enc_Prin_m1 * DonnerTauxConvertion(iMonnaieVisu)

        Enc_Frais_m2 = Enc_Frais_m1 * DonnerTauxConvertion(iMonnaieVisu)

     Case cEURO

        Select Case Enc_Prin_m2
           Case Is > Enc_Montant_m2

              Enc_Montant_m2 = Enc_Prin_m2 + Enc_Frais_m2

           Case Else

              Enc_Frais_m2 = Enc_Montant_m2 - Enc_Prin_m2

        End Select

        Enc_Montant_m1 = Enc_Montant_m2 * DonnerTauxConvertion(iMonnaieVisu)

        Enc_Prin_m1 = Enc_Prin_m2 * DonnerTauxConvertion(iMonnaieVisu)

        Enc_Frais_m1 = Enc_Frais_m2 * DonnerTauxConvertion(iMonnaieVisu)

   End Select

   TxtEnc_Montant_m.Requery

   TxtEnc_Frais_m.Requery

Exit_TxtEnc_Prin_m_AfterUpdate:

    Exit Sub

Err_TxtEnc_Prin_m_AfterUpdate:

    MsgBox Err.Number & " " & Err.description

    Resume Exit_TxtEnc_Prin_m_AfterUpdate
End Sub

Private Sub CmdEncaissementDetail_Click()

   On Error GoTo Err_CmdEncaissementDetail_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case 0

      Case 2046

         Err.Clear

      Case Else

         GoTo Err_CmdEncaissementDetail_Click

   End Select

   Select Case IsNull(Enc_ID)
      Case True

      Case False

         DoCmd.OpenForm "FrmFicheDebiteurDetailEncaissement", acNormal, , , , acDialog, Enc_ID & ";" & iMonnaieVisu

   End Select

Exit_CmdEncaissementDetail_Click:

   Exit Sub

Err_CmdEncaissementDetail_Click:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_CmdEncaissementDetail_Click
End Sub
