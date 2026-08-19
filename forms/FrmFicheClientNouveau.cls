Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_Form_BeforeUpdate

   Select Case True
      Case IsNull(Cli_Code), IsNull(Cli_Nom), IsNull(Cli_Représentant)

         MsgBox "Le Code, le Nom et le Représentant sont obligatoires.", vbInformation, stSource

         Cancel = True

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case DLookup("Cli_Code", "CLIENT", "Cli_Code = '" & Cli_Code & "'")
     Case Cli_Code

        Cancel = True

        TxtCli_Code.SetFocus

        MsgBox "Ce code existe déjà !", vbExclamation, stSource

        GoTo Exit_Form_BeforeUpdate

     Case Else

   End Select

   Select Case True
      Case IsNull(Cli_Mont_Hon_mo)

        MsgBox "Veuillez indiquer la monnaie de traitement du client.", vbInformation, stSource

        Cancel = True

        GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case IsNull(Cli_Date_contrat)
     Case True

        Cancel = True

        TxtCli_Date_contrat.SetFocus

        MsgBox "La date de contrat est obligatoire !", vbExclamation, stSource

        GoTo Exit_Form_BeforeUpdate

     Case False

   End Select

   Select Case True
      Case TxtCli_Mont_Hon_m > 0 And Cli_Pourcentage = 0

      Case TxtCli_Mont_Hon_m = 0 And Cli_Pourcentage > 0

      Case TxtCli_Mont_Hon_m = 0 And Cli_Pourcentage = 0

        Cancel = True

        MsgBox "Le type de contrat est obligatoire.", vbInformation, stSource

        GoTo Exit_Form_BeforeUpdate

      Case Else

        Cancel = True

        MsgBox "Ne choisissez qu'un seul type de Contrat.", vbInformation, stSource

        GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case Cli_Mont_Hon_mo
       Case cFRANC

           Cli_Mont_Hon_m2 = Cli_Mont_Hon_m1 * DonnerTauxConvertion(Cli_Mont_Hon_mo)

       Case cEURO

           Cli_Mont_Hon_m1 = Cli_Mont_Hon_m2 * DonnerTauxConvertion(Cli_Mont_Hon_mo)

   End Select

   Cli_MonnaieEnCours = Cli_Mont_Hon_mo

Exit_Form_BeforeUpdate:

   Exit Sub

Err_Form_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_BeforeUpdate"

   Resume Exit_Form_BeforeUpdate
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

Private Sub CmbCli_Mont_Hon_mo_AfterUpdate()

   On Error GoTo Err_CmbCli_Mont_Hon_mo_AfterUpdate

   Select Case Me.Cli_Mont_Hon_mo
      Case cFRANC

         TxtCli_Mont_Hon_m.ControlSource = "Cli_Mont_Hon_m1"

         Cli_Mont_Hon_m1 = 0

      Case cEURO

         TxtCli_Mont_Hon_m.ControlSource = "Cli_Mont_Hon_m2"

         Cli_Mont_Hon_m2 = 0

   End Select

Exit_CmbCli_Mont_Hon_mo_AfterUpdate:

   Exit Sub

Err_CmbCli_Mont_Hon_mo_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbCli_Mont_Hon_mo_AfterUpdate"

   Resume Exit_CmbCli_Mont_Hon_mo_AfterUpdate
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

Private Sub CmdTypeClient_Click()

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

   DoCmd.OpenForm "FrmParam_SousGroupeClient", acNormal, , , , acDialog

   CmbCli_Groupe.Requery

Exit_CmdGroupe_Click:

   Exit Sub

Err_CmdGroupe_Click:

   MsgBox Err.Number & "  " & Err.description, vbCritical, stSource & " : CmdGroupe_Click"

   Resume Exit_CmdGroupe_Click
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

         Set CtlControl = Forms!FrmFicheClient.Form.ActiveControl

         Forms!FrmFicheClient.Filter = "Cli_Code='" & Cli_Code & "'"

         Forms!FrmFicheClient.FilterOn = True

      Case False

         DoCmd.OpenForm "FrmFicheClient", acNormal, , "Cli_Code='" & Cli_Code & "'", , acWindowNormal

         Set CtlControl = Forms!FrmFicheClient.Form.ActiveControl

   End Select

   CtlControl.SetFocus

   IntTypeAdrCourrier = Nz(DLookup("Adresse_Courrier", "PARAMETRE_GREF"))

   Select Case IntTypeAdrCourrier
      Case 0

         MsgBox "ATTENTION !" & Chr(10) & "Le parametres type d'adresse par defaut est abscent dans les parametres GREF.", vbExclamation, stSource

         GoTo Exit_CmdOK_Click

      Case Else

   End Select

   DoCmd.OpenForm "FrmFicheClientDetailAdresse", acNormal, , , acFormAdd, acWindowNormal, Cli_Code & ";" & IntTypeAdrCourrier

   DoCmd.Close acForm, Me.Name, acSavePrompt

   MsgBox "ATTENTION !" & Chr(10) & "N'oubliez pas de choisir une grille de Frais.", vbExclamation, stSource

Exit_CmdOK_Click:

    Exit Sub

Err_CmdOK_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

    Resume Exit_CmdOK_Click
End Sub
