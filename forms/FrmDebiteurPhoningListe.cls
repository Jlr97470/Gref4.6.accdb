Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Function ImpVisu(ByVal IntVisu As Integer)

End Function

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub Form_Load()
   Dim StrDéb_Code As String
   Dim StrParam() As String
   Dim RstDebiteur  As DAO.Recordset

   On Error GoTo Err_Form_Load

   StrParam = Split(Nz(Me.OpenArgs), ";")

   Select Case UBound(StrParam)
      Case -1

      Case Else

         StrDéb_Code = StrParam(0)

         Set RstDebiteur = Me.RecordsetClone

         RstDebiteur.FindFirst "Déb_Code='" & StrDéb_Code & "'"

         Me.Bookmark = RstDebiteur.Bookmark

         RstDebiteur.Close

   End Select

Exit_Form_Load:

   Set RstDebiteur = Nothing

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Load"

   Resume Exit_Form_Load
End Sub

Private Sub TxtDéb_Situation_AfterUpdate()
   Dim IntCode_Action As Integer
   Dim IntDelai As Integer
   Dim IntGrilleTauxNum As Integer
   Dim IntRep As Integer

   On Error GoTo Err_TxtDéb_Situation_AfterUpdate

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

Exit_TxtDéb_Situation_AfterUpdate:

   Exit Sub

Err_TxtDéb_Situation_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDéb_Situation_AfterUpdate"

   Resume Exit_TxtDéb_Situation_AfterUpdate
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

Private Sub CmdFermer_Click()
   Dim IntEtat As Integer
   Dim StrOrder As String
   Dim StrChamp As String

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

   Select Case DonneEtatFormulaire("FrmDebiteurPhoningSelectionListe", IntEtat)
      Case True

         DoCmd.OpenForm "FrmDebiteurPhoningSelectionListe"

         StrOrder = IIf(Forms!FrmDebiteurPhoningSelectionListe.SQLOrder <> "", Forms!FrmDebiteurPhoningSelectionListe.SQLOrder, "Déb_Code")

         StrChamp = Nz(Forms!FrmDebiteurPhoningSelectionListe.Recordset(StrOrder))

         Forms!FrmDebiteurPhoningSelectionListe.Form.Requery

         Forms!FrmDebiteurPhoningSelectionListe.Recordset.FindFirst StrOrder & ">='" & StrChamp & "'"

      Case False

   End Select

Exit_CmdFermer_Click:

    Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

    Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdDetail_Click()

   On Error GoTo Err_CmdDetail_Click

   DoCmd.Minimize

   DoCmd.OpenForm "FrmFicheDebiteur", acNormal, , "Déb_Code='" & Déb_Code & "'", , acWindowNormal

Exit_CmdDetail_Click:

   Exit Sub

Err_CmdDetail_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdDetail_Click"

   Resume Exit_CmdDetail_Click
End Sub
