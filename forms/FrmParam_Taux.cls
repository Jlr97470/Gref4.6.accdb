Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Function ImpVisu(ByVal IntVisu As Integer)

    On Error GoTo Err_ImpVisu

    DoCmd.OpenReport "EtaTauxRelance", IntVisu

Exit_ImpVisu:

    Exit Function

Err_ImpVisu:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ImpVisu"

    Resume Exit_ImpVisu
End Function

Private Sub CmbTxG_Grille_AfterUpdate()

   On Error GoTo Err_CmbTxG_Grille_AfterUpdate

   Select Case DCount("Tx_N°", "TAUX", "Tx_Grille=" & CmbTxG_Grille)
      Case Is > 0

      Case Else

         CurrentDb.Execute "INSERT INTO TAUX ( Tx_Grille, Tx_Code_Action, Tx_N°, Tx_Délai, Tx_Frais ) " & _
                                       "SELECT " & CmbTxG_Grille & ", RELANCE.Rel_Code_Action, RELANCE.Rel_N°, 0 , 0 FROM RELANCE WHERE Rel_Code_Action<=3;", dbFailOnError

   End Select

   Me.Requery

Exit_CmbTxG_Grille_AfterUpdate:

   Exit Sub

Err_CmbTxG_Grille_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbTxG_Grille_AfterUpdate"

   Resume Exit_CmbTxG_Grille_AfterUpdate
End Sub

Private Sub CmdFemer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdTauxGrille_Click()

   On Error GoTo Err_CmdTauxGrille_Click

   DoCmd.OpenForm "FrmParam_TauxGrille", acNormal, , , , acDialog

   CmbTxG_Grille.Requery

Exit_CmdTauxGrille_Click:

   Exit Sub

Err_CmdTauxGrille_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdTauxGrille_Click"

   Resume Exit_CmdTauxGrille_Click
End Sub
