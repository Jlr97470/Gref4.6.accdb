Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private StrDéb_Code As String
Private iMonnaieVisu As Integer

Public Sub Inverseur(ByVal IntMonnaie As Integer)
    'Cette fonction permet de visualiser une fiche dans une monnaie particulière (argMonnaie)
    On Error GoTo Err_Inverseur

   Select Case IntMonnaie
       Case cFRANC

           TxtDe_Mont_éch_m.ControlSource = "De_Mont_éch_m1"

           TxtTotal.ControlSource = "=Round(Nz(Sum([De_Mont_éch_m1])),2)"

       Case cEURO

           TxtDe_Mont_éch_m.ControlSource = "De_Mont_éch_m2"

           TxtTotal.ControlSource = "=Round(Nz(Sum([De_Mont_éch_m2])),2)"

   End Select

   iMonnaieVisu = IntMonnaie

   TxtDe_Mont_éch_mo.DefaultValue = IntMonnaie

Exit_Inverseur:

    Exit Sub

Err_Inverseur:

    MsgBox Err.Number & " : " & Err.description, vbCritical, stSource & " : Inverseur"

    Resume Exit_Inverseur
End Sub

Private Sub Form_Open(Cancel As Integer)
   Dim StrParam() As String

   On Error GoTo Err_Form_Open

   StrParam = Split(Nz(Me.OpenArgs), ";")

   Select Case UBound(StrParam)
      Case Is < 1

         Cancel = True

         GoTo Exit_Form_Open

      Case Else

         StrDéb_Code = StrParam(0)

         Me.RecordSource = "SELECT DEBITEUR_ECHEANCE.* FROM DEBITEUR_ECHEANCE WHERE De_Code_Débiteur= '" & StrDéb_Code & "';"

         Inverseur Nz(StrParam(1), cEURO)

   End Select

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)

   On Error GoTo Err_Form_BeforeInsert

   De_Code_Débiteur = StrDéb_Code

Exit_Form_BeforeInsert:

   Exit Sub

Err_Form_BeforeInsert:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_BeforeInsert"

   Resume Exit_Form_BeforeInsert
End Sub

Private Sub TxtDe_Date_éch_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_TxtDe_Date_éch_BeforeUpdate

   Select Case IsNull(TxtDe_Date_éch)
      Case True

         MsgBox "Date Obligatoire !!", vbCritical

         Cancel = True

      Case False

   End Select

Exit_TxtDe_Date_éch_BeforeUpdate:

   Exit Sub

Err_TxtDe_Date_éch_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDe_Date_éch_BeforeUpdate"

   Resume Exit_TxtDe_Date_éch_BeforeUpdate
End Sub

Private Sub TxtDe_Mont_éch_m_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_TxtDe_Mont_éch_m_BeforeUpdate

   Select Case IsNull(TxtDe_Date_éch)
      Case True

         MsgBox "Montant Obligatoire !!", vbCritical

         Cancel = True

      Case False

   End Select

Exit_TxtDe_Mont_éch_m_BeforeUpdate:

   Exit Sub

Err_TxtDe_Mont_éch_m_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDe_Mont_éch_m_BeforeUpdate"

   Resume Exit_TxtDe_Mont_éch_m_BeforeUpdate
End Sub

Private Sub TxtDe_Mont_éch_m_AfterUpdate()

   On Error GoTo Err_TxtDe_Mont_éch_AfterUpdate
   'Mise à jour des valeurs de l'autre monnaie
   Select Case iMonnaieVisu
     Case cFRANC

       De_Mont_éch_m2 = De_Mont_éch_m1 * DonnerTauxConvertion(iMonnaieVisu)

     Case cEURO

       De_Mont_éch_m1 = De_Mont_éch_m2 * DonnerTauxConvertion(iMonnaieVisu)

   End Select

Exit_TxtDe_Mont_éch_AfterUpdate:

   Exit Sub

Err_TxtDe_Mont_éch_AfterUpdate:

   MsgBox Err.Number & " : " & Err.description, vbCritical, stSource & " : Inverseur"

   Resume Exit_TxtDe_Mont_éch_AfterUpdate
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.RunCommand acCmdSaveRecord

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " : " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdImprimer_Click()

   On Error GoTo Err_CmdImprimer_Click

Exit_CmdImprimer_Click:

   Exit Sub

Err_CmdImprimer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdImprimer_Click"

   Resume Exit_CmdImprimer_Click
End Sub

Private Sub CmdOK_Click()

   On Error GoTo Err_CmdOK_Click

   DoCmd.RunCommand acCmdSaveRecord

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdOK_Click:

    Exit Sub

Err_CmdOK_Click:

   MsgBox Err.Number & " : " & Err.description, vbCritical, stSource & " : CmdOK_Click"

    Resume Exit_CmdOK_Click
End Sub
