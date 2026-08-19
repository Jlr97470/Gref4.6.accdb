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

   TxtMon_Libelle = DLookup("Mon_Nom", "MONNAIE", "Mon_Id=" & IntMonnaie)

   iMonnaieVisu = IntMonnaie

Exit_Inverseur:

    Exit Sub

Err_Inverseur:

    MsgBox Err.Number & " " & Err.Number, vbCritical, stSource & " : Inverseur"

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

   End Select

   StrCli_Code = StrParam(0)

   Inverseur Nz(StrParam(1), cEURO)

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
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

Private Sub CmdOK_Click()
   Dim RstClient As DAO.Recordset
   Dim RstClientContrat As DAO.Recordset

   On Error GoTo Err_CmdOK_Click

   Select Case IsNull(TxtCli_Date_contrat)
      Case True

        MsgBox "Le date de contrat est obligatoire.", vbInformation, stSource

        GoTo Exit_CmdOK_Click

      Case False

   End Select

   Select Case True
      Case TxtCli_Mont_Hon_m > 0 And TxtCli_Pourcentage = 0

      Case TxtCli_Mont_Hon_m = 0 And TxtCli_Pourcentage > 0

      Case TxtCli_Mont_Hon_m = 0 And TxtCli_Pourcentage = 0

        MsgBox "Le type de contrat est obligatoire.", vbInformation, stSource

        GoTo Exit_CmdOK_Click

      Case Else

        MsgBox "Ne choisissez qu'un seul type de Contrat.", vbInformation, stSource

        GoTo Exit_CmdOK_Click

   End Select

   Set RstClient = CurrentDb.OpenRecordset("SELECT CLIENT.* FROM CLIENT WHERE Cli_Code=" & Chr(34) & StrCli_Code & Chr(34) & ";")

   Set RstClientContrat = CurrentDb.OpenRecordset("CLIENT_CONTRAT")

   RstClientContrat.AddNew

   RstClientContrat("CC_Code_Client") = RstClient("Cli_Code")

   RstClientContrat("CC_Date_Contrat") = RstClient("Cli_Date_contrat")

   RstClientContrat("CC_Mont_Hon_m1") = RstClient("Cli_Mont_Hon_m1")

   RstClientContrat("CC_Mont_Hon_m2") = RstClient("Cli_Mont_Hon_m2")

   RstClientContrat("CC_Mont_Hon_mo") = RstClient("Cli_Mont_Hon_mo")

   RstClientContrat("CC_Pourcentage") = RstClient("Cli_Pourcentage")

   RstClientContrat.Update

   RstClientContrat.Close

   Set RstClientContrat = Nothing

   RstClient.Edit

   RstClient("Cli_Date_contrat") = TxtCli_Date_contrat

   Select Case iMonnaieVisu
      Case cFRANC

         RstClient("Cli_Mont_Hon_m1") = TxtCli_Mont_Hon_m

         RstClient("Cli_Mont_Hon_m2") = RstClient("Cli_Mont_Hon_m1") * DonnerTauxConvertion(iMonnaieVisu)

      Case cEURO

         RstClient("Cli_Mont_Hon_m2") = TxtCli_Mont_Hon_m

         RstClient("Cli_Mont_Hon_m1") = RstClient!Cli_Mont_Hon_m2 * DonnerTauxConvertion(iMonnaieVisu)

   End Select

   RstClient("Cli_Mont_Hon_mo") = iMonnaieVisu

   RstClient("Cli_Pourcentage") = TxtCli_Pourcentage

   RstClient.Update

   RstClient.Close

   Set RstClient = Nothing

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdOK_Click:

    Exit Sub

Err_CmdOK_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

    Resume Exit_CmdOK_Click
End Sub
