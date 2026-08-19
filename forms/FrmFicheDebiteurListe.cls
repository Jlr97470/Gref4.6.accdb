Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private iMonnaieVisu As Integer

Public SQLBase As String
Public SQLWhere As String
Public SQLWhereGlobal As String
Public SQLOrder As String

Public Sub Inverseur(ByVal IntMonnaie As Integer)

   On Error GoTo Err_Inverseur

   Select Case IntMonnaie
       Case cFRANC

           TxtDéb_Mont_Prin_m.ControlSource = "Déb_Mont_Prin_m1"

       Case cEURO

           TxtDéb_Mont_Prin_m.ControlSource = "Déb_Mont_Prin_m2"

   End Select

   Me.Filter = "Déb_MonnaieEnCours = " & IntMonnaie

   Me.FilterOn = True

   CmbMonnaie = IntMonnaie

   iMonnaieVisu = IntMonnaie

Exit_Inverseur:

    Exit Sub

Err_Inverseur:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Inverseur"

    Resume Exit_Inverseur
End Sub

Private Function ImpVisu(ByVal IntVisu As Integer)
   Dim IntRaportEtat As Integer
   Dim IntRepMsg As Integer

   On Error GoTo Err_ImpVisu

   DoCmd.OpenReport "EtaListeDebiteurDetail", acDesign

   Reports("EtaListeDebiteurDetail").Filter = SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere)

   Reports("EtaListeDebiteurDetail").FilterOn = True

   DoCmd.Close acReport, "EtaListeDebiteurDetail", acSavePrompt

   Select Case IntVisu
      Case 0, 2

         On Error Resume Next

         DoCmd.OpenReport "EtaListeDebiteurDetail", IntVisu

         Select Case Err.Number
            Case 0

            Case 2501

               Err.Clear

            Case Else

               GoTo Err_ImpVisu

         End Select

         On Error GoTo Err_ImpVisu

         Do While DonneEtatRaport("EtaListeDebiteurDetail", IntRaportEtat) = True

            DoEvents

         Loop

      Case 5

         IntRepMsg = MsgBox("Souhaitez-vous exporter les données au format Excel ?" & vbCrLf & "Indiquez OUI pour extraire les données au format excel et NON pour le format texte.", vbYesNo + vbQuestion, "Format d'exportation")

         Select Case IntRepMsg
            Case vbYes

               DoCmd.OutputTo acOutputReport, "EtaListeDebiteurDetail", acFormatXLS, , False

            Case vbNo

               DoCmd.OutputTo acOutputReport, "EtaListeDebiteurDetail", acFormatRTF, , False

         End Select

   End Select

Exit_ImpVisu:

   Exit Function

Err_ImpVisu:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Err_ImpVisu"

   Resume Exit_ImpVisu
End Function

Private Sub Form_Open(Cancel As Integer)
   Dim StrParam() As String

   On Error GoTo Err_Form_Open

   StrParam = Split(Nz(Me.OpenArgs), ";")

   Select Case UBound(StrParam)
      Case Is < 1

         Inverseur cEURO

         SQLWhereGlobal = ""

         SQLOrder = ""

      Case Else

         Inverseur Nz(StrParam(1), cEURO)

         SQLWhereGlobal = StrParam(0)

         SQLOrder = "Déb_Code"

   End Select

   SQLBase = "SELECT * FROM DEBITEUR"

   SQLWhere = ""

   Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub CmbMonnaie_AfterUpdate()

   On Error GoTo Err_CmbMonnaie_AfterUpdate

   Inverseur CmbMonnaie

Exit_CmbMonnaie_AfterUpdate:

   Exit Sub

Err_CmbMonnaie_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbMonnaie_AfterUpdate"

   Resume Exit_CmbMonnaie_AfterUpdate
End Sub

Private Sub CmdFermer_Click()

    On Error GoTo Err_CmdFermer_Click

    StrFicheDebiteurListeDéb_Code = ""

    DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

    Exit Sub

Err_CmdFermer_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

    Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdRechercher_Click()

   On Error GoTo Err_CmdRechercher_Click

   DoCmd.OpenForm "FrmFicheDebiteurRechercher", acNormal, , , , acDialog, "FrmFicheDebiteurListe"

   Select Case SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"
     Case Me.RecordSource

     Case Else

        Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

   End Select

Exit_CmdRechercher_Click:

   Exit Sub

Err_CmdRechercher_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdRechercher_Click"

   Resume Exit_CmdRechercher_Click
End Sub

Private Sub CmdOK_Click()

   On Error GoTo Err_CmdOK_Click

    StrFicheDebiteurListeDéb_Code = Déb_Code

    DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdOK_Click:

    Exit Sub

Err_CmdOK_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

    Resume Exit_CmdOK_Click
End Sub
