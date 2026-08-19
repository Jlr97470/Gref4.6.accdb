Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private iMonnaieVisu As Integer 'Monnaie dans laquelle est visualisée la fiche en cours

Public Sub Inverseur(ByVal IntMonnaie As Integer)
    'Cette fonction permet de visualiser une fiche dans une monnaie particulière (argMonnaie)
    On Error GoTo Err_Inverseur

   Select Case IntMonnaie
       Case cFRANC

           TexteT1.Caption = "De " & 0 & " à " & 100 & " F"

           TexteT2.Caption = "De " & 100.01 & " à " & 300 & " F"

           TexteT3.Caption = "De " & 300.01 & " à " & 500 & " F"

           TexteT4.Caption = "De " & 500.01 & " à " & 1000 & " F"

           TexteT5.Caption = "De " & 1000.01 & " à " & 1500 & " F"

           TexteT6.Caption = "De " & 1500.01 & " à " & 2000 & " F"

           TexteT7.Caption = "De " & 2000.01 & " à " & 3000 & " F"

           TexteT8.Caption = "De " & 3000.01 & " à " & 4000 & " F"

           TexteT9.Caption = "De " & 4000.01 & " à " & 5000 & " F"

           TexteT10.Caption = "De " & 5000.01 & " à " & 7000 & " F"

           TexteT11.Caption = "De " & 7000.01 & " à " & 10000 & " F"

           TexteT12.Caption = "De " & 10000.01 & " à " & 15000 & " F"

           TexteT13.Caption = "De " & 15000.01 & " à " & 20000 & " F"

           TexteT14.Caption = "De " & 20000.01 & " à " & 30000 & " F"

           TexteT15.Caption = "De " & 30000.01 & " à " & 40000 & " F"

           TexteT16.Caption = "De " & 40000.01 & " à " & 50000 & " F"

           TexteT17.Caption = "De " & 50000.01 & " à " & 70000 & " F"

           TexteT18.Caption = "Suppérieur à " & 70000 & " F"

           Fra_Tranche1.ControlSource = "Fra_Tranche1_m1"

           Fra_Tranche2.ControlSource = "Fra_Tranche2_m1"

           Fra_Tranche3.ControlSource = "Fra_Tranche3_m1"

           Fra_Tranche4.ControlSource = "Fra_Tranche4_m1"

           Fra_Tranche5.ControlSource = "Fra_Tranche5_m1"

           Fra_Tranche6.ControlSource = "Fra_Tranche6_m1"

           Fra_Tranche7.ControlSource = "Fra_Tranche7_m1"

           Fra_Tranche8.ControlSource = "Fra_Tranche8_m1"

           Fra_Tranche9.ControlSource = "Fra_Tranche9_m1"

           Fra_Tranche10.ControlSource = "Fra_Tranche10_m1"

           Fra_Tranche11.ControlSource = "Fra_Tranche11_m1"

           Fra_Tranche12.ControlSource = "Fra_Tranche12_m1"

           Fra_Tranche13.ControlSource = "Fra_Tranche13_m1"

           Fra_Tranche14.ControlSource = "Fra_Tranche14_m1"

           Fra_Tranche15.ControlSource = "Fra_Tranche15_m1"

           Fra_Tranche16.ControlSource = "Fra_Tranche16_m1"

           Fra_Tranche17.ControlSource = "Fra_Tranche17_m1"

           Fra_Tranche18.ControlSource = "Fra_Tranche18_m1"

       Case cEURO

           TexteT1.Caption = "De " & 0 * DonnerTauxConvertion(cFRANC) & " à " & Format(100 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT2.Caption = "De " & Format(100.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(300 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT3.Caption = "De " & Format(300.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(500 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT4.Caption = "De " & Format(500.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(1000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT5.Caption = "De " & Format(1000.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(1500 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT6.Caption = "De " & Format(1500.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(2000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT7.Caption = "De " & Format(2000.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(3000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT8.Caption = "De " & Format(3000.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(4000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT9.Caption = "De " & Format(4000.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(5000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT10.Caption = "De " & Format(5000.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(7000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT11.Caption = "De " & Format(7000.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(10000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT12.Caption = "De " & Format(10000.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(15000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT13.Caption = "De " & Format(15000.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(20000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT14.Caption = "De " & Format(20000.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(30000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT15.Caption = "De " & Format(30000.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(40000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT16.Caption = "De " & Format(40000.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(50000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT17.Caption = "De " & Format(50000.01 * DonnerTauxConvertion(cFRANC), "# ###.##") & " à " & Format(70000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           TexteT18.Caption = "Suppérieur à " & Format(70000 * DonnerTauxConvertion(cFRANC), "# ###.##") & " E"

           Fra_Tranche1.ControlSource = "Fra_Tranche1_m2"

           Fra_Tranche2.ControlSource = "Fra_Tranche2_m2"

           Fra_Tranche3.ControlSource = "Fra_Tranche3_m2"

           Fra_Tranche4.ControlSource = "Fra_Tranche4_m2"

           Fra_Tranche5.ControlSource = "Fra_Tranche5_m2"

           Fra_Tranche6.ControlSource = "Fra_Tranche6_m2"

           Fra_Tranche7.ControlSource = "Fra_Tranche7_m2"

           Fra_Tranche8.ControlSource = "Fra_Tranche8_m2"

           Fra_Tranche9.ControlSource = "Fra_Tranche9_m2"

           Fra_Tranche10.ControlSource = "Fra_Tranche10_m2"

           Fra_Tranche11.ControlSource = "Fra_Tranche11_m2"

           Fra_Tranche12.ControlSource = "Fra_Tranche12_m2"

           Fra_Tranche13.ControlSource = "Fra_Tranche13_m2"

           Fra_Tranche14.ControlSource = "Fra_Tranche14_m2"

           Fra_Tranche15.ControlSource = "Fra_Tranche15_m2"

           Fra_Tranche16.ControlSource = "Fra_Tranche16_m2"

           Fra_Tranche17.ControlSource = "Fra_Tranche17_m2"

           Fra_Tranche18.ControlSource = "Fra_Tranche18_m2"

   End Select

   iMonnaieVisu = IntMonnaie

Exit_Inverseur:

    Exit Sub

Err_Inverseur:

    MsgBox Err.Number & " : " & Error, vbCritical, stSource & " : Inverseur"

    Resume Exit_Inverseur
End Sub

Private Sub CmbFra_Grille_AfterUpdate()

   On Error GoTo Err_CmbFra_Grille_AfterUpdate

   Select Case DCount("Fra_Grille", "FRAIS", "Fra_Grille=" & CmbFra_Grille)
      Case Is > 0

      Case Else

         CurrentDb.Execute "INSERT INTO FRAIS (Fra_Grille) VALUES (" & CmbFra_Grille & ");", dbFailOnError

   End Select

   Me.Requery

Exit_CmbFra_Grille_AfterUpdate:

   Exit Sub

Err_CmbFra_Grille_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbFra_Grille_AfterUpdate"

   Resume Exit_CmbFra_Grille_AfterUpdate
End Sub

Private Function ImpVisu(IntVisu As Integer)

    On Error GoTo Err_ImpVisu

    DoCmd.OpenReport "EtaFraisRelance", IntVisu, , "Mon_Id=" & iMonnaieVisu & " AND Fra_Grille = " & CmbFra_Grille, acWindowNormal, iMonnaieVisu & ";"

Exit_ImpVisu:

    Exit Function

Err_ImpVisu:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ImpVisu"

    Resume Exit_ImpVisu
End Function

Private Sub CmdFraisGrille_Click()

   On Error GoTo Err_CmdFraisGrille_Click

   DoCmd.OpenForm "FrmParam_FraisGrille", acNormal, , , , acDialog

   CmbFra_Grille.Requery

Exit_CmdFraisGrille_Click:

   Exit Sub

Err_CmdFraisGrille_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFraisGrille_Click"

   Resume Exit_CmdFraisGrille_Click
End Sub

Private Sub CmdInverseur_Click()

   On Error GoTo Err_CmdInverseur_Click

   Select Case iMonnaieVisu
      Case cEURO

        Inverseur cFRANC 'inverser la monnaie de visualisation

      Case cFRANC

        Inverseur cEURO 'inverser la monnaie de visualisation

    End Select

Exit_CmdInverseur_Click:

   Exit Sub

Err_CmdInverseur_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdInverseur_Click"

   Resume Exit_CmdInverseur_Click:
End Sub

Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   Inverseur cEURO 'initialise le formulaire en franc

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Load"

   Resume Exit_Form_Load
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_CmdFermer_Click
End Sub

Private Sub Fra_Tranche1_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche1_m2] = Round(Me.[Fra_Tranche1_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche1_m1] = Round(Me.[Fra_Tranche1_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche2_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche2_m2] = Round(Me.[Fra_Tranche2_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche2_m1] = Round(Me.[Fra_Tranche2_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche3_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche3_m2] = Round(Me.[Fra_Tranche3_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche3_m1] = Round(Me.[Fra_Tranche3_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche4_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche4_m2] = Round(Me.[Fra_Tranche4_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche4_m1] = Round(Me.[Fra_Tranche4_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche5_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche5_m2] = Round(Me.[Fra_Tranche5_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche5_m1] = Round(Me.[Fra_Tranche5_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche6_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche6_m2] = Round(Me.[Fra_Tranche6_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche6_m1] = Round(Me.[Fra_Tranche6_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche7_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche7_m2] = Round(Me.[Fra_Tranche7_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche7_m1] = Round(Me.[Fra_Tranche7_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche8_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche8_m2] = Round(Me.[Fra_Tranche8_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche8_m1] = Round(Me.[Fra_Tranche8_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche9_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche9_m2] = Round(Me.[Fra_Tranche9_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche9_m1] = Round(Me.[Fra_Tranche9_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche10_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche10_m2] = Round(Me.[Fra_Tranche10_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche10_m1] = Round(Me.[Fra_Tranche10_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche11_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche11_m2] = Round(Me.[Fra_Tranche11_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche11_m1] = Round(Me.[Fra_Tranche11_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche12_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche12_m2] = Round(Me.[Fra_Tranche12_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche12_m1] = Round(Me.[Fra_Tranche12_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche13_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche13_m2] = Round(Me.[Fra_Tranche13_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche13_m1] = Round(Me.[Fra_Tranche13_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche14_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche14_m2] = Round(Me.[Fra_Tranche14_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche14_m1] = Round(Me.[Fra_Tranche14_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche15_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche15_m2] = Round(Me.[Fra_Tranche15_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche15_m1] = Round(Me.[Fra_Tranche15_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche16_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche16_m2] = Round(Me.[Fra_Tranche16_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche16_m1] = Round(Me.[Fra_Tranche16_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche17_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche17_m2] = Round(Me.[Fra_Tranche17_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche17_m1] = Round(Me.[Fra_Tranche17_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub

Private Sub Fra_Tranche18_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Fra_Tranche18_m2] = Round(Me.[Fra_Tranche18_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Fra_Tranche18_m1] = Round(Me.[Fra_Tranche18_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub
