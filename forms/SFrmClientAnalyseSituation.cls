Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private iMonnaieVisu As Integer

Public Sub Inverseur(ByVal IntMonnaie As Integer)
    'Cette fonction permet de visualiser une fiche dans une monnaie particulière (argMonnaie)
    On Error GoTo Err_Inverseur

   Select Case IntMonnaie
       Case cFRANC

           TxtSolde.ControlSource = "=[Déb_Mont_Prin_m1]-[Déb_Prin_Enc_m1]"

           TxtTotalSolde.ControlSource = "=Sum([Déb_Mont_Prin_m1]-[Déb_Prin_Enc_m1])"

           TxtTotalDéb_Mont_Reversé_m.ControlSource = "=Sum([Déb_Mont_Reversé_m1])+Round(Sum([Déb_Mont_Reversé_m2])*DonnerTauxConvertion(2),2)"

           TxtTotalDéb_Prin_Enc_m.ControlSource = "=Sum([Déb_Prin_Enc_m1])"

       Case cEURO

           TxtSolde.ControlSource = "=[Déb_Mont_Prin_m2]-[Déb_Prin_Enc_m2]"

           TxtTotalSolde.ControlSource = "=Sum([Déb_Mont_Prin_m2]-[Déb_Prin_Enc_m2])"

           TxtTotalDéb_Mont_Reversé_m.ControlSource = "=Sum([Déb_Mont_Reversé_m2])+Round(Sum([Déb_Mont_Reversé_m1])*DonnerTauxConvertion(1),2)"

           TxtTotalDéb_Prin_Enc_m.ControlSource = "=Sum([Déb_Prin_Enc_m2])"

   End Select

   iMonnaieVisu = IntMonnaie

Exit_Inverseur:

    Exit Sub

Err_Inverseur:

    MsgBox Err.Number & " : " & Err.description, vbCritical, stSource & " : Inverseur"

    Resume Exit_Inverseur
End Sub
