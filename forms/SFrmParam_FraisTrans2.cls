Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private iMonnaieVisu As Integer

Public Sub Inverseur(argMonnaie As Integer)
    'Cette fonction permet de visualiser le formulaire dans une monnaie particulière (argMonnaie)
   On Error GoTo Err_Inverseur

   Select Case argMonnaie
       Case cFRANC

           Tranche.ControlSource = "=iif([Frt_Fin_m1]=0,""Suppérieur à "" & Format([Frt_Début_m1],""# ##0.00""),""De "" & Format([Frt_Début_m1],""# ##0.00"") & "" à "" & Format([Frt_Fin_m1],""# ##0.00"")) & "" F"""

           Frt_Frais.ControlSource = "[Frt_Frais_m1]"

       Case cEURO

           Tranche.ControlSource = "=iif([Frt_Fin_m2]=0,""Suppérieur à "" & Format([Frt_Début_m2],""# ##0.00""),""De "" & Format([Frt_Début_m2],""# ##0.00"") & "" à "" & Format([Frt_Fin_m2],""# ##0.00"")) & "" E"""

           Frt_Frais.ControlSource = "[Frt_Frais_m2]"

   End Select

   iMonnaieVisu = argMonnaie

Exit_Inverseur:

   Exit Sub

Err_Inverseur:

   MsgBox Err.Number & " : " & Err.description, vbCritical, stSource & " : Inverseur"

   Resume Exit_Inverseur
End Sub

Private Sub Frt_Frais_AfterUpdate()
    'Mise à jour des valeurs de l'autre monnaie
    If iMonnaieVisu = cFRANC Then
        Me.[Frt_Frais_m2] = Round(Me.[Frt_Frais_m1] * DonnerTauxConvertion(iMonnaieVisu), 2)
    Else
        Me.[Frt_Frais_m1] = Round(Me.[Frt_Frais_m2] * DonnerTauxConvertion(iMonnaieVisu), 2)
    End If
End Sub
