Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_CmdFermer_Click
End Sub

Private Sub Form_AfterInsert()
   Dim stSQL As String

   On Error GoTo Err_Form_AfterInsert

   With CurrentDb

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",0,100) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",100.01,200) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",200.01,300) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",300.01,400) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",400.01,500) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",500.01,700) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",700.01,900) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",900.01,1100) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",1100.01,1500) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",1500.01,2000) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",2000.01,2500) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",2500.01,3000) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",3000.01,3500) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",3500.01,4000) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",4000.01,4500) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",4500.01,5000) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",5000.01,5500) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "INSERT INTO FRAIS_TRANS (Frt_Grille,Frt_Début_m1,Frt_Fin_m1) VALUES (" & Me.Frt_Grille & ",5500.01,0) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "UPDATE FRAIS_TRANS SET [Frt_Début_m2] = Round((Frt_Début_m1*" & Replace(DonnerTauxConvertion(cFRANC), ",", ".") & "),2) ;"

        .Execute stSQL, dbFailOnError

        stSQL = "UPDATE FRAIS_TRANS SET [Frt_Fin_m2] = Round((Frt_Fin_m1*" & Replace(DonnerTauxConvertion(cFRANC), ",", ".") & "),2) ;"

        .Execute stSQL, dbFailOnError

    End With

Exit_Form_AfterInsert:

   Exit Sub

Err_Form_AfterInsert:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_AfterInsert"

   Resume Exit_Form_AfterInsert
End Sub
