Attribute VB_Name = "ModFonction"
Option Compare Database
Option Explicit

Private LngNumero As Long
Private StrCode As String

Function IsNullVide(ByVal VarObjet As Variant) As Boolean

   On Error GoTo Err_IsNullVide

   Select Case IsNull(VarObjet)
     Case True

        IsNullVide = True

      Case False

        Select Case Len(VarObjet)
           Case 0

              IsNullVide = True

           Case Else

              IsNullVide = False

        End Select

   End Select

Exit_IsNullVide:

   Exit Function

Err_IsNullVide:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_IsNullVide
End Function

Public Function RemplaceChr(ByVal StrChaine As String, ByVal StrChaineRechercher As String, ByVal strChaineRemplace As String, Optional Compare) As String
   Dim StrChaineResultat As String
   Dim IntPos As Integer
   Dim IntCompare As Integer

   On Error GoTo Err_RemplaceChr

   If StrChaine = vbNullString Then

       RemplaceChr = vbNullString

       GoTo Exit_RemplaceChr

   End If

   Select Case IsMissing(Compare)
      Case True

         Compare = 1
      Case False

         Select Case VarType(Compare)
            Case vbByte, vbInteger, vbLong

                  Select Case Compare
                     Case Is < vbEmpty

                        MsgBox "Parametre Compare Incorrect", , "RemplaceChr"

                        GoTo Exit_RemplaceChr

                     Case Is > 2

                        MsgBox "Parametre Compare Incorrect", , "RemplaceChr"

                        GoTo Exit_RemplaceChr


                     Case Else

                        IntCompare = Compare

                  End Select
            Case Else

               IntCompare = vbEmpty

         End Select
   End Select

   StrChaineResultat = vbNullString

   IntPos = InStr(1, StrChaine, StrChaineRechercher, IntCompare)

   While (IntPos > vbEmpty)

       StrChaineResultat = StrChaineResultat + Left(StrChaine, IntPos - 1) & strChaineRemplace

       StrChaine = Right(StrChaine, Len(StrChaine) - Len(StrChaineRechercher) - IntPos + 1)

       IntPos = InStr(1, StrChaine, StrChaineRechercher, IntCompare)

   Wend

   RemplaceChr = StrChaineResultat + StrChaine

Exit_RemplaceChr:

   Exit Function

Err_RemplaceChr:

   MsgBox Err.Number & " " & Err.description, , "RemplaceChr"

   Resume Exit_RemplaceChr
End Function

Function ConvDate(ByVal DatDate As Date) As String
   Dim IntJour As Integer
   Dim IntMois As Integer
   Dim IntAnnee As Integer

   On Error GoTo Err_ConvDate

   IntJour = Day(DatDate)

   IntMois = Month(DatDate)

   IntAnnee = Year(DatDate)

   ConvDate = "#" & IntMois & "/" & IntJour & "/" & IntAnnee & "#"

Exit_ConvDate:

   Exit Function

Err_ConvDate:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ConvDate"

    Resume Exit_ConvDate
End Function

Function DonnerTauxConvertion(ByVal IntMonnaieVisu As Integer) As Single

   On Error GoTo Err_DonnerTauxConvertion

   Select Case IsNull(DLookup("TauxConvertion_m1m2", "PARAMETRE_GREF"))
      Case True

         MsgBox "ATTENTION !" & Chr(10) & "Le parametres taux de convertion est abscent dans les parametres GREF.", vbExclamation, stSource

         GoTo Exit_DonnerTauxConvertion

      Case False

   End Select

   Select Case IntMonnaieVisu
      Case cEURO

         DonnerTauxConvertion = DLookup("TauxConvertion_m1m2", "PARAMETRE_GREF")

      Case cFRANC

         DonnerTauxConvertion = 1 / DLookup("TauxConvertion_m1m2", "PARAMETRE_GREF")

    End Select

Exit_DonnerTauxConvertion:

   Exit Function

Err_DonnerTauxConvertion:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : DonnerTauxConvertion"

   Resume Exit_DonnerTauxConvertion
End Function

Public Function DonnerMontantFrais(ByVal StrCodeClient As String, ByVal CurMontant As Currency, ByVal IntMonnaie As Integer) As Currency
   Dim StrModeCalcul As String * 1
   Dim CurMontantAccessoire As Currency
   Dim SngTxTVA As Single
   Dim SngTxLegal As Single
   Dim IntGrilleFrais As Integer

   On Error GoTo Err_DonnerMontantFrais

   StrModeCalcul = Nz(DLookup("Mode_Calcul_Frais", "PARAMETRE_GREF"))

   Select Case StrModeCalcul
      Case ""

         MsgBox "ATTENTION !" & Chr(10) & "Le parametres mode de calcul des frais est abscent dans les parametres GREF.", vbExclamation, stSource

         GoTo Exit_DonnerMontantFrais

      Case Else

   End Select

   SngTxTVA = Nz(DLookup("Taux_TVA", "PARAMETRE_GREF"))

   Select Case SngTxTVA
      Case 0

         MsgBox "ATTENTION !" & Chr(10) & "Le parametres taux TVA est abscent dans les parametres GREF.", vbExclamation, stSource

         GoTo Exit_DonnerMontantFrais

      Case Else

   End Select

   SngTxLegal = Nz(DLookup("Tx_Interet_legal", "PARAMETRE_GREF"))

   Select Case SngTxLegal
      Case 0

         MsgBox "ATTENTION !" & Chr(10) & "Le parametres taux interet legal est abscent dans les parametres GREF.", vbExclamation, stSource

         GoTo Exit_DonnerMontantFrais

      Case Else

   End Select

   CurMontantAccessoire = Nz(DLookup("Frais_Accessoires", "PARAMETRE_GREF"))

   Select Case CurMontantAccessoire
      Case 0

         MsgBox "ATTENTION !" & Chr(10) & "Le parametres frais accessoires est abscent dans les parametres GREF.", vbExclamation, stSource

         GoTo Exit_DonnerMontantFrais

      Case Else

   End Select

   Select Case IntMonnaie
      Case cFRANC

      Case cEURO

         CurMontantAccessoire = CurMontantAccessoire * DonnerTauxConvertion(cFRANC)

   End Select

   Select Case True
      Case (StrModeCalcul = "1") Or (StrModeCalcul = "3" And DLookup("Cli_NewCalculFrais", "CLIENT", "Cli_Code = '" & StrCodeClient & "'") = False)

         IntGrilleFrais = DLookup("Cli_Grille_Frais", "CLIENT", "Cli_Code = '" & StrCodeClient & "'")

         DonnerMontantFrais = DonnerFraisGrille(IntGrilleFrais, CurMontant, IntMonnaie)

      Case (StrModeCalcul = "2") Or (StrModeCalcul = "3" And DLookup("Cli_NewCalculFrais", "CLIENT", "Cli_Code = '" & StrCodeClient & "'") = True)

         DonnerMontantFrais = CurMontant * SngTxLegal * (1 + SngTxTVA) + CurMontantAccessoire * (1 + SngTxTVA)

   End Select

Exit_DonnerMontantFrais:

   Exit Function

Err_DonnerMontantFrais:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : DonnerMontantFrais"

   Resume Exit_DonnerMontantFrais
End Function

Function DonnerFraisGrille(ByVal IntNumeroGrille As Integer, ByVal CurMontant As Currency, ByVal IntMonnaie As Integer) As Currency
    Dim RstFrais As DAO.Recordset

    On Error GoTo Err_DonnerFraisGrille

    Set RstFrais = CurrentDb.OpenRecordset("SELECT * FROM FRAIS WHERE Fra_Grille = " & IntNumeroGrille)

    'On recherche la tranche correspondant au montant de la créance
    Select Case IntMonnaie
      Case cFRANC

         Select Case CurMontant
            Case 0 To 100

                DonnerFraisGrille = RstFrais!Fra_Tranche1_m1

            Case 100.01 To 300

                DonnerFraisGrille = RstFrais!Fra_Tranche2_m1

            Case 300.01 To 500

                DonnerFraisGrille = RstFrais!Fra_Tranche3_m1

            Case 500.01 To 1000

                DonnerFraisGrille = RstFrais!Fra_Tranche4_m1

            Case 1000.01 To 1500

                DonnerFraisGrille = RstFrais!Fra_Tranche5_m1

            Case 1500.01 To 2000

                DonnerFraisGrille = RstFrais!Fra_Tranche6_m1

            Case 2000.01 To 3000

                DonnerFraisGrille = RstFrais!Fra_Tranche7_m1

            Case 3000.01 To 4000

                DonnerFraisGrille = RstFrais!Fra_Tranche8_m1

            Case 4000.01 To 5000

                DonnerFraisGrille = RstFrais!Fra_Tranche9_m1

            Case 5000.01 To 7000

                DonnerFraisGrille = RstFrais!Fra_Tranche10_m1

            Case 7000.01 To 10000

                DonnerFraisGrille = RstFrais!Fra_Tranche11_m1

            Case 10000.01 To 15000

                DonnerFraisGrille = RstFrais!Fra_Tranche12_m1

            Case 15000.01 To 20000

                DonnerFraisGrille = RstFrais!Fra_Tranche13_m1

            Case 20000.01 To 30000

                DonnerFraisGrille = RstFrais!Fra_Tranche14_m1

            Case 30000.01 To 40000

                DonnerFraisGrille = RstFrais!Fra_Tranche15_m1

            Case 40000.01 To 50000

                DonnerFraisGrille = RstFrais!Fra_Tranche16_m1

            Case 50000.01 To 70000

                DonnerFraisGrille = RstFrais!Fra_Tranche17_m1

            Case Is > 70000.01

                DonnerFraisGrille = RstFrais!Fra_Tranche18_m1

         End Select

      Case cEURO

        Select Case Round(CurMontant * DonnerTauxConvertion(cEURO), 2)
            Case 0 To 100

                DonnerFraisGrille = RstFrais!Fra_Tranche1_m2

            Case 100.01 To 300

                DonnerFraisGrille = RstFrais!Fra_Tranche2_m2

            Case 300.01 To 500

                DonnerFraisGrille = RstFrais!Fra_Tranche3_m2

            Case 500.01 To 1000

                DonnerFraisGrille = RstFrais!Fra_Tranche4_m2

            Case 1000.01 To 1500

                DonnerFraisGrille = RstFrais!Fra_Tranche5_m2

            Case 1500.01 To 2000

                DonnerFraisGrille = RstFrais!Fra_Tranche6_m2

            Case 2000.01 To 3000

                DonnerFraisGrille = RstFrais!Fra_Tranche7_m2

            Case 3000.01 To 4000

                DonnerFraisGrille = RstFrais!Fra_Tranche8_m2

            Case 4000.01 To 5000

                DonnerFraisGrille = RstFrais!Fra_Tranche9_m2

            Case 5000.01 To 7000

                DonnerFraisGrille = RstFrais!Fra_Tranche10_m2

            Case 7000.01 To 10000

                DonnerFraisGrille = RstFrais!Fra_Tranche11_m2

            Case 10000.01 To 15000

                DonnerFraisGrille = RstFrais!Fra_Tranche12_m2

            Case 15000.01 To 20000

                DonnerFraisGrille = RstFrais!Fra_Tranche13_m2

            Case 20000.01 To 30000

                DonnerFraisGrille = RstFrais!Fra_Tranche14_m2

            Case 30000.01 To 40000

                DonnerFraisGrille = RstFrais!Fra_Tranche15_m2

            Case 40000.01 To 50000

                DonnerFraisGrille = RstFrais!Fra_Tranche16_m2

            Case 50000.01 To 70000

                DonnerFraisGrille = RstFrais!Fra_Tranche17_m2

            Case Is > 70000.01

                DonnerFraisGrille = RstFrais!Fra_Tranche18_m2

        End Select

    End Select

    RstFrais.Close

Exit_DonnerFraisGrille:

    Set RstFrais = Nothing

    Exit Function

Err_DonnerFraisGrille:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : DonnerFraisGrille"

    Resume Exit_DonnerFraisGrille
End Function

Function DonnerNouveauCodeDebiteur(ByVal StrCode_Client As String) As String
   Dim BolBlocage As Boolean
   Dim StrCode_Debiteur As String
   Dim StrTmpCode As String 'on est obligé de passer par une variable de type variant au cas où tmpcode=null

   On Error GoTo Err_DonnerNouveauCodeDebiteur

   BolBlocage = DLookup("Cli_Bloquer", "CLIENT", "Cli_Code =" & "'" & StrCode_Client & "'")

   Select Case BolBlocage
      Case True

         MsgBox "Client bloquer : création de nouveaux débiteurs impossible !", vbInformation, stSource

         DonnerNouveauCodeDebiteur = ""

      Case False

         StrTmpCode = Nz(DLookup("Cli_Code_der_Déb", "CLIENT", "Cli_Code =" & "'" & StrCode_Client & "'"), "00000")

         DonnerNouveauCodeDebiteur = Format(Int(Right(StrTmpCode, 5)) + 1, "00000")

         DonnerNouveauCodeDebiteur = StrCode_Client & DonnerNouveauCodeDebiteur

   End Select

Exit_DonnerNouveauCodeDebiteur:

    Exit Function

Err_DonnerNouveauCodeDebiteur:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : DonnerNouveauCodeDebiteur"

    Resume Exit_DonnerNouveauCodeDebiteur
End Function

Function GenererRelance(ByVal DatDateDebut As Variant, ByVal DatDateFin As Variant) As Boolean
   Dim IntMes As Integer
   Dim IntMaxRelance As Integer
   Dim IntNivRelance As Integer
   Dim IntGrilleTauxNum As Integer
   Dim IntDelai As Integer
   Dim SngTauxFrais As Single
   Dim BolWhere As Boolean
   Dim StrSQL As String
   Dim StrSQLWhere As String
   Dim WksEspace As DAO.Workspace
   Dim DBBase As DAO.Database
   Dim RstDebiteur As DAO.Recordset
   Dim RstRelance As DAO.Recordset

   On Error GoTo Err_GenererRelance

   Screen.MousePointer = 11

   GenererRelance = True

   StrSQLWhere = "(Déb_Date_proch_rel Is Null)"

   Select Case IsNull(DatDateDebut)
      Case True

      Case False

         StrSQLWhere = StrSQLWhere & " OR (Déb_Date_proch_rel>=" & ConvDate(DatDateDebut)

         BolWhere = True

   End Select

   Select Case IsNull(DatDateFin)
      Case True

         Select Case BolWhere
            Case True

               StrSQLWhere = StrSQLWhere & ")"

            Case False

         End Select

      Case False

         Select Case BolWhere
            Case True

               StrSQLWhere = StrSQLWhere & " AND Déb_Date_proch_rel<=" & ConvDate(DatDateFin) & ")"

            Case False

               StrSQLWhere = StrSQLWhere & " OR (Déb_Date_proch_rel<=" & ConvDate(DatDateFin) & ")"

         End Select

   End Select

   Set WksEspace = DBEngine.Workspaces(0)    ' Cherche la zone de travail par défaut.

   Set DBBase = WksEspace.Databases(0)

   Select Case DCount("TRel_Déb_Code", "TMP_RELANCE")
      Case 0

         WksEspace.BeginTrans

      Case Else

         IntMes = MsgBox("Les anciennes relances n'ont pas toutes été supprimées." & Chr(13) & Chr(10) & "Pour pouvoir éditer les nouvelles relances, toutes les anciennes doivent être supprimées." & Chr(13) & Chr(10) & "SUPPRIMER ?", 3 + 32 + 256)

         Select Case IntMes
            Case 6

               WksEspace.BeginTrans

               DBBase.Execute "DELETE * FROM TMP_RELANCE", dbFailOnError

            Case 7

               GoTo Exit_GenererRelance

            Case Else

               GenererRelance = False

               GoTo Exit_GenererRelance

         End Select

   End Select

   StrSQL = "SELECT CLIENT.Cli_Code, DEBITEUR.Déb_Code, DEBITEUR.Déb_Situation, DEBITEUR.Déb_Code_Enc, DEBITEUR.Déb_niv_relance, DEBITEUR.Déb_Date_proch_rel, DEBITEUR.Déb_Date_der_rel, DEBITEUR.Déb_Frais_m1, DEBITEUR.Déb_Frais_m2, DEBITEUR.Déb_MonnaieEnCours, CODE_GESTION.CG_Code_Action, ADRESSE_DEBITEUR.AD_Tel, ADRESSE_DEBITEUR.AD_Tel2, ADRESSE_DEBITEUR.AD_CodePostal " & _
                 "FROM (CODE_GESTION INNER JOIN (CLIENT INNER JOIN DEBITEUR ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client) ON CODE_GESTION.CG_Code = DEBITEUR.Déb_Situation) INNER JOIN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Code = ADRESSE_DEBITEUR.AD_Code_Débiteur " & _
                 "WHERE CODE_GESTION.CG_Code_Action In (1,2) AND ADRESSE_DEBITEUR.AD_TypeAdresse_ID=DLookUp('Adresse_Courrier','PARAMETRE_GREF') " & IIf(StrSQLWhere = "", "", " AND (" & StrSQLWhere) & ");"

   Set RstDebiteur = DBBase.OpenRecordset(StrSQL, dbOpenDynaset)

   Set RstRelance = DBBase.OpenRecordset("TMP_RELANCE", dbOpenDynaset)

   Select Case RstDebiteur.EOF
      Case True

         WksEspace.CommitTrans

         GoTo Exit_GenererRelance

      Case False

   End Select

   RstDebiteur.MoveLast

   RstDebiteur.MoveFirst

   MajPatientez 0, 0, RstDebiteur.RecordCount

   Do Until RstDebiteur.EOF = True

      MajPatientez RstDebiteur.AbsolutePosition

      IntNivRelance = RstDebiteur("Déb_niv_relance") + IIf(IsNull(RstDebiteur("Déb_Date_der_rel")), 0, 1)

      IntMaxRelance = DLookup("Cli_Nb_Relance", "CLIENT", "Cli_Code = '" & RstDebiteur("Cli_Code") & "'") '*** Nb maxi de relances accepté par le client pour ses débiteurs

      IntGrilleTauxNum = DLookup("Cli_GrilleTaux", "CLIENT", "Cli_Code = '" & RstDebiteur("Cli_Code") & "'") '*** Numéro de la grille de taux du client

      SngTauxFrais = 0

      IntDelai = 0

      RstDebiteur.Edit

      Select Case RstDebiteur("CG_Code_Action")
         Case 1

            Select Case IntNivRelance
               Case Is > IntMaxRelance

                  Select Case IsNull(RstDebiteur("AD_Tel")) And IsNull(RstDebiteur("AD_Tel2"))
                     Case True

                        SngTauxFrais = DLookup("Tx_Frais", "TAUX", "Tx_N° = 7 AND Tx_Grille = " & IntGrilleTauxNum & " AND Tx_Code_Action = 3")

                        RstDebiteur("Déb_niv_relance") = 7

                        RstDebiteur("Déb_Situation") = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action = 3")

                        RstDebiteur("Déb_Code_Enc") = Nz(DLookup("Ens_Code_Enc", "ENCAISSEUR", "Ens_CodePostal='" & Left(RstDebiteur("AD_CodePostal"), 2) & "'"), "XX")

                     Case False

                        IntDelai = DLookup("Tx_Délai", "TAUX", "Tx_N° = 9 AND Tx_Grille = " & IntGrilleTauxNum & " AND Tx_Code_Action = 2")

                        SngTauxFrais = DLookup("Tx_Frais", "TAUX", "Tx_N° = 8 AND Tx_Grille = " & IntGrilleTauxNum & " AND Tx_Code_Action = 2")

                        RstDebiteur("Déb_niv_relance") = 8

                        RstDebiteur("Déb_Situation") = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action = 2")

                  End Select

               Case Else

                  Select Case IntNivRelance
                     Case IntMaxRelance

                        IntDelai = DLookup("Tx_Délai", "TAUX", "Tx_N° = 7 AND Tx_Grille = " & IntGrilleTauxNum & " AND Tx_Code_Action = 3")

                     Case Else

                        IntDelai = DLookup("Tx_Délai", "TAUX", "Tx_N° = " & (IntNivRelance + 1) & " AND Tx_Grille = " & IntGrilleTauxNum & " AND Tx_Code_Action = 1")

                  End Select

                  SngTauxFrais = DLookup("Tx_Frais", "TAUX", "Tx_N° = " & IntNivRelance & " AND Tx_Grille = " & IntGrilleTauxNum & " AND Tx_Code_Action = 1")

                  RstRelance.AddNew

                  RstRelance("TRel_Déb_Code") = RstDebiteur("Déb_Code")

                  RstRelance("TRel_Niv_relance") = IntNivRelance

                  RstRelance.Update

                  RstDebiteur("Déb_niv_relance") = IntNivRelance

            End Select

         Case 2

            SngTauxFrais = DLookup("Tx_Frais", "TAUX", "Tx_N° = 7 AND Tx_Grille = " & IntGrilleTauxNum & " AND Tx_Code_Action = 3")

            RstDebiteur("Déb_niv_relance") = 7

            RstDebiteur("Déb_Situation") = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action = 3")

            RstDebiteur("Déb_Code_Enc") = Nz(DLookup("Ens_Code_Enc", "ENCAISSEUR", "Ens_CodePostal='" & Left(RstDebiteur("AD_CodePostal"), 2) & "'"), "XX")

      End Select

      RstDebiteur("Déb_Date_der_rel") = Date

      Select Case IntDelai
         Case 0

            RstDebiteur("Déb_Date_proch_rel") = Null

         Case Else

            RstDebiteur("Déb_Date_proch_rel") = Date + IntDelai

      End Select

      Select Case SngTauxFrais
         Case 0

         Case Else

            Select Case RstDebiteur("Déb_MonnaieEnCours")
               Case cFRANC

                  RstDebiteur("Déb_Frais_m1") = RstDebiteur("Déb_Frais_m1") + RstDebiteur("Déb_Frais_m1") * SngTauxFrais / 100

                  RstDebiteur("Déb_Frais_m2") = RstDebiteur("Déb_Frais_m1") * DonnerTauxConvertion(RstDebiteur!Déb_MonnaieEnCours)

               Case cEURO

                  RstDebiteur("Déb_Frais_m2") = RstDebiteur("Déb_Frais_m2") + RstDebiteur("Déb_Frais_m2") * SngTauxFrais / 100

                  RstDebiteur("Déb_Frais_m1") = RstDebiteur("Déb_Frais_m2") * DonnerTauxConvertion(RstDebiteur("Déb_MonnaieEnCours"))

            End Select

      End Select

      RstDebiteur.Update

      RstDebiteur.MoveNext

   Loop

   RstDebiteur.Close

   RstRelance.Close

   WksEspace.CommitTrans

Exit_GenererRelance:

   Screen.MousePointer = 0

   DoCmd.Close acForm, "FrmPatientez", acSavePrompt

   Set RstRelance = Nothing

   Set RstDebiteur = Nothing

   Set RstRelance = Nothing

   Set DBBase = Nothing

   Set WksEspace = Nothing

   Exit Function

Err_GenererRelance:

   GenererRelance = False

   WksEspace.Rollback

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : GenererRelance"

   Resume Exit_GenererRelance
End Function

Public Function GenererReglement() As Boolean
   Dim WksEspace As DAO.Workspace
   Dim DBBase As DAO.Database
   Dim RstTable As DAO.Recordset
   Dim RstEncaissement As DAO.Recordset
   Dim RstReglement As DAO.Recordset
   Dim RstReglementDetail As DAO.Recordset
   Dim StrSQL As String
   Dim LngReg_Id As Long
   Dim LngRegD_Id As Long
   Dim StrReg_Cli_Code As String
   Dim StrRegD_Déb_Code As String
   Dim IntRegD_ETy_Mode As Integer
   Dim CurReg_Montant_Encaisser_m1 As Currency
   Dim CurReg_Montant_Encaisser_m2 As Currency
   Dim CurReg_Montant_Direct_m1 As Currency
   Dim CurReg_Montant_Direct_m2 As Currency
   Dim CurRegD_Montant_m1 As Currency
   Dim CurRegD_Montant_m2 As Currency

   On Error GoTo Err_GenererReglement

   Screen.MousePointer = 11

   GenererReglement = True

   Set WksEspace = DBEngine.Workspaces(0) ' Cherche la zone de travail par défaut.

   Set DBBase = WksEspace.Databases(0)      ' Cherche la base de données active.

   WksEspace.BeginTrans  ' Début de la transaction.

   StrSQL = "SELECT CLIENT.Cli_Code, CLIENT.Cli_Pourcentage, DEBITEUR.Déb_Code, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb_Mont_Prin_m2, ENCAISSEMENT.Enc_Id, ENCAISSEMENT_TYPE.ETy_Mode, ENCAISSEMENT.Enc_Prin_m1, ENCAISSEMENT.Enc_Prin_m2, CLIENT.Cli_MonnaieEnCours " & _
                 "FROM (ENCAISSEMENT_TYPE INNER JOIN (CLIENT INNER JOIN (DEBITEUR INNER JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur) ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client) ON ENCAISSEMENT_TYPE.ETy_Code = ENCAISSEMENT.Enc_Type) INNER JOIN TMP_REGLEMENT ON ENCAISSEMENT.Enc_ID = TMP_REGLEMENT.TReg_Enc_ID " & _
                 "WHERE (((TMP_REGLEMENT.TReg_Flag) = True)) ORDER BY CLIENT.Cli_Code, DEBITEUR.Déb_Code, ENCAISSEMENT_TYPE.ETy_Mode DESC;"

   Set RstEncaissement = DBBase.OpenRecordset(StrSQL, dbOpenDynaset)

   Set RstReglement = DBBase.OpenRecordset("SELECT * FROM REGLEMENT")

   Set RstReglementDetail = DBBase.OpenRecordset("SELECT * FROM REGLEMENT_DETAIL")

   Select Case RstEncaissement.EOF
      Case True

         GoTo Exit_GenererReglement

      Case False

   End Select

   RstEncaissement.MoveLast

   RstEncaissement.MoveFirst

   MajPatientez 0, 0, RstEncaissement.RecordCount

   Do While Not RstEncaissement.EOF

      MajPatientez RstEncaissement.AbsolutePosition

      Select Case RstEncaissement("Cli_Code") <> StrReg_Cli_Code Or StrRegD_Déb_Code <> RstEncaissement("Déb_Code") Or IntRegD_ETy_Mode <> RstEncaissement("ETy_Mode")
         Case True

            Select Case IntRegD_ETy_Mode
               Case 1

                  CurReg_Montant_Encaisser_m1 = CurReg_Montant_Encaisser_m1 + Round(CurRegD_Montant_m1, 2)

                  CurReg_Montant_Encaisser_m2 = CurReg_Montant_Encaisser_m2 + Round(CurRegD_Montant_m2, 2)

               Case 2

                  CurReg_Montant_Direct_m1 = CurReg_Montant_Direct_m1 + Round(CurRegD_Montant_m1, 2)

                  CurReg_Montant_Direct_m2 = CurReg_Montant_Direct_m2 + Round(CurRegD_Montant_m2, 2)

            End Select

            Select Case LngReg_Id
               Case 0

               Case Else

                  RstReglement.FindFirst "Reg_Id=" & LngReg_Id

                  RstReglement.Edit

                  RstReglement("Reg_Montant_Encaisser_m1") = CurReg_Montant_Encaisser_m1

                  RstReglement("Reg_Montant_Encaisser_m2") = CurReg_Montant_Encaisser_m2

                  RstReglement("Reg_Montant_Direct_m1") = CurReg_Montant_Direct_m1

                  RstReglement("Reg_Montant_Direct_m2") = CurReg_Montant_Direct_m2

                  RstReglement.Update

            End Select

         Case False

      End Select

      Select Case RstEncaissement("Cli_Code") <> StrReg_Cli_Code
         Case True

            StrReg_Cli_Code = RstEncaissement("Cli_Code")

            CurReg_Montant_Encaisser_m1 = 0

            CurReg_Montant_Encaisser_m2 = 0

            CurReg_Montant_Direct_m1 = 0

            CurReg_Montant_Direct_m2 = 0

            RstReglement.AddNew

            RstReglement("Reg_Date") = Date

            RstReglement("Reg_Cli_Code") = StrReg_Cli_Code

            RstReglement("Reg_Cli_Pourcentage") = RstEncaissement("Cli_Pourcentage")

            RstReglement("Reg_TVA") = DLookup("Taux_TVA", "PARAMETRE_GREF")

            RstReglement("Reg_Montant_Encaisser_m1") = CurReg_Montant_Encaisser_m1

            RstReglement("Reg_Montant_Encaisser_m2") = CurReg_Montant_Encaisser_m2

            RstReglement("Reg_Montant_Direct_m1") = CurReg_Montant_Direct_m1

            RstReglement("Reg_Montant_Direct_m2") = CurReg_Montant_Direct_m2

            RstReglement("Reg_Montant_mo") = RstEncaissement("Cli_MonnaieEnCours")

            RstReglement("Reg_Editer") = True

            LngReg_Id = RstReglement("Reg_Id")

            RstReglement.Update

         Case False

      End Select

      Select Case StrRegD_Déb_Code <> RstEncaissement("Déb_Code") Or IntRegD_ETy_Mode <> RstEncaissement("ETy_Mode")
         Case True

            StrRegD_Déb_Code = RstEncaissement("Déb_Code")

            IntRegD_ETy_Mode = RstEncaissement("ETy_Mode")

            CurRegD_Montant_m1 = RstEncaissement("Enc_Prin_m1")

            CurRegD_Montant_m2 = RstEncaissement("Enc_Prin_m2")

            RstReglementDetail.AddNew

            RstReglementDetail("RegD_Reg_Id") = LngReg_Id

            RstReglementDetail("RegD_Déb_Code") = StrRegD_Déb_Code

            RstReglementDetail("RegD_ETy_Mode") = IntRegD_ETy_Mode

            RstReglementDetail("RegD_Type") = "ACOMPTE"

            RstReglementDetail("RegD_Montant_m1") = CurRegD_Montant_m1

            RstReglementDetail("RegD_Montant_m2") = CurRegD_Montant_m2

            RstReglementDetail("RegD_Montant_mo") = RstEncaissement("Cli_MonnaieEnCours")

            LngRegD_Id = RstReglementDetail("RegD_Id")

            RstReglementDetail.Update

         Case False

            CurRegD_Montant_m1 = CurRegD_Montant_m1 + RstEncaissement("Enc_Prin_m1")

            CurRegD_Montant_m2 = CurRegD_Montant_m2 + RstEncaissement("Enc_Prin_m2")

            RstReglementDetail.FindFirst "RegD_Id=" & LngRegD_Id

            RstReglementDetail.Edit

            RstReglementDetail("RegD_Montant_m1") = CurRegD_Montant_m1

            RstReglementDetail("RegD_Montant_m2") = CurRegD_Montant_m2

            RstReglementDetail.Update

      End Select

      DBBase.Execute "UPDATE ENCAISSEMENT SET ENCAISSEMENT.Enc_RegD_Id = " & LngRegD_Id & ", ENCAISSEMENT.Enc_Situation='R' WHERE Enc_ID=" & RstEncaissement("Enc_Id") & ";", dbFailOnError

      Select Case RstEncaissement("Cli_MonnaieEnCours")
         Case cFRANC

            Set RstTable = DBBase.OpenRecordset("SELECT Sum(Enc_Prin_m1) As Déb_Prin_Enc_m1 FROM ENCAISSEMENT WHERE Enc_RegD_Id<=" & LngRegD_Id & " AND Enc_Code_Débiteur='" & StrRegD_Déb_Code & "'")

            Select Case Round(RstEncaissement("Déb_Mont_Prin_m1"), 2) - Round(Nz(RstTable("Déb_Prin_Enc_m1"), 0), 2)
               Case Is <= 0

                  RstReglementDetail.FindFirst "RegD_Id=" & LngRegD_Id

                  RstReglementDetail.Edit

                  RstReglementDetail("RegD_Type") = "SOLDE"

                  RstReglementDetail.Update

               Case Else

            End Select

            RstTable.Close

         Case cEURO

            Set RstTable = DBBase.OpenRecordset("SELECT Sum(Enc_Prin_m2) As Déb_Prin_Enc_m2 FROM ENCAISSEMENT WHERE Enc_RegD_Id<=" & LngRegD_Id & " AND Enc_Code_Débiteur='" & StrRegD_Déb_Code & "'")

            Select Case Round(RstEncaissement("Déb_Mont_Prin_m2"), 2) - Round(Nz(RstTable("Déb_Prin_Enc_m2"), 0), 2)
               Case Is <= 0

                  RstReglementDetail.FindFirst "RegD_Id=" & LngRegD_Id

                  RstReglementDetail.Edit

                  RstReglementDetail("RegD_Type") = "SOLDE"

                  RstReglementDetail.Update

               Case Else

            End Select

            RstTable.Close

      End Select

      RstEncaissement.MoveNext

   Loop

   Select Case IntRegD_ETy_Mode
      Case 1

         CurReg_Montant_Encaisser_m1 = CurReg_Montant_Encaisser_m1 + Round(CurRegD_Montant_m1, 2)

         CurReg_Montant_Encaisser_m2 = CurReg_Montant_Encaisser_m2 + Round(CurRegD_Montant_m2, 2)

      Case 2

         CurReg_Montant_Direct_m1 = CurReg_Montant_Direct_m1 + Round(CurRegD_Montant_m1, 2)

         CurReg_Montant_Direct_m2 = CurReg_Montant_Direct_m2 + Round(CurRegD_Montant_m2, 2)

   End Select

   Select Case LngReg_Id
      Case 0

      Case Else

         RstReglement.FindFirst "Reg_Id=" & LngReg_Id

         RstReglement.Edit

         RstReglement("Reg_Montant_Encaisser_m1") = CurReg_Montant_Encaisser_m1

         RstReglement("Reg_Montant_Encaisser_m2") = CurReg_Montant_Encaisser_m2

         RstReglement("Reg_Montant_Direct_m1") = CurReg_Montant_Direct_m1

         RstReglement("Reg_Montant_Direct_m2") = CurReg_Montant_Direct_m2

         RstReglement.Update

   End Select

   RstEncaissement.Close

   RstReglementDetail.Close

   RstReglement.Close

   StrSQL = "DELETE * FROM TMP_REGLEMENT;"

   DBBase.Execute StrSQL, dbFailOnError

   WksEspace.CommitTrans

Exit_GenererReglement:

   Screen.MousePointer = 0

   DoCmd.Close acForm, "FrmPatientez", acSavePrompt

   Set RstEncaissement = Nothing

   Set RstReglementDetail = Nothing

   Set RstReglement = Nothing

   Set RstTable = Nothing

   Set DBBase = Nothing

   Set WksEspace = Nothing

   Exit Function

Err_GenererReglement:

   GenererReglement = False

   WksEspace.Rollback

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : GenererReglement"

   Resume Exit_GenererReglement
End Function

Public Function GenererFacture(ByVal SngHonoraire As Single, ByVal StrFactureNumero As String) As Boolean
   Dim WksEspace As DAO.Workspace
   Dim DBBase As DAO.Database
   Dim RstEncaissement As DAO.Recordset
   Dim RstFacture As DAO.Recordset
   Dim RstFactureDetail As DAO.Recordset
   Dim StrSQL As String
   Dim LngFac_Id As Long
   Dim LngFacD_Id As Long
   Dim StrFac_Cli_Code As String
   Dim StrFacD_Déb_Code As String
   Dim IntFacD_ETy_Mode As Integer
   Dim CurFac_Montant_Encaisser_m1 As Currency
   Dim CurFac_Montant_Encaisser_m2 As Currency
   Dim CurFac_Montant_Direct_m1 As Currency
   Dim CurFac_Montant_Direct_m2 As Currency
   Dim CurFacD_Montant_m1 As Currency
   Dim CurFacD_Montant_m2 As Currency

   On Error GoTo Err_GenererFacture

   Screen.MousePointer = 11

   GenererFacture = True

   Set WksEspace = DBEngine.Workspaces(0) ' Cherche la zone de travail par défaut.

   Set DBBase = WksEspace.Databases(0)      ' Cherche la base de données active.

   WksEspace.BeginTrans  ' Début de la transaction.

   StrSQL = "SELECT CLIENT.Cli_Code, DEBITEUR.Déb_Code, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb_Mont_Prin_m2, ENCAISSEMENT.Enc_Id, ENCAISSEMENT_TYPE.ETy_Mode, ENCAISSEMENT.Enc_Prin_m1, ENCAISSEMENT.Enc_Prin_m2, CLIENT.Cli_MonnaieEnCours " & _
                 "FROM (ENCAISSEMENT_TYPE INNER JOIN (CLIENT INNER JOIN (DEBITEUR INNER JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur) ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client) ON ENCAISSEMENT_TYPE.ETy_Code = ENCAISSEMENT.Enc_Type) INNER JOIN TMP_Facture ON ENCAISSEMENT.Enc_ID = TMP_FACTURE.TFac_Enc_ID " & _
                 "WHERE (((TMP_FACTURE.TFac_Flag) = True)) ORDER BY CLIENT.Cli_Code, DEBITEUR.Déb_Code, ENCAISSEMENT_TYPE.ETy_Mode DESC;"

   Set RstEncaissement = DBBase.OpenRecordset(StrSQL, dbOpenDynaset)

   Set RstFacture = DBBase.OpenRecordset("SELECT * FROM FACTURE")

   Set RstFactureDetail = DBBase.OpenRecordset("SELECT * FROM FACTURE_DETAIL")

   RstEncaissement.MoveLast

   RstEncaissement.MoveFirst

   MajPatientez 0, 0, RstEncaissement.RecordCount

   Do While Not RstEncaissement.EOF

      MajPatientez RstEncaissement.AbsolutePosition

      Select Case RstEncaissement("Cli_Code") <> StrFac_Cli_Code Or StrFacD_Déb_Code <> RstEncaissement("Déb_Code") Or IntFacD_ETy_Mode <> RstEncaissement("ETy_Mode")
         Case True

            Select Case IntFacD_ETy_Mode
               Case 1

                  CurFac_Montant_Encaisser_m1 = CurFac_Montant_Encaisser_m1 + Round(CurFacD_Montant_m1, 2)

                  CurFac_Montant_Encaisser_m2 = CurFac_Montant_Encaisser_m2 + Round(CurFacD_Montant_m2, 2)

               Case 2

                  CurFac_Montant_Direct_m1 = CurFac_Montant_Direct_m1 + Round(CurFacD_Montant_m1, 2)

                  CurFac_Montant_Direct_m2 = CurFac_Montant_Direct_m2 + Round(CurFacD_Montant_m2, 2)

            End Select

            Select Case LngFac_Id
               Case 0

               Case Else

                  RstFacture.FindFirst "Fac_Id=" & LngFac_Id

                  RstFacture.Edit

                  RstFacture("Fac_Montant_Encaisser_m1") = CurFac_Montant_Encaisser_m1

                  RstFacture("Fac_Montant_Encaisser_m2") = CurFac_Montant_Encaisser_m2

                  RstFacture("Fac_Montant_Direct_m1") = CurFac_Montant_Direct_m1

                  RstFacture("Fac_Montant_Direct_m2") = CurFac_Montant_Direct_m2

                  RstFacture.Update

            End Select

         Case False

      End Select

      Select Case RstEncaissement("Cli_Code") <> StrFac_Cli_Code
         Case True

            StrFac_Cli_Code = RstEncaissement("Cli_Code")

            CurFac_Montant_Encaisser_m1 = 0

            CurFac_Montant_Encaisser_m2 = 0

            CurFac_Montant_Direct_m1 = 0

            CurFac_Montant_Direct_m2 = 0

            RstFacture.AddNew

            RstFacture("Fac_Date") = Date

            RstFacture("Fac_CodeCli") = StrFac_Cli_Code

            RstFacture("Fac_Num") = Format(StrFactureNumero, "00000000")

            RstFacture("Fac_Honoraire") = SngHonoraire

            RstFacture("Fac_TVA") = DLookup("Taux_TVA", "PARAMETRE_GREF")

            RstFacture("Fac_Montant_Encaisser_m1") = CurFac_Montant_Encaisser_m1

            RstFacture("Fac_Montant_Encaisser_m2") = CurFac_Montant_Encaisser_m2

            RstFacture("Fac_Montant_Direct_m1") = CurFac_Montant_Direct_m1

            RstFacture("Fac_Montant_Direct_m2") = CurFac_Montant_Direct_m2

            RstFacture("Fac_Montant_mo") = RstEncaissement("Cli_MonnaieEnCours")

            RstFacture("Fac_Editer") = True

            LngFac_Id = RstFacture("Fac_Id")

            RstFacture.Update

            StrFactureNumero = StrFactureNumero + 1

         Case False

      End Select

      Select Case StrFacD_Déb_Code <> RstEncaissement("Déb_Code") Or IntFacD_ETy_Mode <> RstEncaissement("ETy_Mode")
         Case True

            StrFacD_Déb_Code = RstEncaissement("Déb_Code")

            IntFacD_ETy_Mode = RstEncaissement("ETy_Mode")

            CurFacD_Montant_m1 = RstEncaissement("Enc_Prin_m1")

            CurFacD_Montant_m2 = RstEncaissement("Enc_Prin_m2")

            RstFactureDetail.AddNew

            RstFactureDetail("FacD_Fac_Id") = LngFac_Id

            RstFactureDetail("FacD_Déb_Code") = StrFacD_Déb_Code

            RstFactureDetail("FacD_ETy_Mode") = IntFacD_ETy_Mode

            RstFactureDetail("FacD_Montant_m1") = CurFacD_Montant_m1

            RstFactureDetail("FacD_Montant_m2") = CurFacD_Montant_m2

            RstFactureDetail("FacD_Montant_mo") = RstEncaissement("Cli_MonnaieEnCours")

            LngFacD_Id = RstFactureDetail("FacD_Id")

            RstFactureDetail.Update

         Case False

            CurFacD_Montant_m1 = CurFacD_Montant_m1 + RstEncaissement("Enc_Prin_m1")

            CurFacD_Montant_m2 = CurFacD_Montant_m2 + RstEncaissement("Enc_Prin_m2")

            RstFactureDetail.FindFirst "FacD_Id=" & LngFacD_Id

            RstFactureDetail.Edit

            RstFactureDetail("FacD_Montant_m1") = CurFacD_Montant_m1

            RstFactureDetail("FacD_Montant_m2") = CurFacD_Montant_m2

            RstFactureDetail.Update

      End Select

      DBBase.Execute "UPDATE ENCAISSEMENT SET ENCAISSEMENT.Enc_FacD_Id = " & LngFacD_Id & ", ENCAISSEMENT.Enc_Situation='F' WHERE Enc_ID=" & RstEncaissement("Enc_Id") & ";", dbFailOnError

      RstEncaissement.MoveNext

   Loop

   Select Case IntFacD_ETy_Mode
      Case 1

         CurFac_Montant_Encaisser_m1 = CurFac_Montant_Encaisser_m1 + Round(CurFacD_Montant_m1, 2)

         CurFac_Montant_Encaisser_m2 = CurFac_Montant_Encaisser_m2 + Round(CurFacD_Montant_m2, 2)

      Case 2

         CurFac_Montant_Direct_m1 = CurFac_Montant_Direct_m1 + Round(CurFacD_Montant_m1, 2)

         CurFac_Montant_Direct_m2 = CurFac_Montant_Direct_m2 + Round(CurFacD_Montant_m2, 2)

   End Select

   Select Case LngFac_Id
      Case 0

      Case Else

         RstFacture.FindFirst "Fac_Id=" & LngFac_Id

         RstFacture.Edit

         RstFacture("Fac_Montant_Encaisser_m1") = CurFac_Montant_Encaisser_m1

         RstFacture("Fac_Montant_Encaisser_m2") = CurFac_Montant_Encaisser_m2

         RstFacture("Fac_Montant_Direct_m1") = CurFac_Montant_Direct_m1

         RstFacture("Fac_Montant_Direct_m2") = CurFac_Montant_Direct_m2

         RstFacture.Update

   End Select

   RstEncaissement.Close

   RstFactureDetail.Close

   RstFacture.Close

   StrSQL = "DELETE * FROM TMP_FACTURE;"

   DBBase.Execute StrSQL, dbFailOnError

   WksEspace.CommitTrans

Exit_GenererFacture:

   Screen.MousePointer = 0

   DoCmd.Close acForm, "FrmPatientez", acSavePrompt

   Set RstEncaissement = Nothing

   Set RstFactureDetail = Nothing

   Set RstFacture = Nothing

   Set DBBase = Nothing

   Set WksEspace = Nothing

   Exit Function

Err_GenererFacture:

   GenererFacture = False

   WksEspace.Rollback

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : GenererFacture"

   Resume Exit_GenererFacture
End Function

Function GenererIrrecouvrable(ByVal StrCli_Code As Variant, ByVal DatDateFin As Variant) As Boolean
   Dim IntMes As Integer
   Dim BolWhere As Boolean
   Dim StrSQLWhere As String
   Dim StrSQL As String
   Dim WksEspace As DAO.Workspace
   Dim DBBase As DAO.Database

   On Error GoTo Err_GenererIrrecouvrable

   Screen.MousePointer = 11

   GenererIrrecouvrable = True

   Select Case IsNull(StrCli_Code)
      Case True

         GoTo Exit_GenererIrrecouvrable

      Case False

   End Select

   StrSQLWhere = "(Cli_Code='" & StrCli_Code & "' "

   Select Case IsNull(DatDateFin)
      Case True

         StrSQLWhere = StrSQLWhere & "AND (Déb_Date_Cré Is Null))"

      Case False

         StrSQLWhere = StrSQLWhere & "AND ((Déb_Date_Cré Is Null) OR Déb_Date_Cré<=" & ConvDate(DatDateFin) & "))"

   End Select

   Set WksEspace = DBEngine.Workspaces(0)    ' Cherche la zone de travail par défaut.

   Set DBBase = WksEspace.Databases(0)

   WksEspace.BeginTrans

   StrSQL = "INSERT INTO [TMP_IRRECOUVRABLE] (TIrr_Déb_Code,TIrr_Deja_Editer) " & _
                 "SELECT DEBITEUR.Déb_Code, Not IsNull([Déb_Date_Der_Rel]) AS TIrr_Deja_Editer " & _
                 "FROM CLIENT INNER JOIN (CODE_GESTION INNER JOIN DEBITEUR ON CODE_GESTION.CG_Code = DEBITEUR.Déb_Situation) ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client " & _
                 "WHERE (CODE_GESTION.CG_Code_Action = 6) AND " & StrSQLWhere

   DBBase.Execute StrSQL, dbFailOnError

   WksEspace.CommitTrans

Exit_GenererIrrecouvrable:

   Screen.MousePointer = 0

   DoCmd.Close acForm, "FrmPatientez", acSavePrompt

   Set DBBase = Nothing

   Set WksEspace = Nothing

   Exit Function

Err_GenererIrrecouvrable:

   GenererIrrecouvrable = False

   WksEspace.Rollback

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : GenererIrrecouvrable"

   Resume Exit_GenererIrrecouvrable
End Function
