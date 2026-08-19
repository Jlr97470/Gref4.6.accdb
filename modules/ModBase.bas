Attribute VB_Name = "ModBase"
Option Compare Database
Option Explicit

Public Function VerifierBase() As Boolean
   Dim FsoSystemFichier As New Scripting.FileSystemObject
   Dim StrPathMDB As String
   Dim StrFichMDB As String
   Dim StrFichArchMDB As String
   Dim TdfTable As DAO.TableDef

   On Error GoTo Err_VerifierBase

   VerifierBase = True

   CurrentDb.Properties("AppTitle") = Nz(DLookup("Bas_Nom", "BASE"))

   StrPathMDB = Nz(DLookup("Bas_Path", "BASE"))

   StrFichMDB = Nz(DLookup("Bas_MDB", "BASE"))

   StrFichArchMDB = DLookup("Bas_Archive_MDB", "BASE")

   Select Case FsoSystemFichier.FileExists(StrPathMDB & StrFichMDB)
      Case True

      Case False

         VerifierBase = False

         MsgBox "Base " & StrPathMDB & StrFichMDB & " Inexistante !!"

         GoTo Exit_VerifierBase

   End Select

   Select Case FsoSystemFichier.FileExists(StrPathMDB & StrFichArchMDB)
      Case True

      Case False

         VerifierBase = False

         MsgBox "Base " & StrPathMDB & StrFichArchMDB & " Inexistante !!"

         GoTo Exit_VerifierBase

   End Select

   For Each TdfTable In CurrentDb.TableDefs

      Select Case TdfTable.Connect
         Case ""

         Case Else

            Select Case DLookup("TblNumero", "TABLE", "TblNom='" & TdfTable.Name & "'")
               Case 1

                  SauveTableInformationConnexion TdfTable, "", "", StrPathMDB, StrFichMDB

               Case 2

                  SauveTableInformationConnexion TdfTable, "", "", StrPathMDB, StrFichArchMDB

            End Select

      End Select

   Next

   CurrentDb.TableDefs.Refresh

'   MajBDD46

Exit_VerifierBase:

   Exit Function

Err_VerifierBase:

   VerifierBase = False

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : VerifierBase"

   Resume Exit_VerifierBase
End Function

Public Function LectureTableInformationConnexion(ByVal TdfTable As TableDef, ByRef StrType As String, ByRef StrParametre, ByRef StrChemin As String, ByRef StrFichier As String) As Boolean
   Dim StrConnexion As String
   Dim IntConnexionLongueur As Integer
   Dim StrCheminFichier As String
   Dim IntCheminFichierLongueur As Integer

   On Error GoTo Err_LectureTableInformationConnexion

   StrConnexion = TdfTable.Connect

   IntConnexionLongueur = Len(StrConnexion)

   Select Case IntConnexionLongueur
      Case vbEmpty

         StrType = vbNullString

         StrParametre = vbNullString

         StrChemin = vbNullString

         StrFichier = vbNullString

         LectureTableInformationConnexion = False

         GoTo Exit_LectureTableInformationConnexion

      Case Else

   End Select

   StrType = Mid(StrConnexion, 1, InStr(1, StrConnexion, ";") - 1)

   StrCheminFichier = Mid(StrConnexion, InStr(1, StrConnexion, "DATABASE=") + 9)

   StrParametre = RemplaceChr(RemplaceChr(StrConnexion, StrType & ";", vbNullString), "DATABASE=" & StrCheminFichier, vbNullString)

   IntCheminFichierLongueur = Len(StrCheminFichier)

   Select Case IntCheminFichierLongueur
      Case vbEmpty

         StrChemin = vbNullString

         StrFichier = vbNullString

         LectureTableInformationConnexion = False

      Case Else

         Select Case StrType
            Case vbNullString, "Excel 5.0"

               StrChemin = Mid(StrCheminFichier, 1, InStrRev(StrCheminFichier, "\") - 1)

               StrFichier = Mid(StrCheminFichier, InStrRev(StrCheminFichier, "\") + 1)

            Case "Text"

               StrChemin = StrCheminFichier

               StrFichier = TdfTable.SourceTableName

            Case Else

         End Select

         LectureTableInformationConnexion = True

   End Select

Exit_LectureTableInformationConnexion:

   Exit Function

Err_LectureTableInformationConnexion:

   LectureTableInformationConnexion = False

   MsgBox Err.Number & " " & Err.description, , stSource & " : LectureTableInformationConnexion"

   Resume Exit_LectureTableInformationConnexion
End Function

Public Function SauveTableInformationConnexion(ByRef TdfTable As DAO.TableDef, ByRef StrType As String, ByRef StrParametre, ByRef StrChemin As String, ByRef StrFichier As String) As Boolean
   Dim strConnect As String

   On Error GoTo Err_SauveTableInformationConnexion

   SauveTableInformationConnexion = True

   Select Case StrType
      Case vbNullString

         strConnect = ";DATABASE=" & StrChemin & StrFichier

      Case "Excel 5.0"

         strConnect = StrType & ";" & StrParametre & ";DATABASE=" & StrChemin & StrFichier

      Case "Text"

         strConnect = StrType & ";" & StrParametre & ";DATABASE=" & StrChemin

      Case Else

   End Select

   Select Case TdfTable.Connect
      Case strConnect

      Case Else

         TdfTable.Connect = strConnect

         TdfTable.RefreshLink

   End Select

Exit_SauveTableInformationConnexion:

   Exit Function

Err_SauveTableInformationConnexion:

   SauveTableInformationConnexion = False

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : SauveTableInformationConnexion"

   Resume Exit_SauveTableInformationConnexion
End Function

Public Function MajBDD46() As Boolean
   Dim FsoFichierSystem As New Scripting.FileSystemObject
   Dim FilFichier As Scripting.File
   Dim DBBaseSource As DAO.Database
   Dim DBBaseDestination As DAO.Database
   Dim TdfTableSource As DAO.TableDef
   Dim TdfTableDestination As DAO.TableDef
   Dim RstTableSource As DAO.Recordset
   Dim RstTableDestination As DAO.Recordset
   Dim RstTableDestination1 As DAO.Recordset
   Dim RstTable As DAO.Recordset
   Dim FldChamp As DAO.Field
   Dim StrSQL As String
   Dim StrPath As String
   Dim StrFichierSource As String
   Dim StrFichierDestination As String
   Dim IntFichierNumero As Integer
   Dim LngReg_Id As Long
   Dim DatReg_Date As Date
   Dim StrReg_Cli_Code As String
   Dim SngReg_Cli_Pourcentage As Single
   Dim IntReg_Montant_mo As Integer
   Dim CurReg_Montant_Encaisser_m1 As Currency
   Dim CurReg_Montant_Encaisser_m2 As Currency
   Dim CurReg_Montant_Direct_m1 As Currency
   Dim CurReg_Montant_Direct_m2 As Currency
   Dim LngRegD_Id As Long
   Dim StrRegD_Déb_Code As String
   Dim IntRegD_ETy_Mode As Integer
   Dim CurRegD_Montant_m1 As Currency
   Dim CurRegD_Montant_m2 As Currency
   Dim LngEnc_Fac_Id As Long
   Dim LngFac_Id As Long
   Dim StrFacMois As String
   Dim IntFacNum As Integer
   Dim IntFac_Montant_mo As Integer
   Dim CurFac_Montant_Encaisser_m1 As Currency
   Dim CurFac_Montant_Encaisser_m2 As Currency
   Dim CurFac_Montant_Direct_m1 As Currency
   Dim CurFac_Montant_Direct_m2 As Currency
   Dim LngFacD_Id As Long
   Dim StrFacD_Déb_Code As String
   Dim IntFacD_ETy_Mode As Integer
   Dim CurFacD_Montant_m1 As Currency
   Dim CurFacD_Montant_m2 As Currency

   On Error GoTo Err_MajBDD46

   MajBDD46 = True

   StrPath = DLookup("Bas_Path", "BASE")

   StrFichierSource = "TGREF2003OLD.MDB"

   StrFichierDestination = DLookup("Bas_MDB", "BASE")

   Select Case FsoFichierSystem.FileExists(StrPath & StrFichierSource)
      Case True

      Case False

         GoTo Exit_MajBDD46

   End Select

   IntFichierNumero = FreeFile

   Open StrPath & "MajBDD.LOG" For Output As #IntFichierNumero

   Print #IntFichierNumero, "Maj 4.6 : " & Now()

   Set DBBaseSource = DBEngine.OpenDatabase(StrPath & StrFichierSource)

   Set DBBaseDestination = DBEngine.OpenDatabase(StrPath & StrFichierDestination)

   Set TdfTableSource = DBBaseSource.TableDefs("Paramètres GREF")

   Set TdfTableDestination = DBBaseDestination.TableDefs("PARAMETRE_GREF")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         Select Case FldChamp.Name
            Case "Mot_de_Passe", "TauxConvertion_m1m2", "Taux_TVA", "Adresse_Courrier", "Code_Sous_Traitant", "Situation_Déb_Défaut", "Mode_Calcul_Frais", "Tx_Interet_legal", "Frais_Accessoires", "Frais_Déplacement_m1", "Frais_Déplacement_m2", "Site"

               RstTableDestination(FldChamp.Name) = FldChamp.Value

            Case Else

         End Select

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " PARAMETRE_GREF:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("ADRESSE_TYPE")

   Set TdfTableDestination = DBBaseDestination.TableDefs("ADRESSE_TYPE")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " ADRESSE_TYPE:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("FRAIS_GRILLE")

   Set TdfTableDestination = DBBaseDestination.TableDefs("FRAIS_GRILLE")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " FRAIS_GRILLE:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("FRAIS")

   Set TdfTableDestination = DBBaseDestination.TableDefs("FRAIS")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = Round(FldChamp.Value, 2)

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " FRAIS:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("FRAIS_TRANS_GRILLE")

   Set TdfTableDestination = DBBaseDestination.TableDefs("FRAIS_TRANS_GRILLE")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " FRAIS_TRANS_GRILLE:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("FRAIS_TRANS")

   Set TdfTableDestination = DBBaseDestination.TableDefs("FRAIS_TRANS")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " FRAIS_TRANS:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

'   Set TdfTableSource = DBBaseSource.TableDefs("TAUX_GRILLE")
'
'   Set TdfTableDestination = DBBaseDestination.TableDefs("TAUX_GRILLE")
'
'   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)
'
'   Set RstTableDestination = TdfTableDestination.OpenRecordset()
'
'   RstTableSource.MoveLast
'
'   RstTableSource.MoveFirst
'
'   MajPatientez 0, 0, RstTableSource.RecordCount
'
'   Do While Not RstTableSource.EOF
'
'      MajPatientez RstTableSource.AbsolutePosition
'
'      RstTableDestination.AddNew
'
'      For Each FldChamp In RstTableSource.Fields
'
'         RstTableDestination(FldChamp.Name) = FldChamp.Value
'
'      Next
'
'      RstTableDestination.Update
'
'      RstTableSource.MoveNext
'
'   Loop
'
'   Print #IntFichierNumero, "TAUX_GRILLE:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount
'
'   RstTableSource.Close
'
'   RstTableDestination.Close
'
'   StrSQL = "SELECT Tx_N° AS TxR_N°, Tx_Nom AS TxR_Nom FROM TAUX GROUP BY Tx_N°, Tx_Nom;"
'
'   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)
'
'   Set TdfTableDestination = DBBaseDestination.TableDefs("TAUX_RELANCE")
'
'   Set RstTableDestination = TdfTableDestination.OpenRecordset()
'
'   RstTableSource.MoveLast
'
'   RstTableSource.MoveFirst
'
'   MajPatientez 0, 0, RstTableSource.RecordCount
'
'   Do While Not RstTableSource.EOF
'
'      MajPatientez RstTableSource.AbsolutePosition
'
'      RstTableDestination.AddNew
'
'      For Each FldChamp In RstTableSource.Fields
'
'         RstTableDestination(FldChamp.Name) = FldChamp.Value
'
'      Next
'
'      RstTableDestination.Update
'
'      RstTableSource.MoveNext
'
'   Loop
'
'   Print #IntFichierNumero, "TAUX_RELANCE:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount
'
'   RstTableSource.Close
'
'   RstTableDestination.Close
'
'   Set TdfTableSource = DBBaseSource.TableDefs("TAUX")
'
'   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)
'
'   Set TdfTableDestination = DBBaseDestination.TableDefs("TAUX")
'
'   Set RstTableDestination = TdfTableDestination.OpenRecordset()
'
'   RstTableSource.MoveLast
'
'   RstTableSource.MoveFirst
'
'   MajPatientez 0, 0, RstTableSource.RecordCount
'
'   Do While Not RstTableSource.EOF
'
'      MajPatientez RstTableSource.AbsolutePosition
'
'      RstTableDestination.AddNew
'
'      For Each FldChamp In RstTableSource.Fields
'
'         Select Case FldChamp.Name
'            Case "Tx_Nom"
'
'            Case Else
'
'               RstTableDestination(FldChamp.Name) = FldChamp.Value
'
'         End Select
'
'      Next
'
'      RstTableDestination.Update
'
'      RstTableSource.MoveNext
'
'   Loop
'
'   Print #IntFichierNumero, "TAUX:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount
'
'   RstTableSource.Close
'
'   RstTableDestination.Close
'
'   Set TdfTableSource = DBBaseSource.TableDefs("CODE_ACTION")
'
'   Set TdfTableDestination = DBBaseDestination.TableDefs("CODE_ACTION")
'
'   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)
'
'   Set RstTableDestination = TdfTableDestination.OpenRecordset()
'
'   RstTableSource.MoveLast
'
'   RstTableSource.MoveFirst
'
'   MajPatientez 0, 0, RstTableSource.RecordCount
'
'   Do While Not RstTableSource.EOF
'
'      MajPatientez RstTableSource.AbsolutePosition
'
'      RstTableDestination.AddNew
'
'      For Each FldChamp In RstTableSource.Fields
'
'         RstTableDestination(FldChamp.Name) = FldChamp.Value
'
'      Next
'
'      RstTableDestination.Update
'
'      RstTableSource.MoveNext
'
'   Loop
'
'   Print #IntFichierNumero, "CODE_ACTION:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount
'
'   RstTableSource.Close
'
'   RstTableDestination.Close
'
'   Set TdfTableSource = DBBaseSource.TableDefs("CODE_GESTION")
'
'   Set TdfTableDestination = DBBaseDestination.TableDefs("CODE_GESTION")
'
'   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)
'
'   Set RstTableDestination = TdfTableDestination.OpenRecordset()
'
'   RstTableSource.MoveLast
'
'   RstTableSource.MoveFirst
'
'   MajPatientez 0, 0, RstTableSource.RecordCount
'
'   Do While Not RstTableSource.EOF
'
'      MajPatientez RstTableSource.AbsolutePosition
'
'      RstTableDestination.AddNew
'
'      For Each FldChamp In RstTableSource.Fields
'
'         RstTableDestination(FldChamp.Name) = FldChamp.Value
'
'      Next
'
'      RstTableDestination.Update
'
'      RstTableSource.MoveNext
'
'   Loop
'
'   Print #IntFichierNumero, "CODE_GESTION:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount
'
'   RstTableSource.Close
'
'   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("ENCAISSEMENT_MODE")

   Set TdfTableDestination = DBBaseDestination.TableDefs("ENCAISSEMENT_MODE")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " ENCAISSEMENT_MODE:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("ENCAISSEMENT_TYPE")

   Set TdfTableDestination = DBBaseDestination.TableDefs("ENCAISSEMENT_TYPE")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " ENCAISSEMENT_TYPE:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("MOTIF_IMPAYE_GRILLE")

   Set TdfTableDestination = DBBaseDestination.TableDefs("MOTIF_IMPAYE_GRILLE")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " MOTIF_IMPAYE_GRILLE:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("MOTIF_IMPAYE")

   Set TdfTableDestination = DBBaseDestination.TableDefs("MOTIF_IMPAYE")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " MOTIF_IMPAYE:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("GROUPE")

   Set TdfTableDestination = DBBaseDestination.TableDefs("GROUPE")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " GROUPE:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("REPRESENTANT")

   Set TdfTableDestination = DBBaseDestination.TableDefs("REPRESENTANT")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " REPRESENTANT:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("CLIENT")

   Set TdfTableDestination = DBBaseDestination.TableDefs("CLIENT")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         Select Case FldChamp.Name
            Case "Cli_Code"

               RstTableDestination("Cli_Code") = RstTableSource("Cli_Code")

               Select Case True
                  Case RstTableDestination("Cli_Code") Like "7*"

                     RstTableDestination("Cli_Groupe") = 5

                  Case RstTableDestination("Cli_Code") Like "C*"

                     RstTableDestination("Cli_Groupe") = 6

                  Case RstTableDestination("Cli_Code") Like "AU*"

                     RstTableDestination("Cli_Groupe") = 7

                  Case RstTableDestination("Cli_Code") Like "ED*"

                     RstTableDestination("Cli_Groupe") = 8

               End Select

            Case "Cli_Date_der_reg", "Cli_Der_numFac", "Cli_Flag1", "Cli_Groupe"

            Case "Cli_Pourcentage"

               RstTableDestination(FldChamp.Name) = IIf(IsNull(FldChamp.Value), 0, FldChamp.Value)

            Case Else

               RstTableDestination(FldChamp.Name) = FldChamp.Value

         End Select

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " CLIENT:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrSQL = "SELECT ADRESSE_CLIENT.AC_Code_Client, ADRESSE_CLIENT.AC_TypeAdresse_ID, Last(ADRESSE_CLIENT.AC_GenreComplémentaire) AS AC_GenreComplémentaire, Last(ADRESSE_CLIENT.AC_NomComplémentaire) AS AC_NomComplémentaire, Last(ADRESSE_CLIENT.AC_Adr1) AS AC_Adr1, Last(ADRESSE_CLIENT.AC_Adr2) AS AC_Adr2, Last(ADRESSE_CLIENT.AC_Adr3) AS AC_Adr3, Last(ADRESSE_CLIENT.AC_Adr4) AS AC_Adr4, Last(ADRESSE_CLIENT.AC_CodePostal) AS AC_CodePostal, Last(ADRESSE_CLIENT.AC_Ville) AS AC_Ville, Last(ADRESSE_CLIENT.AC_Pays) AS AC_Pays, Last(ADRESSE_CLIENT.AC_Tel) AS AC_Tel, Last(ADRESSE_CLIENT.AC_Tel2) AS AC_Tel2, Last(ADRESSE_CLIENT.AC_Fax) AS AC_Fax, Last(ADRESSE_CLIENT.AC_email) AS AC_email, Last(ADRESSE_CLIENT.AC_Note) AS AC_Note " & _
                  "FROM ADRESSE_CLIENT WHERE (((ADRESSE_CLIENT.AC_TypeAdresse_ID) Is Not Null)) GROUP BY ADRESSE_CLIENT.AC_Code_Client, ADRESSE_CLIENT.AC_TypeAdresse_ID;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("ADRESSE_CLIENT")

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " ADRESSE_CLIENT:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrSQL = "SELECT HIST_CONTRAT.HC_Code_Client, HIST_CONTRAT.HC_Date_Contrat, Last(HIST_CONTRAT.HC_Mont_Hon_m1) AS HC_Mont_Hon_m1, Last(HIST_CONTRAT.HC_Mont_Hon_m2) AS HC_Mont_Hon_m2, Last(HIST_CONTRAT.HC_Mont_Hon_mo) AS HC_Mont_Hon_mo, Last(HIST_CONTRAT.HC_Pourcentage) AS HC_Pourcentage " & _
                 "FROM HIST_CONTRAT GROUP BY HIST_CONTRAT.HC_Code_Client, HIST_CONTRAT.HC_Date_Contrat HAVING ((Not (HIST_CONTRAT.HC_Date_Contrat) Is Null));"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("CLIENT_CONTRAT")

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination("C" & Mid(FldChamp.Name, 2)) = Nz(FldChamp.Value, RstTableDestination("C" & Mid(FldChamp.Name, 2)).DefaultValue)

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " CLIENT_CONTRAT:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrSQL = "SELECT DEBITEUR.*, (SELECT First(AD_CodePostal) FROM ADRESSE_DEBITEUR WHERE AD_Code_Débiteur=Déb_Code;) AS AD_CodePostal FROM DEBITEUR;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("DEBITEUR")

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         Select Case FldChamp.Name
            Case "Déb_Num_Chèque"

               Select Case IsDate(FldChamp.Value)
                  Case True

                     RstTableDestination("Déb_Date_Chèque") = RstTableSource("Déb_Num_Chèque")

                  Case False

                     RstTableDestination("Déb_Num_Chèque") = RstTableSource("Déb_Num_Chèque")

               End Select

               Select Case IsDate(RstTableSource("Déb_Date_Chèque"))
                  Case True

                     RstTableDestination("Déb_Date_Chèque") = RstTableSource("Déb_Date_Chèque")

                  Case False

               End Select

               Select Case RstTableSource("Déb_Code_Client")
                  Case "88961"

                     RstTableDestination("Déb_Num_Chèque") = RstTableSource("Déb_Réf1_Client")

                  Case Else

               End Select

            Case "Déb_Situation"

               RstTableDestination("Déb_Situation") = RstTableSource("Déb_Situation")

               RstTableDestination("Déb_Code_Enc") = RstTableSource("Déb_Code_Enc")

               Select Case DLookup("CG_Code_Action", "CODE_GESTION", "CG_Code='" & RstTableSource("Déb_Situation") & "'")
                  Case 1

                     Select Case RstTableSource("Déb_niv_relance")
                        Case 1, 2, 3, 4, 5, 6

                           Select Case RstTableSource("Déb_Date_proch_rel") >= "01/01/2005" Or IsNull(RstTableSource("Déb_Date_proch_rel"))
                              Case True

                                 RstTableDestination("Déb_niv_relance") = RstTableSource("Déb_niv_relance")

                                 RstTableDestination("Déb_Date_proch_rel") = RstTableSource("Déb_Date_proch_rel")

                                 RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                              Case False

                                 RstTableDestination("Déb_Situation") = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action= 3 ")

                                 RstTableDestination("Déb_niv_relance") = 7

                                 RstTableDestination("Déb_Date_proch_rel") = Null

                                 RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                                 RstTableDestination("Déb_Code_Enc") = Nz(DLookup("Ens_Code_Enc", "ENCAISSEUR", "Ens_CodePostal='" & Left(RstTableSource("AD_CodePostal"), 2) & "'"), "XX")

                           End Select

                        Case Else

                           RstTableDestination("Déb_Situation") = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action= 3 ")

                           RstTableDestination("Déb_niv_relance") = 7

                           RstTableDestination("Déb_Date_proch_rel") = Null

                           RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                           RstTableDestination("Déb_Code_Enc") = RstTableSource("Déb_Code_Enc")

                     End Select

                  Case 2

                  Case 3

                     RstTableDestination("Déb_niv_relance") = 7

                     RstTableDestination("Déb_Date_proch_rel") = Null

                     RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                     Select Case RstTableSource("Déb_Date_der_rel") >= "31/08/2005"
                        Case True

                           Select Case IsNull(RstTableSource("Déb_Code_Enc"))
                              Case True

                                 RstTableDestination("Déb_Code_Enc") = Nz(DLookup("Ens_Code_Enc", "ENCAISSEUR", "Ens_CodePostal='" & Left(RstTableSource("AD_CodePostal"), 2) & "'"), "XX")

                              Case False

                           End Select

                        Case False

                     End Select

                  Case 4, 7, 8, 9, 10

                     RstTableDestination("Déb_niv_relance") = IIf(RstTableSource("Déb_niv_relance") = "*", 7, RstTableSource("Déb_niv_relance"))

                     RstTableDestination("Déb_Date_proch_rel") = RstTableSource("Déb_Date_proch_rel")

                     RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                  Case 5

                     RstTableDestination("Déb_niv_relance") = 1

                     RstTableDestination("Déb_Date_proch_rel") = Null

                     RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                  Case 6

                     RstTableDestination("Déb_niv_relance") = 1

                     RstTableDestination("Déb_Date_proch_rel") = Null

                     RstTableDestination("Déb_Date_der_rel") = Nz(RstTableSource("Déb_Date_Edit_Irr"), IIf(RstTableSource("Déb_CI_Edit"), "01/01/1900", Null))

               End Select

            Case "Déb_Date_Chèque", "Déb_Lot_Imp_Id", "Déb_Date_Exp_Irr", "Déb_Date_der_enc", "Déb_Prin_enc_m1", "Déb_Prin_enc_m2", "Déb_Frais_enc_m1", "Déb_Frais_enc_m2", "Déb_Mont_reversé_m1", "Déb_Mont_reversé_m2", "Déb_Date_der_règ", "Déb_Direct_Réglé", "Déb_CI_Select", "Déb_Clé_Débiteur", "Déb_Clé_Adr_Débiteur", "Déb_CI_Edit", "Déb_Archiver", "Déb_Date_Edit_Irr", "Déb_Date_créat_Ater", "Déb_niv_relance", "Déb_Date_proch_rel", "Déb_Date_der_rel", "Déb_Code_Enc", "AD_CodePostal"

            Case Else

               RstTableDestination(FldChamp.Name) = FldChamp.Value

         End Select

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " DEBITEUR:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrFichierSource = "ARCHDEB2003OLD.MDB"

   Set DBBaseSource = DBEngine.OpenDatabase(StrPath & StrFichierSource)

   StrSQL = "SELECT ARCHIVE_DEBITEUR.*, (SELECT First(AD_CodePostal) FROM ARCHIVE_ADRESSE_DEBITEUR WHERE AD_Code_Débiteur=Déb_Code;) AS AD_CodePostal FROM ARCHIVE_DEBITEUR;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("DEBITEUR")

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         Select Case FldChamp.Name
            Case "Déb_Num_Chèque"

               Select Case IsDate(FldChamp.Value)
                  Case True

                     RstTableDestination("Déb_Date_Chèque") = RstTableSource("Déb_Num_Chèque")

                  Case False

                     RstTableDestination("Déb_Num_Chèque") = RstTableSource("Déb_Num_Chèque")

               End Select

               Select Case IsDate(RstTableSource("Déb_Date_Chèque"))
                  Case True

                     RstTableDestination("Déb_Date_Chèque") = RstTableSource("Déb_Date_Chèque")

                  Case False

               End Select

            Case "Déb_Situation"

               RstTableDestination("Déb_Situation") = RstTableSource("Déb_Situation")

               RstTableDestination("Déb_Code_Enc") = RstTableSource("Déb_Code_Enc")

               Select Case DLookup("CG_Code_Action", "CODE_GESTION", "CG_Code='" & RstTableSource("Déb_Situation") & "'")
                  Case 1

                     Select Case RstTableSource("Déb_niv_relance")
                        Case 1, 2, 3, 4, 5, 6

                           Select Case RstTableSource("Déb_Date_proch_rel") >= "01/01/2005" Or IsNull(RstTableSource("Déb_Date_proch_rel"))
                              Case True

                                 RstTableDestination("Déb_niv_relance") = RstTableSource("Déb_niv_relance")

                                 RstTableDestination("Déb_Date_proch_rel") = RstTableSource("Déb_Date_proch_rel")

                                 RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                              Case False

                                 RstTableDestination("Déb_Situation") = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action= 3 ")

                                 RstTableDestination("Déb_niv_relance") = 7

                                 RstTableDestination("Déb_Date_proch_rel") = Null

                                 RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                                 RstTableDestination("Déb_Code_Enc") = Nz(DLookup("Ens_Code_Enc", "ENCAISSEUR", "Ens_CodePostal='" & Left(RstTableSource("AD_CodePostal"), 2) & "'"), "XX")

                           End Select

                        Case Else

                           RstTableDestination("Déb_Situation") = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action= 3 ")

                           RstTableDestination("Déb_niv_relance") = 7

                           RstTableDestination("Déb_Date_proch_rel") = Null

                           RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                           RstTableDestination("Déb_Code_Enc") = RstTableSource("Déb_Code_Enc")

                     End Select

                  Case 2

                  Case 3

                     RstTableDestination("Déb_niv_relance") = 7

                     RstTableDestination("Déb_Date_proch_rel") = Null

                     RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                     Select Case RstTableSource("Déb_Date_der_rel") >= "31/08/2005"
                        Case True

                           Select Case IsNull(RstTableSource("Déb_Code_Enc"))
                              Case True

                                 RstTableDestination("Déb_Code_Enc") = Nz(DLookup("Ens_Code_Enc", "ENCAISSEUR", "Ens_CodePostal='" & Left(RstTableSource("AD_CodePostal"), 2) & "'"), "XX")

                              Case False

                           End Select

                        Case False

                     End Select

                  Case 4, 7, 8, 9, 10

                     RstTableDestination("Déb_niv_relance") = IIf(RstTableSource("Déb_niv_relance") = "*", 7, RstTableSource("Déb_niv_relance"))

                     RstTableDestination("Déb_Date_proch_rel") = RstTableSource("Déb_Date_proch_rel")

                     RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                  Case 5

                     RstTableDestination("Déb_niv_relance") = 1

                     RstTableDestination("Déb_Date_proch_rel") = Null

                     RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                  Case 6

                     RstTableDestination("Déb_niv_relance") = 1

                     RstTableDestination("Déb_Date_proch_rel") = Null

                     RstTableDestination("Déb_Date_der_rel") = Nz(RstTableSource("Déb_Date_Edit_Irr"), IIf(RstTableSource("Déb_CI_Edit"), "01/01/1900", Null))

               End Select

            Case "Déb_Date_Chèque", "Déb_Lot_Imp_Id", "Déb_Date_Exp_Irr", "Déb_Date_der_enc", "Déb_Prin_enc_m1", "Déb_Prin_enc_m2", "Déb_Frais_enc_m1", "Déb_Frais_enc_m2", "Déb_Mont_reversé_m1", "Déb_Mont_reversé_m2", "Déb_Date_der_règ", "Déb_Direct_Réglé", "Déb_CI_Select", "Déb_Clé_Débiteur", "Déb_Clé_Adr_Débiteur", "Déb_CI_Edit", "Déb_Archiver", "Déb_Date_Edit_Irr", "Déb_Date_créat_Ater", "Déb_niv_relance", "Déb_Date_proch_rel", "Déb_Date_der_rel", "Déb_Code_Enc", "AD_CodePostal"

            Case Else

               RstTableDestination(FldChamp.Name) = FldChamp.Value

         End Select

      Next

      RstTableDestination.Update

      DBBaseDestination.Execute "INSERT INTO TMP_DEBITEUR (TDéb_Code) VALUES('" & RstTableSource("Déb_Code") & "');", dbFailOnError

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " ARCHIVE_DEBITEUROLD:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrFichierSource = "ARCHDEB2003BACK.MDB"

   Set DBBaseSource = DBEngine.OpenDatabase(StrPath & StrFichierSource)

   StrSQL = "SELECT ARCHIVE_DEBITEUR.*, (SELECT First(AD_CodePostal) FROM ARCHIVE_ADRESSE_DEBITEUR WHERE AD_Code_Débiteur=Déb_Code;) AS AD_CodePostal FROM ARCHIVE_DEBITEUR;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("DEBITEUR")

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         Select Case FldChamp.Name
            Case "Déb_Num_Chèque"

               Select Case IsDate(FldChamp.Value)
                  Case True

                     RstTableDestination("Déb_Date_Chèque") = RstTableSource("Déb_Num_Chèque")

                  Case False

                     RstTableDestination("Déb_Num_Chèque") = RstTableSource("Déb_Num_Chèque")

               End Select

               Select Case IsDate(RstTableSource("Déb_Date_Chèque"))
                  Case True

                     RstTableDestination("Déb_Date_Chèque") = RstTableSource("Déb_Date_Chèque")

                  Case False

               End Select

            Case "Déb_Situation"

               RstTableDestination("Déb_Situation") = RstTableSource("Déb_Situation")

               RstTableDestination("Déb_Code_Enc") = RstTableSource("Déb_Code_Enc")

               Select Case DLookup("CG_Code_Action", "CODE_GESTION", "CG_Code='" & RstTableSource("Déb_Situation") & "'")
                  Case 1

                     Select Case RstTableSource("Déb_niv_relance")
                        Case 1, 2, 3, 4, 5, 6

                           Select Case RstTableSource("Déb_Date_proch_rel") >= "01/01/2005" Or IsNull(RstTableSource("Déb_Date_proch_rel"))
                              Case True

                                 RstTableDestination("Déb_niv_relance") = RstTableSource("Déb_niv_relance")

                                 RstTableDestination("Déb_Date_proch_rel") = RstTableSource("Déb_Date_proch_rel")

                                 RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                              Case False

                                 RstTableDestination("Déb_Situation") = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action= 3 ")

                                 RstTableDestination("Déb_niv_relance") = 7

                                 RstTableDestination("Déb_Date_proch_rel") = Null

                                 RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                                 RstTableDestination("Déb_Code_Enc") = Nz(DLookup("Ens_Code_Enc", "ENCAISSEUR", "Ens_CodePostal='" & Left(RstTableSource("AD_CodePostal"), 2) & "'"), "XX")

                           End Select

                        Case Else

                           RstTableDestination("Déb_Situation") = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action= 3 ")

                           RstTableDestination("Déb_niv_relance") = 7

                           RstTableDestination("Déb_Date_proch_rel") = Null

                           RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                           RstTableDestination("Déb_Code_Enc") = RstTableSource("Déb_Code_Enc")

                     End Select

                  Case 2

                  Case 3

                     RstTableDestination("Déb_niv_relance") = 7

                     RstTableDestination("Déb_Date_proch_rel") = Null

                     RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                     Select Case RstTableSource("Déb_Date_der_rel") >= "31/08/2005"
                        Case True

                           Select Case IsNull(RstTableSource("Déb_Code_Enc"))
                              Case True

                                 RstTableDestination("Déb_Code_Enc") = Nz(DLookup("Ens_Code_Enc", "ENCAISSEUR", "Ens_CodePostal='" & Left(RstTableSource("AD_CodePostal"), 2) & "'"), "XX")

                              Case False

                           End Select

                        Case False

                     End Select

                  Case 4, 7, 8, 9, 10

                     RstTableDestination("Déb_niv_relance") = IIf(RstTableSource("Déb_niv_relance") = "*", 7, RstTableSource("Déb_niv_relance"))

                     RstTableDestination("Déb_Date_proch_rel") = RstTableSource("Déb_Date_proch_rel")

                     RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                  Case 5

                     RstTableDestination("Déb_niv_relance") = 1

                     RstTableDestination("Déb_Date_proch_rel") = Null

                     RstTableDestination("Déb_Date_der_rel") = RstTableSource("Déb_Date_der_rel")

                  Case 6

                     RstTableDestination("Déb_niv_relance") = 1

                     RstTableDestination("Déb_Date_proch_rel") = Null

                     RstTableDestination("Déb_Date_der_rel") = Nz(RstTableSource("Déb_Date_Edit_Irr"), IIf(RstTableSource("Déb_CI_Edit"), "01/01/1900", Null))

               End Select

            Case "Déb_Date_Chèque", "Déb_Lot_Imp_Id", "Déb_Date_Exp_Irr", "Déb_Date_der_enc", "Déb_Prin_enc_m1", "Déb_Prin_enc_m2", "Déb_Frais_enc_m1", "Déb_Frais_enc_m2", "Déb_Mont_reversé_m1", "Déb_Mont_reversé_m2", "Déb_Date_der_règ", "Déb_Direct_Réglé", "Déb_CI_Select", "Déb_Clé_Débiteur", "Déb_Clé_Adr_Débiteur", "Déb_CI_Edit", "Déb_Archiver", "Déb_Date_Edit_Irr", "Déb_Date_créat_Ater", "Déb_niv_relance", "Déb_Date_proch_rel", "Déb_Date_der_rel", "Déb_Code_Enc", "AD_CodePostal"

            Case Else

               RstTableDestination(FldChamp.Name) = FldChamp.Value

         End Select

      Next

      RstTableDestination.Update

      DBBaseDestination.Execute "INSERT INTO TMP_DEBITEUR (TDéb_Code) VALUES('" & RstTableSource("Déb_Code") & "');", dbFailOnError

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " ARCHIVE_DEBITEURBACK:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrFichierSource = "TGREF2003OLD.MDB"

   Set DBBaseSource = DBEngine.OpenDatabase(StrPath & StrFichierSource)

   StrSQL = "SELECT ADRESSE_DEBITEUR.AD_Code_Débiteur, ADRESSE_DEBITEUR.AD_TypeAdresse_ID, Last(ADRESSE_DEBITEUR.AD_GenreComplémentaire) AS AD_GenreComplémentaire, Last(ADRESSE_DEBITEUR.AD_NomComplémentaire) AS AD_NomComplémentaire, Last(ADRESSE_DEBITEUR.AD_Adr1) AS AD_Adr1, Last(ADRESSE_DEBITEUR.AD_Adr2) AS AD_Adr2, Last(ADRESSE_DEBITEUR.AD_Adr3) AS AD_Adr3, Last(ADRESSE_DEBITEUR.AD_Adr4) AS AD_Adr4, Last(ADRESSE_DEBITEUR.AD_CodePostal) AS AD_CodePostal, Last(ADRESSE_DEBITEUR.AD_Ville) AS AD_Ville, Last(ADRESSE_DEBITEUR.AD_Pays) AS AD_Pays, Last(ADRESSE_DEBITEUR.AD_Tel) AS AD_Tel, Last(ADRESSE_DEBITEUR.AD_Tel2) AS AD_Tel2, Last(ADRESSE_DEBITEUR.AD_Fax) AS AD_Fax, Last(ADRESSE_DEBITEUR.AD_email) AS AD_email, Last(ADRESSE_DEBITEUR.AD_Note) AS AD_Note " & _
               "FROM ADRESSE_DEBITEUR WHERE (ADRESSE_DEBITEUR.AD_TypeAdresse_ID Is Not Null) GROUP BY ADRESSE_DEBITEUR.AD_Code_Débiteur, ADRESSE_DEBITEUR.AD_TypeAdresse_ID;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("ADRESSE_DEBITEUR")

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " ADRESSE_DEBITEUR:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrFichierSource = "ARCHDEB2003OLD.MDB"

   Set DBBaseSource = DBEngine.OpenDatabase(StrPath & StrFichierSource)

   StrSQL = "SELECT ARCHIVE_ADRESSE_DEBITEUR.AD_Code_Débiteur, ARCHIVE_ADRESSE_DEBITEUR.AD_TypeAdresse_ID, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_GenreComplémentaire) AS AD_GenreComplémentaire, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_NomComplémentaire) AS AD_NomComplémentaire, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Adr1) AS AD_Adr1, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Adr2) AS AD_Adr2, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Adr3) AS AD_Adr3, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Adr4) AS AD_Adr4, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_CodePostal) AS AD_CodePostal, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Ville) AS AD_Ville, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Pays) AS AD_Pays, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Tel) AS AD_Tel, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Tel2) AS AD_Tel2, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Fax) AS AD_Fax, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_email) AS AD_email, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Note) AS AD_Note " & _
                 "FROM ARCHIVE_ADRESSE_DEBITEUR WHERE (ARCHIVE_ADRESSE_DEBITEUR.AD_TypeAdresse_ID Is Not Null) GROUP BY ARCHIVE_ADRESSE_DEBITEUR.AD_Code_Débiteur, ARCHIVE_ADRESSE_DEBITEUR.AD_TypeAdresse_ID;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("ADRESSE_DEBITEUR")

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " ADRESSE_DEBITEUROLD:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrFichierSource = "ARCHDEB2003BACK.MDB"

   Set DBBaseSource = DBEngine.OpenDatabase(StrPath & StrFichierSource)

   StrSQL = "SELECT ARCHIVE_ADRESSE_DEBITEUR.AD_Code_Débiteur, ARCHIVE_ADRESSE_DEBITEUR.AD_TypeAdresse_ID, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_GenreComplémentaire) AS AD_GenreComplémentaire, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_NomComplémentaire) AS AD_NomComplémentaire, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Adr1) AS AD_Adr1, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Adr2) AS AD_Adr2, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Adr3) AS AD_Adr3, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Adr4) AS AD_Adr4, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_CodePostal) AS AD_CodePostal, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Ville) AS AD_Ville, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Pays) AS AD_Pays, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Tel) AS AD_Tel, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Tel2) AS AD_Tel2, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Fax) AS AD_Fax, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_email) AS AD_email, Last(ARCHIVE_ADRESSE_DEBITEUR.AD_Note) AS AD_Note " & _
                 "FROM ARCHIVE_ADRESSE_DEBITEUR WHERE (ARCHIVE_ADRESSE_DEBITEUR.AD_TypeAdresse_ID Is Not Null) GROUP BY ARCHIVE_ADRESSE_DEBITEUR.AD_Code_Débiteur, ARCHIVE_ADRESSE_DEBITEUR.AD_TypeAdresse_ID;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("ADRESSE_DEBITEUR")

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " ADRESSE_DEBITEURBACK:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrFichierSource = "TGREF2003OLD.MDB"

   Set DBBaseSource = DBEngine.OpenDatabase(StrPath & StrFichierSource)

   StrSQL = "SELECT ENCAISSEMENT.Enc_ID, ENCAISSEMENT.Enc_Code_Débiteur,  IIf(IsNull([ENCAISSEMENT].[Enc_Date]),Enc_ID,[ENCAISSEMENT].[Enc_Date]) AS Enc_Date, ENCAISSEMENT.Enc_Type, ENCAISSEMENT.Enc_Montant_m1, ENCAISSEMENT.Enc_Montant_m2, ENCAISSEMENT.Enc_Montant_mo, ENCAISSEMENT.Enc_Prin_m1, ENCAISSEMENT.Enc_Prin_m2, ENCAISSEMENT.Enc_Frais_m1, ENCAISSEMENT.Enc_Frais_m2, ENCAISSEMENT.Enc_Type_Paiement, ENCAISSEMENT.Enc_Observation, Switch([Enc_Facturé]=True,'F',[ENCAISSEMENT].[Enc_Situation]='P','R',[ENCAISSEMENT].[Enc_Situation]<>'P',Null) AS Enc_Situation " & _
                 "FROM ENCAISSEMENT;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("ENCAISSEMENT")

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      Select Case RstTableDestination("Enc_Montant_mo")
         Case cFRANC

            Select Case (Round(RstTableDestination("Enc_Prin_m1"), 2) <> RstTableDestination("Enc_Prin_m1")) Or (Round(RstTableDestination("Enc_Frais_m1"), 2) <> RstTableDestination("Enc_Frais_m1")) Or (Round(RstTableDestination("Enc_Montant_m1"), 2) <> RstTableDestination("Enc_Montant_m1"))
               Case True

                  RstTableDestination("Enc_Montant_m1") = Round(RstTableDestination("Enc_Montant_m1"), 2)

                  RstTableDestination("Enc_Prin_m1") = Round(RstTableDestination("Enc_Prin_m1"), 2)

                  RstTableDestination("Enc_Frais_m1") = RstTableDestination("Enc_Montant_m1") - RstTableDestination("Enc_Prin_m1")

                  RstTableDestination("Enc_Montant_m2") = RstTableDestination("Enc_Montant_m1") * DonnerTauxConvertion(cFRANC)

                  RstTableDestination("Enc_Prin_m2") = RstTableDestination("Enc_Prin_m1") * DonnerTauxConvertion(cFRANC)

                  RstTableDestination("Enc_Frais_m2") = RstTableDestination("Enc_Frais_m1") * DonnerTauxConvertion(cFRANC)

               Case False

            End Select

         Case cEURO

            Select Case (Round(RstTableDestination("Enc_Prin_m2"), 2) <> RstTableDestination("Enc_Prin_m2")) Or (Round(RstTableDestination("Enc_Frais_m2"), 2) <> RstTableDestination("Enc_Frais_m2")) Or (Round(RstTableDestination("Enc_Montant_m2"), 2) <> RstTableDestination("Enc_Montant_m2"))
               Case True

                  RstTableDestination("Enc_Montant_m2") = Round(RstTableDestination("Enc_Montant_m2"), 2)

                  RstTableDestination("Enc_Prin_m2") = Round(RstTableDestination("Enc_Prin_m2"), 2)

                  RstTableDestination("Enc_Frais_m2") = RstTableDestination("Enc_Montant_m2") - RstTableDestination("Enc_Prin_m2")

                  RstTableDestination("Enc_Montant_m1") = RstTableDestination("Enc_Montant_m2") * DonnerTauxConvertion(cEURO)

                  RstTableDestination("Enc_Prin_m1") = RstTableDestination("Enc_Prin_m2") * DonnerTauxConvertion(cEURO)

                  RstTableDestination("Enc_Frais_m1") = RstTableDestination("Enc_Frais_m2") * DonnerTauxConvertion(cEURO)

               Case False

            End Select

      End Select

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " ENCAISSEMENT:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrFichierSource = "ARCHDEB2003OLD.MDB"

   Set DBBaseSource = DBEngine.OpenDatabase(StrPath & StrFichierSource)

   StrSQL = "SELECT ARCHIVE_ENCAISSEMENT.Enc_ID, ARCHIVE_ENCAISSEMENT.Enc_Code_Débiteur,  IIf(IsNull([ARCHIVE_ENCAISSEMENT].[Enc_Date]),Enc_ID,[ARCHIVE_ENCAISSEMENT].[Enc_Date]) AS Enc_Date, ARCHIVE_ENCAISSEMENT.Enc_Type, ARCHIVE_ENCAISSEMENT.Enc_Montant_m1, ARCHIVE_ENCAISSEMENT.Enc_Montant_m2, ARCHIVE_ENCAISSEMENT.Enc_Montant_mo, ARCHIVE_ENCAISSEMENT.Enc_Prin_m1, ARCHIVE_ENCAISSEMENT.Enc_Prin_m2, ARCHIVE_ENCAISSEMENT.Enc_Frais_m1, ARCHIVE_ENCAISSEMENT.Enc_Frais_m2, ARCHIVE_ENCAISSEMENT.Enc_Type_Paiement, ARCHIVE_ENCAISSEMENT.Enc_Observation, Switch([Enc_Facturé]=True,'F',[ARCHIVE_ENCAISSEMENT].[Enc_Situation]='P','R',[ARCHIVE_ENCAISSEMENT].[Enc_Situation]<>'P',Null) AS Enc_Situation " & _
                 "FROM ARCHIVE_ENCAISSEMENT;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("ENCAISSEMENT")

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         RstTableDestination(FldChamp.Name) = FldChamp.Value

      Next

      Select Case RstTableDestination("Enc_Montant_mo")
         Case cFRANC

            Select Case (Round(RstTableDestination("Enc_Prin_m1"), 2) <> RstTableDestination("Enc_Prin_m1")) Or (Round(RstTableDestination("Enc_Frais_m1"), 2) <> RstTableDestination("Enc_Frais_m1")) Or (Round(RstTableDestination("Enc_Montant_m1"), 2) <> RstTableDestination("Enc_Montant_m1"))
               Case True

                  RstTableDestination("Enc_Montant_m1") = Round(RstTableDestination("Enc_Montant_m1"), 2)

                  RstTableDestination("Enc_Prin_m1") = Round(RstTableDestination("Enc_Prin_m1"), 2)

                  RstTableDestination("Enc_Frais_m1") = RstTableDestination("Enc_Montant_m1") - RstTableDestination("Enc_Prin_m1")

                  RstTableDestination("Enc_Montant_m2") = RstTableDestination("Enc_Montant_m1") * DonnerTauxConvertion(cFRANC)

                  RstTableDestination("Enc_Prin_m2") = RstTableDestination("Enc_Prin_m1") * DonnerTauxConvertion(cFRANC)

                  RstTableDestination("Enc_Frais_m2") = RstTableDestination("Enc_Frais_m1") * DonnerTauxConvertion(cFRANC)

               Case False

            End Select

         Case cEURO

            Select Case (Round(RstTableDestination("Enc_Prin_m2"), 2) <> RstTableDestination("Enc_Prin_m2")) Or (Round(RstTableDestination("Enc_Frais_m2"), 2) <> RstTableDestination("Enc_Frais_m2")) Or (Round(RstTableDestination("Enc_Montant_m2"), 2) <> RstTableDestination("Enc_Montant_m2"))
               Case True

                  RstTableDestination("Enc_Montant_m2") = Round(RstTableDestination("Enc_Montant_m2"), 2)

                  RstTableDestination("Enc_Prin_m2") = Round(RstTableDestination("Enc_Prin_m2"), 2)

                  RstTableDestination("Enc_Frais_m2") = RstTableDestination("Enc_Montant_m2") - RstTableDestination("Enc_Prin_m2")

                  RstTableDestination("Enc_Montant_m1") = RstTableDestination("Enc_Montant_m2") * DonnerTauxConvertion(cEURO)

                  RstTableDestination("Enc_Prin_m1") = RstTableDestination("Enc_Prin_m2") * DonnerTauxConvertion(cEURO)

                  RstTableDestination("Enc_Frais_m1") = RstTableDestination("Enc_Frais_m2") * DonnerTauxConvertion(cEURO)

               Case False

            End Select

      End Select

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " ENCAISSEMENTOLD:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrFichierSource = "ARCHDEB2003BACK.MDB"

   Set DBBaseSource = DBEngine.OpenDatabase(StrPath & StrFichierSource)

   StrSQL = "SELECT ARCHIVE_ENCAISSEMENT.Enc_ID, ARCHIVE_ENCAISSEMENT.Enc_Code_Débiteur,  IIf(IsNull([ARCHIVE_ENCAISSEMENT].[Enc_Date]),Enc_ID,[ARCHIVE_ENCAISSEMENT].[Enc_Date]) AS Enc_Date, ARCHIVE_ENCAISSEMENT.Enc_Type, ARCHIVE_ENCAISSEMENT.Enc_Montant_m1, ARCHIVE_ENCAISSEMENT.Enc_Montant_m2, ARCHIVE_ENCAISSEMENT.Enc_Montant_mo, ARCHIVE_ENCAISSEMENT.Enc_Prin_m1, ARCHIVE_ENCAISSEMENT.Enc_Prin_m2, ARCHIVE_ENCAISSEMENT.Enc_Frais_m1, ARCHIVE_ENCAISSEMENT.Enc_Frais_m2, ARCHIVE_ENCAISSEMENT.Enc_Type_Paiement, ARCHIVE_ENCAISSEMENT.Enc_Observation, Switch([Enc_Facturé]=True,'F',[ARCHIVE_ENCAISSEMENT].[Enc_Situation]='P','R',[ARCHIVE_ENCAISSEMENT].[Enc_Situation]<>'P',Null) AS Enc_Situation " & _
                 "FROM ARCHIVE_ENCAISSEMENT;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("ENCAISSEMENT")

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         Select Case FldChamp.Name
            Case "Enc_Id"

               Set RstTable = DBBaseDestination.OpenRecordset("SELECT Max(Enc_Id) FROM ENCAISSEMENT;")

               RstTableDestination(FldChamp.Name) = RstTable(0) + 1

               RstTable.Close

            Case Else

               RstTableDestination(FldChamp.Name) = FldChamp.Value

         End Select

      Next

      Select Case RstTableDestination("Enc_Montant_mo")
         Case cFRANC

            Select Case (Round(RstTableDestination("Enc_Prin_m1"), 2) <> RstTableDestination("Enc_Prin_m1")) Or (Round(RstTableDestination("Enc_Frais_m1"), 2) <> RstTableDestination("Enc_Frais_m1")) Or (Round(RstTableDestination("Enc_Montant_m1"), 2) <> RstTableDestination("Enc_Montant_m1"))
               Case True

                  RstTableDestination("Enc_Montant_m1") = Round(RstTableDestination("Enc_Montant_m1"), 2)

                  RstTableDestination("Enc_Prin_m1") = Round(RstTableDestination("Enc_Prin_m1"), 2)

                  RstTableDestination("Enc_Frais_m1") = RstTableDestination("Enc_Montant_m1") - RstTableDestination("Enc_Prin_m1")

                  RstTableDestination("Enc_Montant_m2") = RstTableDestination("Enc_Montant_m1") * DonnerTauxConvertion(cFRANC)

                  RstTableDestination("Enc_Prin_m2") = RstTableDestination("Enc_Prin_m1") * DonnerTauxConvertion(cFRANC)

                  RstTableDestination("Enc_Frais_m2") = RstTableDestination("Enc_Frais_m1") * DonnerTauxConvertion(cFRANC)

               Case False

            End Select

         Case cEURO

            Select Case (Round(RstTableDestination("Enc_Prin_m2"), 2) <> RstTableDestination("Enc_Prin_m2")) Or (Round(RstTableDestination("Enc_Frais_m2"), 2) <> RstTableDestination("Enc_Frais_m2")) Or (Round(RstTableDestination("Enc_Montant_m2"), 2) <> RstTableDestination("Enc_Montant_m2"))
               Case True

                  RstTableDestination("Enc_Montant_m2") = Round(RstTableDestination("Enc_Montant_m2"), 2)

                  RstTableDestination("Enc_Prin_m2") = Round(RstTableDestination("Enc_Prin_m2"), 2)

                  RstTableDestination("Enc_Frais_m2") = RstTableDestination("Enc_Montant_m2") - RstTableDestination("Enc_Prin_m2")

                  RstTableDestination("Enc_Montant_m1") = RstTableDestination("Enc_Montant_m2") * DonnerTauxConvertion(cEURO)

                  RstTableDestination("Enc_Prin_m1") = RstTableDestination("Enc_Prin_m2") * DonnerTauxConvertion(cEURO)

                  RstTableDestination("Enc_Frais_m1") = RstTableDestination("Enc_Frais_m2") * DonnerTauxConvertion(cEURO)

               Case False

            End Select

      End Select

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Now, "Long Time") & " ENCAISSEMENTBACK:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrFichierSource = "TGREF2003OLD.MDB"

   Set DBBaseSource = DBEngine.OpenDatabase(StrPath & StrFichierSource)

   StrSQL = "SELECT IIf(IsNull([ENCAISSEMENT].[Enc_Date_Rgt]) And [Enc_Situation]='P',#1/1/1900#,[ENCAISSEMENT].[Enc_Date_Rgt]) AS Enc_Date_Rgt, CLIENT.Cli_Code, CLIENT.Cli_Date_contrat, nz([CLIENT].[Cli_Pourcentage],0) AS Cli_Pourcentage, DEBITEUR.Déb_Code, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb_Mont_Prin_m2, CLIENT.Cli_MonnaieEnCours, ENCAISSEMENT_TYPE.ETy_Mode, ENCAISSEMENT.Enc_ID, IIf(ENCAISSEMENT.Enc_Montant_mo=1,Round(ENCAISSEMENT.Enc_Prin_m1,2),ENCAISSEMENT.Enc_Prin_m1) AS Enc_Prin_m1, IIf(ENCAISSEMENT.Enc_Montant_mo=2,Round(ENCAISSEMENT.Enc_Prin_m2,2),ENCAISSEMENT.Enc_Prin_m2) AS Enc_Prin_m2 " & _
                 "FROM (CLIENT INNER JOIN DEBITEUR ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client) INNER JOIN (ENCAISSEMENT INNER JOIN ENCAISSEMENT_TYPE ON ENCAISSEMENT.Enc_Type = ENCAISSEMENT_TYPE.ETy_Code) ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur " & _
                 "WHERE (((ENCAISSEMENT.Enc_Situation)='P')) " & _
                 "GROUP BY IIf(IsNull([ENCAISSEMENT].[Enc_Date_Rgt]) And [Enc_Situation]='P',#1/1/1900#,[ENCAISSEMENT].[Enc_Date_Rgt]), CLIENT.Cli_Code, CLIENT.Cli_Date_contrat, nz([CLIENT].[Cli_Pourcentage],0), DEBITEUR.Déb_Code, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb_Mont_Prin_m2, CLIENT.Cli_MonnaieEnCours, ENCAISSEMENT_TYPE.ETy_Mode, ENCAISSEMENT.Enc_ID, IIf(ENCAISSEMENT.Enc_Montant_mo=1,Round(ENCAISSEMENT.Enc_Prin_m1,2),ENCAISSEMENT.Enc_Prin_m1), IIf(ENCAISSEMENT.Enc_Montant_mo=2,Round(ENCAISSEMENT.Enc_Prin_m2,2),ENCAISSEMENT.Enc_Prin_m2) " & _
                 "HAVING (((IIf(IsNull([ENCAISSEMENT].[Enc_Date_Rgt]) And [Enc_Situation]='P',#1/1/1900#,[ENCAISSEMENT].[Enc_Date_Rgt])) Is Not Null)) " & _
                 "ORDER BY IIf(IsNull([ENCAISSEMENT].[Enc_Date_Rgt]) And [Enc_Situation]='P',#1/1/1900#,[ENCAISSEMENT].[Enc_Date_Rgt]), CLIENT.Cli_Code, DEBITEUR.Déb_Code, ENCAISSEMENT_TYPE.ETy_Mode DESC , ENCAISSEMENT.Enc_ID;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("REGLEMENT")

   Set RstTableDestination = TdfTableDestination.OpenRecordset(dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("REGLEMENT_DETAIL")

   Set RstTableDestination1 = TdfTableDestination.OpenRecordset(dbOpenDynaset)

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      Select Case DatReg_Date <> RstTableSource("Enc_Date_Rgt") Or StrReg_Cli_Code <> RstTableSource("Cli_Code") Or StrRegD_Déb_Code <> RstTableSource("Déb_Code") Or IntRegD_ETy_Mode <> RstTableSource("ETy_Mode")
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

                  RstTableDestination.FindFirst "Reg_Id=" & LngReg_Id

                  RstTableDestination.Edit

                  RstTableDestination("Reg_Montant_Encaisser_m1") = CurReg_Montant_Encaisser_m1

                  RstTableDestination("Reg_Montant_Encaisser_m2") = CurReg_Montant_Encaisser_m2

                  RstTableDestination("Reg_Montant_Direct_m1") = CurReg_Montant_Direct_m1

                  RstTableDestination("Reg_Montant_Direct_m2") = CurReg_Montant_Direct_m2

                  RstTableDestination.Update

            End Select

         Case False

      End Select

      Select Case DatReg_Date <> RstTableSource("Enc_Date_Rgt") Or StrReg_Cli_Code <> RstTableSource("Cli_Code")
         Case True

            DatReg_Date = RstTableSource("Enc_Date_Rgt")

            StrReg_Cli_Code = RstTableSource("Cli_Code")

            CurReg_Montant_Encaisser_m1 = 0

            CurReg_Montant_Encaisser_m2 = 0

            CurReg_Montant_Direct_m1 = 0

            CurReg_Montant_Direct_m2 = 0

            Select Case RstTableSource("Enc_Date_Rgt")
               Case Is >= RstTableSource("Cli_Date_contrat")

                  SngReg_Cli_Pourcentage = RstTableSource("Cli_Pourcentage")

                  IntReg_Montant_mo = RstTableSource("Cli_MonnaieEnCours")

               Case Else

                  Set RstTable = DBBaseDestination.OpenRecordset("SELECT CC_Pourcentage FROM CLIENT_CONTRAT WHERE CC_Code_Client='" & RstTableSource("Cli_Code") & "' AND CC_Date_contrat<=" & ConvDate(RstTableSource("Enc_Date_Rgt")) & " ORDER BY CC_Date_contrat DESC")

                  Select Case RstTable.EOF
                     Case True

                        SngReg_Cli_Pourcentage = RstTableSource("Cli_Pourcentage")

                     Case False

                        SngReg_Cli_Pourcentage = IIf(IsNull(RstTable("CC_Pourcentage")), RstTableSource("Cli_Pourcentage"), RstTable("CC_Pourcentage"))

                  End Select

                  RstTable.Close

                  Set RstTable = DBBaseDestination.OpenRecordset("SELECT CC_Mont_Hon_mo FROM CLIENT_CONTRAT WHERE CC_Code_Client='" & RstTableSource("Cli_Code") & "' AND CC_Date_contrat<=" & ConvDate(RstTableSource("Enc_Date_Rgt")) & " ORDER BY CC_Date_contrat DESC")

                  Select Case RstTable.EOF
                     Case True

                        IntReg_Montant_mo = RstTableSource("Cli_MonnaieEnCours")

                     Case False

                        IntReg_Montant_mo = IIf(IsNull(RstTable("CC_Mont_Hon_mo")), RstTableSource("Cli_MonnaieEnCours"), RstTable("CC_Mont_Hon_mo"))

                  End Select

                  RstTable.Close

            End Select

            RstTableDestination.AddNew

            RstTableDestination("Reg_Date") = DatReg_Date

            RstTableDestination("Reg_Cli_Code") = StrReg_Cli_Code

            RstTableDestination("Reg_TVA") = 0.196

            RstTableDestination("Reg_Cli_Pourcentage") = SngReg_Cli_Pourcentage

            RstTableDestination("Reg_Montant_Encaisser_m1") = CurReg_Montant_Encaisser_m1

            RstTableDestination("Reg_Montant_Encaisser_m2") = CurReg_Montant_Encaisser_m2

            RstTableDestination("Reg_Montant_Direct_m1") = CurReg_Montant_Direct_m1

            RstTableDestination("Reg_Montant_Direct_m2") = CurReg_Montant_Direct_m2

            RstTableDestination("Reg_Montant_mo") = IntReg_Montant_mo

            RstTableDestination("Reg_Deja_Editer") = True

            LngReg_Id = RstTableDestination("Reg_Id")

            RstTableDestination.Update

         Case False

      End Select

      Select Case StrRegD_Déb_Code <> RstTableSource("Déb_Code") Or IntRegD_ETy_Mode <> RstTableSource("ETy_Mode")
         Case True

            StrRegD_Déb_Code = RstTableSource("Déb_Code")

            IntRegD_ETy_Mode = RstTableSource("ETy_Mode")

            CurRegD_Montant_m1 = RstTableSource("Enc_Prin_m1")

            CurRegD_Montant_m2 = RstTableSource("Enc_Prin_m2")

            RstTableDestination1.AddNew

            RstTableDestination1("RegD_Reg_Id") = LngReg_Id

            RstTableDestination1("RegD_Déb_Code") = StrRegD_Déb_Code

            RstTableDestination1("RegD_ETy_Mode") = IntRegD_ETy_Mode

            RstTableDestination1("RegD_Type") = "ACOMPTE"

            RstTableDestination1("RegD_Montant_m1") = CurRegD_Montant_m1

            RstTableDestination1("RegD_Montant_m2") = CurRegD_Montant_m2

            RstTableDestination1("RegD_Montant_mo") = IntReg_Montant_mo

            LngRegD_Id = RstTableDestination1("RegD_Id")

            RstTableDestination1.Update

         Case False

            CurRegD_Montant_m1 = CurRegD_Montant_m1 + RstTableSource("Enc_Prin_m1")

            CurRegD_Montant_m2 = CurRegD_Montant_m2 + RstTableSource("Enc_Prin_m2")

            RstTableDestination1.FindFirst "RegD_Id=" & LngRegD_Id

            RstTableDestination1.Edit

            RstTableDestination1("RegD_Montant_m1") = CurRegD_Montant_m1

            RstTableDestination1("RegD_Montant_m2") = CurRegD_Montant_m2

            RstTableDestination1.Update

      End Select

      DBBaseDestination.Execute "UPDATE ENCAISSEMENT SET ENCAISSEMENT.Enc_RegD_Id = " & LngRegD_Id & " WHERE Enc_ID=" & RstTableSource("Enc_Id") & ";", dbFailOnError

      Select Case IntReg_Montant_mo
         Case cFRANC

            Set RstTable = DBBaseDestination.OpenRecordset("SELECT Sum(Enc_Prin_m1) As Déb_Prin_Enc_m1 FROM ENCAISSEMENT WHERE Enc_RegD_Id<=" & LngRegD_Id & " AND Enc_Code_Débiteur='" & StrRegD_Déb_Code & "'")

            Select Case Round(RstTableSource("Déb_Mont_Prin_m1"), 2) - Round(Nz(RstTable("Déb_Prin_Enc_m1"), 0), 2)
               Case Is <= 0

                  RstTableDestination1.FindFirst "RegD_Id=" & LngRegD_Id

                  RstTableDestination1.Edit

                  RstTableDestination1("RegD_Type") = "SOLDE"

                  RstTableDestination1.Update

               Case Else

            End Select

            RstTable.Close

         Case cEURO

            Set RstTable = DBBaseDestination.OpenRecordset("SELECT Sum(Enc_Prin_m2) As Déb_Prin_Enc_m2 FROM ENCAISSEMENT WHERE Enc_RegD_Id<=" & LngRegD_Id & " AND Enc_Code_Débiteur='" & StrRegD_Déb_Code & "'")

            Select Case Round(RstTableSource("Déb_Mont_Prin_m2"), 2) - Round(Nz(RstTable("Déb_Prin_Enc_m2"), 0), 2)
               Case Is <= 0

                  RstTableDestination1.FindFirst "RegD_Id=" & LngRegD_Id

                  RstTableDestination1.Edit

                  RstTableDestination1("RegD_Type") = "SOLDE"

                  RstTableDestination1.Update

               Case Else

            End Select

            RstTable.Close

      End Select

      RstTableSource.MoveNext

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

         RstTableDestination.FindFirst "Reg_Id=" & LngReg_Id

         RstTableDestination.Edit

         RstTableDestination("Reg_Montant_Encaisser_m1") = CurReg_Montant_Encaisser_m1

         RstTableDestination("Reg_Montant_Encaisser_m2") = CurReg_Montant_Encaisser_m2

         RstTableDestination("Reg_Montant_Direct_m1") = CurReg_Montant_Direct_m1

         RstTableDestination("Reg_Montant_Direct_m2") = CurReg_Montant_Direct_m2

         RstTableDestination.Update

   End Select

   Print #IntFichierNumero, Format(Now, "Long Time") & " REGLEMENT:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   Print #IntFichierNumero, Format(Now, "Long Time") & " REGLEMENT_DETAIL:" & RstTableSource.RecordCount & "=> " & RstTableDestination1.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   RstTableDestination1.Close

   StrSQL = "SELECT FACTURE.Fac_Id, FACTURE.Fac_Date, FACTURE.Fac_CodeCli, FACTURE.Fac_Honoraire, FACTURE.Fac_Montant_mo, ENCAISSEMENT.Enc_Code_Débiteur, ENCAISSEMENT.Enc_Prin_m1, ENCAISSEMENT.Enc_Prin_m2, ENCAISSEMENT_TYPE.ETy_Mode, ENCAISSEMENT.Enc_ID " & _
                 "FROM FACTURE INNER JOIN (ENCAISSEMENT INNER JOIN ENCAISSEMENT_TYPE ON ENCAISSEMENT.Enc_Type = ENCAISSEMENT_TYPE.ETy_Code) ON FACTURE.Fac_Id = ENCAISSEMENT.Enc_FacId " & _
                 "GROUP BY FACTURE.Fac_Id, FACTURE.Fac_Date, FACTURE.Fac_CodeCli, FACTURE.Fac_Honoraire, FACTURE.Fac_Montant_mo, ENCAISSEMENT.Enc_Code_Débiteur, ENCAISSEMENT.Enc_Prin_m1, ENCAISSEMENT.Enc_Prin_m2, ENCAISSEMENT_TYPE.ETy_Mode, ENCAISSEMENT.Enc_ID " & _
                 "ORDER BY FACTURE.Fac_Date, FACTURE.Fac_CodeCli,  ENCAISSEMENT.Enc_Code_Débiteur, ENCAISSEMENT_TYPE.ETy_Mode DESC;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("FACTURE")

   Set RstTableDestination = TdfTableDestination.OpenRecordset(dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("FACTURE_DETAIL")

   Set RstTableDestination1 = TdfTableDestination.OpenRecordset(dbOpenDynaset)

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      Select Case LngEnc_Fac_Id <> RstTableSource("Fac_Id") Or StrFacD_Déb_Code <> RstTableSource("Enc_Code_Débiteur") Or IntFacD_ETy_Mode <> RstTableSource("ETy_Mode")
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

                  RstTableDestination.FindFirst "Fac_Id=" & LngFac_Id

                  RstTableDestination.Edit

                  RstTableDestination("Fac_Montant_Encaisser_m1") = CurFac_Montant_Encaisser_m1

                  RstTableDestination("Fac_Montant_Encaisser_m2") = CurFac_Montant_Encaisser_m2

                  RstTableDestination("Fac_Montant_Direct_m1") = CurFac_Montant_Direct_m1

                  RstTableDestination("Fac_Montant_Direct_m2") = CurFac_Montant_Direct_m2

                  RstTableDestination.Update

            End Select

         Case False

      End Select

      Select Case LngEnc_Fac_Id <> RstTableSource("Fac_Id")
         Case True

            LngEnc_Fac_Id = RstTableSource("Fac_Id")

            Select Case Format(Right(Year(RstTableSource("Fac_Date")), 2), "00") & Format(Month(RstTableSource("Fac_Date")), "00") <> StrFacMois
               Case True

                   StrFacMois = Format(Right(Year(RstTableSource("Fac_Date")), 2), "00") & Format(Month(RstTableSource("Fac_Date")), "00")

                   IntFacNum = 1

               Case False

                  IntFacNum = IntFacNum + 1

            End Select

            CurFac_Montant_Direct_m1 = 0

            CurFac_Montant_Direct_m2 = 0

            CurFac_Montant_Encaisser_m1 = 0

            CurFac_Montant_Encaisser_m2 = 0

            IntFac_Montant_mo = RstTableSource("Fac_Montant_mo")

            RstTableDestination.AddNew

            RstTableDestination("Fac_Num") = StrFacMois & Format(IntFacNum, "0000")

            RstTableDestination("Fac_Date") = Format(RstTableSource("Fac_Date"), "Short Date")

            RstTableDestination("Fac_CodeCli") = RstTableSource("Fac_CodeCli")

            RstTableDestination("Fac_Honoraire") = RstTableSource("Fac_Honoraire")

            RstTableDestination("Fac_TVA") = 0.196

            RstTableDestination("Fac_Montant_Encaisser_m1") = CurFac_Montant_Encaisser_m1

            RstTableDestination("Fac_Montant_Encaisser_m2") = CurFac_Montant_Encaisser_m2

            RstTableDestination("Fac_Montant_Direct_m1") = CurFac_Montant_Direct_m1

            RstTableDestination("Fac_Montant_Direct_m2") = CurFac_Montant_Direct_m2

            RstTableDestination("Fac_Montant_mo") = IntFac_Montant_mo

            RstTableDestination("Fac_Deja_Editer") = True

            LngFac_Id = RstTableDestination("Fac_Id")

            RstTableDestination.Update

         Case False

      End Select

      Select Case StrFacD_Déb_Code <> RstTableSource("Enc_Code_Débiteur") Or IntFacD_ETy_Mode <> RstTableSource("ETy_Mode")
         Case True

            StrFacD_Déb_Code = RstTableSource("Enc_Code_Débiteur")

            IntFacD_ETy_Mode = RstTableSource("ETy_Mode")

            CurFacD_Montant_m1 = RstTableSource("Enc_Prin_m1")

            CurFacD_Montant_m2 = RstTableSource("Enc_Prin_m2")

            RstTableDestination1.AddNew

            RstTableDestination1("FacD_Fac_Id") = LngFac_Id

            RstTableDestination1("FacD_Déb_Code") = StrFacD_Déb_Code

            RstTableDestination1("FacD_ETy_Mode") = IntFacD_ETy_Mode

            RstTableDestination1("FacD_Montant_m1") = CurFacD_Montant_m1

            RstTableDestination1("FacD_Montant_m2") = CurFacD_Montant_m2

            RstTableDestination1("FacD_Montant_mo") = IntFac_Montant_mo

            LngFacD_Id = RstTableDestination1("FacD_Id")

            RstTableDestination1.Update

         Case False

            CurFacD_Montant_m1 = CurFacD_Montant_m1 + RstTableSource("Enc_Prin_m1")

            CurFacD_Montant_m2 = CurFacD_Montant_m2 + RstTableSource("Enc_Prin_m2")

            RstTableDestination1.FindFirst "FacD_Id=" & LngFacD_Id

            RstTableDestination1.Edit

            RstTableDestination1("FacD_Montant_m1") = CurFacD_Montant_m1

            RstTableDestination1("FacD_Montant_m2") = CurFacD_Montant_m2

            RstTableDestination1.Update

      End Select

      RstTableSource.MoveNext

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

         RstTableDestination.FindFirst "Fac_Id=" & LngFac_Id

         RstTableDestination.Edit

         RstTableDestination("Fac_Montant_Encaisser_m1") = CurFac_Montant_Encaisser_m1

         RstTableDestination("Fac_Montant_Encaisser_m2") = CurFac_Montant_Encaisser_m2

         RstTableDestination("Fac_Montant_Direct_m1") = CurFac_Montant_Direct_m1

         RstTableDestination("Fac_Montant_Direct_m2") = CurFac_Montant_Direct_m2

         RstTableDestination.Update

   End Select

   Print #IntFichierNumero, Format(Now, "Long Time") & " FACTURE:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   Print #IntFichierNumero, Format(Now, "Long Time") & " FACTURE_DETAIL:" & RstTableSource.RecordCount & "=> " & RstTableDestination1.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   RstTableDestination1.Close

   CurrentDb.Execute "INSERT INTO ARCHIVE_DEBITEUR SELECT DEBITEUR.* FROM DEBITEUR INNER JOIN TMP_DEBITEUR ON DEBITEUR.Déb_Code = TMP_DEBITEUR.TDéb_Code;", dbFailOnError

   CurrentDb.Execute "INSERT INTO ARCHIVE_ADRESSE_DEBITEUR SELECT ADRESSE_DEBITEUR.* FROM ADRESSE_DEBITEUR INNER JOIN TMP_DEBITEUR ON ADRESSE_DEBITEUR.AD_Code_Débiteur = TMP_DEBITEUR.TDéb_Code;", dbFailOnError

   CurrentDb.Execute "INSERT INTO ARCHIVE_ENCAISSEMENT SELECT ENCAISSEMENT.* FROM ENCAISSEMENT INNER JOIN TMP_DEBITEUR ON ENCAISSEMENT.Enc_Code_Débiteur = TMP_DEBITEUR.TDéb_Code;", dbFailOnError

   CurrentDb.Execute "DELETE ENCAISSEMENT.* FROM ENCAISSEMENT WHERE Enc_Code_Débiteur In (SELECT TDéb_Code FROM TMP_DEBITEUR;);", dbFailOnError

   CurrentDb.Execute "DELETE ADRESSE_DEBITEUR.* FROM ADRESSE_DEBITEUR WHERE AD_Code_Débiteur In (SELECT TDéb_Code FROM TMP_DEBITEUR;);", dbFailOnError

   CurrentDb.Execute "DELETE DEBITEUR.* FROM DEBITEUR INNER JOIN TMP_DEBITEUR ON DEBITEUR.Déb_Code = TMP_DEBITEUR.TDéb_Code;", dbFailOnError

   DBBaseSource.Close

   DBBaseDestination.Close

Exit_MajBDD46:

   DoCmd.Close acForm, "FrmPatientez", acSavePrompt

   Close #IntFichierNumero

   Set RstTable = Nothing

   Set RstTableSource = Nothing

   Set RstTableDestination = Nothing

   Set RstTableDestination1 = Nothing

   Set TdfTableSource = Nothing

   Set TdfTableDestination = Nothing

   Set DBBaseSource = Nothing

   Set DBBaseDestination = Nothing

   Set FilFichier = Nothing

   Set FsoFichierSystem = Nothing

   Exit Function

Err_MajBDD46:

   MajBDD46 = False

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : MajBDD46"

   Resume Exit_MajBDD46
End Function
