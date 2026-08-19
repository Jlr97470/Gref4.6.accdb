Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DateGrouping =1
    GridX =5
    GridY =5
    Width =9546
    ItemSuffix =42
    Left =360
    Top =2580
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xd461fb5531e9e240
    End
    RecordSource ="SELECT DEBITEUR.Déb_Code, DEBITEUR.Déb_Qualité, DEBITEUR.Déb_Nom, DEBITEUR.Déb_D"
        "ate_der_rel, DEBITEUR.Déb_Code_Enc, ADRESSE_DEBITEUR.AD_Adr1, ADRESSE_DEBITEUR.A"
        "D_Adr2, ADRESSE_DEBITEUR.AD_Adr3, ADRESSE_DEBITEUR.AD_Adr4, ADRESSE_DEBITEUR.AD_"
        "CodePostal, ADRESSE_DEBITEUR.AD_Ville, ADRESSE_DEBITEUR.AD_Pays, ADRESSE_DEBITEU"
        "R.AD_Tel, CLIENT.Cli_Nom, ADRESSE_CLIENT.AC_Adr1, ADRESSE_CLIENT.AC_Adr2, ADRESS"
        "E_CLIENT.AC_Adr3, ADRESSE_CLIENT.AC_Adr4, ADRESSE_CLIENT.AC_CodePostal, ADRESSE_"
        "CLIENT.AC_Ville, ADRESSE_CLIENT.AC_Pays, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb"
        "_Frais_m1, DEBITEUR.Déb_Mont_Prin_m2, DEBITEUR.Déb_Frais_m2, DEBITEUR.Déb_Typ_Cr"
        "é, DEBITEUR.Déb_MonnaieEnCours, CLIENT.Cli_Grille_FTrans, ENCAISSEMENT.Enc_Date,"
        " ENCAISSEMENT.Enc_Montant_m1, ENCAISSEMENT.Enc_Montant_m2, ENCAISSEMENT.Enc_Code"
        "_Débiteur FROM (CODE_GESTION INNER JOIN ((CLIENT INNER JOIN ADRESSE_CLIENT ON CL"
        "IENT.Cli_Code=ADRESSE_CLIENT.AC_Code_Client) INNER JOIN (DEBITEUR LEFT JOIN ENCA"
        "ISSEMENT ON DEBITEUR.Déb_Code=ENCAISSEMENT.Enc_Code_Débiteur) ON CLIENT.Cli_Code"
        "=DEBITEUR.Déb_Code_Client) ON CODE_GESTION.CG_Code=DEBITEUR.Déb_Situation) INNER"
        " JOIN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Code=ADRESSE_DEBITEUR.AD_Code_Débiteur WH"
        "ERE (((ADRESSE_DEBITEUR.AD_TypeAdresse_ID)=DLookUp(\"Adresse_Courrier\",\"PARAME"
        "TRE_GREF\")) AND ((ADRESSE_CLIENT.AC_TypeAdresse_ID)=DLookUp(\"Adresse_Courrier\""
        ",\"PARAMETRE_GREF\")) AND ((Len([Déb_Code_Enc]))=2) AND ((CODE_GESTION.CG_Code_A"
        "ction)=3)); "
    Caption ="Fiches Encaisseurs"
    OnOpen ="[Event Procedure]"
    Begin
        Begin Label
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Rectangle
            Width =850
            Height =850
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Déb_Code"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Enc_Code_Débiteur"
        End
        Begin PageHeader
            Height =480
            Name ="ZoneEntêtePage"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7830
                    Height =480
                    FontSize =18
                    Name ="Déb_Code_Enc"
                    ControlSource ="Déb_Code_Enc"
                    StatusBarText ="Debiteur Code Encaisseur"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =7920
            Name ="EntêteGroupe0"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =3285
                    Top =5906
                    Width =1380
                    Height =285
                    FontSize =10
                    Name ="TxtFraisDeplacement"
                    ControlSource ="=DLookUp(\"Frais_Déplacement_m1\",\"PARAMETRE_GREF\")"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =525
                            Top =5925
                            Width =2355
                            Height =510
                            FontSize =10
                            Name ="EtiFraisDeplacement"
                            Caption ="Frais de déplacement"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =5505
                    Top =2610
                    Width =1680
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =4950
                            Top =2610
                            Width =525
                            Height =285
                            FontSize =10
                            Name ="EtiDéb_Code"
                            Caption ="Ref :"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =3297
                    Top =4710
                    Width =1380
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="TxtDue"
                    ControlSource ="=Round([Déb_Mont_Prin_m1],2)-Round(nz(DSum(\"Enc_Prin_m1\",\"ENCAISSEMENT\",\"En"
                        "c_Code_Débiteur='\" & [Déb_Code] & \"'\"),0),2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =525
                            Top =4710
                            Width =2460
                            Height =285
                            FontSize =10
                            Name ="EtiDue"
                            Caption ="Sommes dues en principal"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =5775
                    Top =4710
                    Width =3735
                    Height =3210
                    FontSize =10
                    TabIndex =3
                    Name ="TxtDéb_Typ_Cré"
                    ControlSource ="Déb_Typ_Cré"
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =6120
                            Top =4275
                            Width =1485
                            Height =285
                            FontSize =10
                            Name ="EtiDéb_Typ_Cré"
                            Caption ="Libellé Créance"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =5505
                    Top =2955
                    Width =1680
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="TxtAD_Tel"
                    ControlSource ="AD_Tel"
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =4950
                            Top =2955
                            Width =525
                            Height =285
                            FontSize =10
                            Name ="EtiAD_Tel"
                            Caption ="Tel :"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =3327
                    Top =4275
                    Width =1350
                    Height =285
                    FontSize =10
                    Name ="EtiMontant"
                    Caption ="Montant"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =4665
                    Top =5951
                    Width =1110
                    TabIndex =5
                    Name ="TxtFraisDeplacementMonnaie"
                    ControlSource ="=[TxtFraisDeplacement]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =4677
                    Top =4755
                    Width =1095
                    TabIndex =6
                    Name ="TxtDueMonnaie"
                    ControlSource ="=[TxtDue]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =525
                    Top =585
                    Width =3750
                    Height =270
                    FontSize =10
                    TabIndex =7
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =525
                    Top =855
                    Width =3750
                    Height =270
                    FontSize =10
                    TabIndex =8
                    Name ="TxtAC_Adr1"
                    ControlSource ="AC_Adr1"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =525
                    Top =1125
                    Width =3750
                    Height =270
                    FontSize =10
                    TabIndex =9
                    Name ="TxtAC_Adr2"
                    ControlSource ="AC_Adr2"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =525
                    Top =1940
                    Width =975
                    Height =270
                    FontSize =10
                    TabIndex =10
                    Name ="TxtAC_CodePostal"
                    ControlSource ="AC_CodePostal"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =1507
                    Top =1935
                    Width =2775
                    Height =270
                    FontSize =10
                    TabIndex =11
                    Name ="TxtAC_Ville"
                    ControlSource ="AC_Ville"
                End
                Begin Label
                    TextFontFamily =34
                    Left =525
                    Top =285
                    Width =1275
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="EtiCreancier"
                    Caption ="CREANCIER"
                End
                Begin Label
                    TextFontFamily =34
                    Left =4935
                    Top =285
                    Width =1125
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="EtiDebiteur"
                    Caption ="DEBITEUR"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =525
                    Top =1395
                    Width =3750
                    Height =270
                    FontSize =10
                    TabIndex =12
                    Name ="TxtAC_Adr3"
                    ControlSource ="AC_Adr3"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =525
                    Top =1665
                    Width =3750
                    Height =270
                    FontSize =10
                    TabIndex =13
                    Name ="TxtAC_Adr4"
                    ControlSource ="AC_Adr4"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =525
                    Top =2205
                    Width =2895
                    Height =270
                    FontSize =10
                    TabIndex =14
                    Name ="TxtAC_Pays"
                    ControlSource ="AC_Pays"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =4935
                    Top =585
                    Width =4575
                    Height =270
                    FontSize =10
                    TabIndex =15
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4935
                    Top =855
                    Width =3750
                    Height =270
                    FontSize =10
                    TabIndex =16
                    Name ="TxtAD_Adr1"
                    ControlSource ="AD_Adr1"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4935
                    Top =1125
                    Width =3750
                    Height =270
                    FontSize =10
                    TabIndex =17
                    Name ="TxtAD_Adr2"
                    ControlSource ="AD_Adr2"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4935
                    Top =1940
                    Width =975
                    Height =270
                    FontSize =10
                    TabIndex =18
                    Name ="TxtAD_CodePostal"
                    ControlSource ="AD_CodePostal"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =5917
                    Top =1935
                    Width =2775
                    Height =270
                    FontSize =10
                    TabIndex =19
                    Name ="TxtAD_Ville"
                    ControlSource ="AD_Ville"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4935
                    Top =1395
                    Width =3750
                    Height =270
                    FontSize =10
                    TabIndex =20
                    Name ="TxtAD_Adr3"
                    ControlSource ="AD_Adr3"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4935
                    Top =1665
                    Width =3750
                    Height =270
                    FontSize =10
                    TabIndex =21
                    Name ="TxtAD_Adr4"
                    ControlSource ="AD_Adr4"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4935
                    Top =2205
                    Width =2895
                    Height =270
                    FontSize =10
                    TabIndex =22
                    Name ="TxtAD_Pays"
                    ControlSource ="AD_Pays"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =3300
                    Top =5310
                    Width =1380
                    Height =285
                    FontSize =10
                    TabIndex =23
                    Name ="TxtFraisTransactionnel"
                    ControlSource ="=DLookUp(\"Frt_Frais_m1\",\"FRAIS_TRANS\",\"[Frt_Grille]=\" & [Cli_Grille_FTrans"
                        "] & \" AND (([Frt_Début_m1]<=CCur('\" & [TxtDue] & \"') And [Frt_Fin_m1]>=CCur('"
                        "\" & [TxtDue] & \"')) Or ([Frt_Fin_m1]=0 And [Frt_Début_m1]<=CCur('\" & [TxtDue]"
                        " & \"')))\")"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =525
                            Top =5160
                            Width =2355
                            Height =510
                            FontSize =10
                            Name ="EtiFraisTransactionnel"
                            Caption ="Frais d'intervention à titre\015\012transactionnel"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =4680
                    Top =5355
                    Width =1095
                    TabIndex =24
                    Name ="TxtFraisTransactionnelMonnaie"
                    ControlSource ="=[TxtFraisTransactionnel]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =3285
                    Top =6555
                    Width =1380
                    Height =285
                    FontSize =10
                    TabIndex =25
                    Name ="TxtTotal"
                    ControlSource ="=[TxtDue]+[TxtFraisTransactionnel]+[TxtFraisDeplacement]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =525
                            Top =6555
                            Width =2355
                            Height =285
                            FontSize =10
                            Name ="EtiTotal"
                            Caption ="TOTAL EXIGIBLE"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =4665
                    Top =6600
                    Width =1110
                    TabIndex =26
                    Name ="TxtTotalMonnaie"
                    ControlSource ="=[TxtTotal]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =285
            BreakLevel =1
            Name ="EntêteGroupe1"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =3255
                    Width =1380
                    Height =285
                    FontSize =10
                    Name ="TxtDéb_Mont_Prin_m"
                    ControlSource ="=IIf(IsNull([Enc_Code_Débiteur]),Null,[Déb_Mont_Prin_m1])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Width =2355
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="EtiDéb_Mont_Prin_m"
                    ControlSource ="=IIf(IsNull([Enc_Code_Débiteur]),Null,\"Principal Initial\")"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =285
            Name ="Détail"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1695
                    Width =1566
                    Height =285
                    ColumnWidth =945
                    FontSize =10
                    Name ="TxtEnc_Date"
                    ControlSource ="Enc_Date"
                    Format ="Short Date"
                    StatusBarText ="Encaissement Date"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3255
                    Width =1356
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="TxtEnc_Montant_m"
                    ControlSource ="Enc_Montant_m1"
                    Format ="Standard"
                    StatusBarText ="Encaissement Montant En Franc"
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Width =1695
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="EtiAcompte"
                    ControlSource ="=IIf(IsNull([Enc_Code_Débiteur]),Null,\"Acompte\")"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="ZonePiedPage"
        End
    End
End
CodeBehindForm
' See "EtaCourrierFicheEncaisseur.cls"
