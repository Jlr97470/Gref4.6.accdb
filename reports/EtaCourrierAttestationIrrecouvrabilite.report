Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DateGrouping =1
    GridX =5
    GridY =5
    Width =9066
    ItemSuffix =56
    Left =2040
    Top =1620
    OnNoData ="[Event Procedure]"
    Filter ="(Déb_MonnaieEnCours=2)"
    RecSrcDt = Begin
        0x301f0fc0f0e8e240
    End
    RecordSource ="SELECT DISTINCTROW CLIENT.Cli_Code, CLIENT.Cli_Nom, CLIENT.Cli_Int_Nom, ADRESSE_"
        "CLIENT.AC_TypeAdresse_ID, ADRESSE_CLIENT.AC_Adr1, ADRESSE_CLIENT.AC_Adr2, ADRESS"
        "E_CLIENT.AC_Adr3, ADRESSE_CLIENT.AC_Adr4, ADRESSE_CLIENT.AC_CodePostal, ADRESSE_"
        "CLIENT.AC_Ville, ADRESSE_CLIENT.AC_Pays, TMP_IRRECOUVRABLE.TIrr_Déb_Code, DEBITE"
        "UR.Déb_Nom, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb_Mont_Prin_m2, DEBITEUR.Déb_M"
        "onnaieEnCours, DEBITEUR.Déb_Réf1_Client, DEBITEUR.Déb_Réf2_Client, CODE_GESTION."
        "CG_Libellé FROM (CODE_GESTION INNER JOIN ((CLIENT INNER JOIN ADRESSE_CLIENT ON C"
        "LIENT.Cli_Code=ADRESSE_CLIENT.AC_Code_Client) INNER JOIN DEBITEUR ON CLIENT.Cli_"
        "Code=DEBITEUR.Déb_Code_Client) ON CODE_GESTION.CG_Code=DEBITEUR.Déb_Situation) I"
        "NNER JOIN TMP_IRRECOUVRABLE ON DEBITEUR.Déb_Code=TMP_IRRECOUVRABLE.TIrr_Déb_Code"
        " WHERE (((ADRESSE_CLIENT.AC_TypeAdresse_ID)=DLookUp(\"Adresse_Courrier\",\"PARAM"
        "ETRE_GREF\")) AND ((TMP_IRRECOUVRABLE.TIrr_Editer)=True)); "
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
        Begin Line
            Width =1701
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
            GroupFooter = NotDefault
            ControlSource ="TIrr_Déb_Code"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Cli_Code"
        End
        Begin BreakLevel
            ControlSource ="Déb_Nom"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =7042
            Name ="EntêteGroupe"
            Begin
                Begin Label
                    TextFontFamily =34
                    Top =6127
                    Width =9045
                    Height =915
                    FontSize =12
                    Name ="EtiCourrierAttestationIrrcouvrabilite"
                    Caption ="               La présente ATTESTATION vous servira de JUSTIFICATIF COMPTABLE et"
                        " éventuellement FISCAL.\015\012"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =630
                    Top =5062
                    Width =1680
                    Height =315
                    FontSize =12
                    Name ="TxtCli_Code"
                    ControlSource ="Cli_Code"
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Top =5062
                            Width =585
                            Height =330
                            FontSize =12
                            Name ="Texte5"
                            Caption ="Ref :"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =6870
                    Top =5452
                    Width =1935
                    Height =315
                    FontSize =12
                    TabIndex =1
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =5895
                            Top =5452
                            Width =960
                            Height =330
                            FontSize =12
                            Name ="EtiDate"
                            Caption ="Paris le "
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =34
                    Top =4732
                    Width =4485
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtCli_Nom1"
                    ControlSource ="Cli_Nom"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =3288
                    Width =2598
                    Height =1256
                    FontSize =10
                    Name ="EtiEnteteCaen"
                    Caption ="CABINET GREF / N.R.S.\015\0129, rue Pasteur\015\01214000 CAEN\015\012Tel  02.31."
                        "85.52.27\015\012Fax  02.31.86.07.50"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4155
                    Top =1402
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =3
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4155
                    Top =1942
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =4
                    Name ="TxtAC_Adr1"
                    ControlSource ="AC_Adr1"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4155
                    Top =2212
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =5
                    Name ="TxtAC_Adr2"
                    ControlSource ="AC_Adr2"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4155
                    Top =3015
                    Width =975
                    Height =270
                    FontSize =10
                    TabIndex =6
                    Name ="TxtAC_CodePostal"
                    ControlSource ="AC_CodePostal"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =5130
                    Top =3022
                    Width =3825
                    Height =270
                    FontSize =10
                    TabIndex =7
                    Name ="TxtAC_Ville"
                    ControlSource ="AC_Ville"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4155
                    Top =1672
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =8
                    Name ="TxtCli_Int_Nom"
                    ControlSource ="Cli_Int_Nom"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4155
                    Top =2482
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =9
                    Name ="TxtAC_Adr3"
                    ControlSource ="AC_Adr3"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4155
                    Top =2752
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =10
                    Name ="TxtAC_Adr4"
                    ControlSource ="AC_Adr4"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4155
                    Top =3292
                    Width =3195
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    Name ="TxtAC_Pays"
                    ControlSource ="AC_Pays"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1140
                    Top =3922
                    Width =6795
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =12
                    Name ="TxtCourrierAttestationIrrecouvrabilite"
                    ControlSource ="=\"ATTESTATION D'IRRECOUVRABILITE EN \" & UCase(DLookUp(\"Mon_Nom\",\"MONNAIE\","
                        "\"Mon_Id=\" & [Déb_MonnaieEnCours]))"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            RepeatSection = NotDefault
            Height =1140
            BreakLevel =1
            Name ="EntêteGroupe1"
            Begin
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =8505
                    Top =870
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="Texte51"
                    Caption ="Motif"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =4980
                    Top =855
                    Width =2175
                    Height =285
                    FontSize =10
                    Name ="Texte52"
                    Caption ="Montant"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Top =855
                    Width =1320
                    Height =285
                    FontSize =10
                    Name ="EtiTIrr_Déb_Code"
                    Caption ="Débiteur"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =1320
                    Top =855
                    Width =1305
                    Height =285
                    FontSize =10
                    Name ="EtiDéb_Nom"
                    Caption ="Nom"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =566
            Name ="Détail1"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =7155
                    Top =15
                    Width =1896
                    Height =255
                    ColumnOrder =5
                    FontSize =10
                    Name ="TxtCG_Libellé"
                    ControlSource ="CG_Libellé"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =15
                    Top =15
                    Width =1251
                    Height =255
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =1
                    Name ="TxtTIrr_Déb_Code"
                    ControlSource ="TIrr_Déb_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =1275
                    Top =15
                    Width =4011
                    Height =255
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =2
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =5115
                    Top =15
                    Width =1131
                    Height =255
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =3
                    Name ="TxtMontant_m"
                    ControlSource ="=Round([Déb_Mont_Prin_m1],2)-Round(nz(DSum(\"Enc_Prin_m1\",\"ENCAISSEMENT\",\"En"
                        "c_Code_Débiteur='\" & [TIrr_Déb_Code] & \"'\"),0),2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =6225
                    Top =15
                    Width =921
                    Height =255
                    ColumnOrder =1
                    TabIndex =4
                    Name ="TxtMontant_mMonnaie"
                    ControlSource ="=[TxtMontant_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =510
                    Top =285
                    Width =2595
                    Height =225
                    ColumnOrder =0
                    TabIndex =5
                    Name ="TxtDéb_Réf1_Client"
                    ControlSource ="Déb_Réf1_Client"
                End
                Begin Label
                    TextFontFamily =34
                    Left =15
                    Top =285
                    Width =495
                    Height =225
                    Name ="EtiDéb_Réf1_Client"
                    Caption ="Réf1 :"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =3600
                    Top =285
                    Width =1695
                    Height =225
                    ColumnOrder =6
                    TabIndex =6
                    Name ="TxtDéb_Réf2_Client"
                    ControlSource ="Déb_Réf2_Client"
                End
                Begin Label
                    TextFontFamily =34
                    Left =3105
                    Top =285
                    Width =495
                    Height =225
                    Name ="EtiDéb_Réf2_Client"
                    Caption ="Réf2 :"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =1110
            BreakLevel =1
            Name ="PiedGroupe2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =4305
                    Top =424
                    Width =1926
                    Height =345
                    ColumnOrder =1
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtTotalMontant_m"
                    ControlSource ="=Sum(Round([Déb_Mont_Prin_m1],2)-Round(nz(DSum(\"Enc_Prin_m1\",\"ENCAISSEMENT\","
                        "\"Enc_Code_Débiteur='\" & [TIrr_Déb_Code] & \"'\"),0),2))"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =3499
                            Top =423
                            Width =810
                            Height =345
                            FontSize =12
                            FontWeight =700
                            Name ="EtiTotalMontant_m"
                            Caption ="Total :"
                        End
                    End
                End
                Begin Line
                    Top =345
                    Width =9015
                    Name ="Ligne47"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =4302
                    Top =760
                    Width =1926
                    Height =330
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    Name ="TxtTotalMontant_mMonnaie"
                    ControlSource ="=[TxtTotalMontant_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =3499
                            Top =763
                            Width =810
                            Height =345
                            FontSize =12
                            Name ="EtiTxtTotalMontant_mMonnaie"
                            Caption ="soit"
                        End
                    End
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =0
            Name ="PiedGroupe1"
        End
    End
End
CodeBehindForm
' See "EtaCourrierAttestationIrrecouvrabilite.cls"
