Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DateGrouping =1
    GridX =5
    GridY =5
    Width =9348
    ItemSuffix =61
    Left =495
    Top =1725
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x421259952de8e240
    End
    RecordSource ="SELECT DISTINCTROW CLIENT.Cli_Code, CLIENT.Cli_Nom, CLIENT.Cli_MonnaieEnCours, D"
        "EBITEUR.Déb_Code, DEBITEUR.Déb_Nom, DEBITEUR.Déb_Date_Cré, DEBITEUR.Déb_Réf1_Cli"
        "ent, DEBITEUR.Déb_Réf2_Client, ENCAISSEMENT.Enc_Date, ENCAISSEMENT.Enc_Type, ENC"
        "AISSEMENT.Enc_Prin_m1, ENCAISSEMENT.Enc_Prin_m2 FROM ((CLIENT INNER JOIN ADRESSE"
        "_CLIENT ON CLIENT.Cli_Code=ADRESSE_CLIENT.AC_Code_Client) INNER JOIN (DEBITEUR I"
        "NNER JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code=ENCAISSEMENT.Enc_Code_Débiteur) ON C"
        "LIENT.Cli_Code=DEBITEUR.Déb_Code_Client) INNER JOIN TMP_REGLEMENT ON ENCAISSEMEN"
        "T.Enc_ID=TMP_REGLEMENT.TReg_Enc_ID WHERE (((TMP_REGLEMENT.TReg_Flag)=True) AND ("
        "(ADRESSE_CLIENT.AC_TypeAdresse_ID)=DLookUp(\"Adresse_Courrier\",\"PARAMETRE_GREF"
        "\"))); "
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
            KeepTogether =1
            ControlSource ="Cli_Code"
        End
        Begin BreakLevel
            ControlSource ="Déb_Code"
        End
        Begin PageHeader
            Height =1474
            Name ="EntêtePage"
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =30
                    Top =622
                    Width =9258
                    Height =480
                    FontSize =14
                    FontWeight =700
                    Name ="EtiBrouillardReglement"
                    Caption ="Brouillard des Règlements : "
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =5865
                    Top =671
                    Width =3405
                    Height =285
                    FontSize =10
                    Name ="TxtDateJour"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                End
                Begin Line
                    BorderWidth =2
                    Top =566
                    Width =9348
                    Name ="TrtBroulliardReglement"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =1680
            Name ="EntêteGroupe"
            Begin
                Begin Label
                    TextFontFamily =34
                    Left =570
                    Top =705
                    Width =2025
                    Height =330
                    FontSize =12
                    Name ="EtiCli_Code"
                    Caption ="Référence Client :"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =3780
                    Top =705
                    Width =5040
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =2610
                    Top =705
                    Width =1140
                    Height =330
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtCli_Code"
                    ControlSource ="Cli_Code"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =480
                    Top =645
                    Width =8370
                    Height =480
                    Name ="RecCli"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =4095
                    Top =1365
                    Width =1005
                    Height =285
                    FontSize =10
                    Name ="EtiDéb_Date_Cré"
                    Caption ="Créance"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =5310
                    Top =1365
                    Width =1050
                    Height =285
                    FontSize =10
                    Name ="EtiEnc_Date"
                    Caption ="Encaiss."
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =270
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =15
                    Top =15
                    Width =1311
                    Height =255
                    FontSize =10
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =1335
                    Top =15
                    Width =2781
                    Height =255
                    FontSize =10
                    TabIndex =1
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =5310
                    Top =15
                    Width =1041
                    Height =255
                    FontSize =10
                    TabIndex =2
                    Name ="TxtEnc_Date"
                    ControlSource ="Enc_Date"
                    Format ="Short Date"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =34
                    Left =6360
                    Top =15
                    Width =1311
                    Height =255
                    FontSize =10
                    TabIndex =3
                    Name ="TxtEnc_Prin_m"
                    ControlSource ="Enc_Prin_m1"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =4110
                    Top =15
                    Width =1011
                    Height =255
                    FontSize =10
                    TabIndex =4
                    Name ="TxtDéb_Date_Cré"
                    ControlSource ="Déb_Date_Cré"
                    Format ="Short Date"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =8775
                    Top =15
                    Width =531
                    Height =255
                    FontSize =10
                    TabIndex =5
                    Name ="TxtEnc_Type"
                    ControlSource ="Enc_Type"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =34
                    Left =7680
                    Width =1086
                    Height =270
                    TabIndex =6
                    Name ="TxtEnc_Prin_mMonnaie"
                    ControlSource ="=[TxtEnc_Prin_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =1474
            Name ="PiedGroupe"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =5625
                    Top =675
                    Width =2031
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="TxtTotalEnc_Prin_m"
                    ControlSource ="=Round(Nz(Sum([Enc_Prin_m1]),0),2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =3435
                            Top =675
                            Width =2145
                            Height =330
                            FontSize =12
                            FontWeight =700
                            Name ="EtiTotalEnc_Prin_m"
                            Caption ="Total encaissé :"
                        End
                    End
                End
                Begin Line
                    Left =60
                    Top =1360
                    Width =9225
                    Name ="TrtMontant"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =5625
                    Top =1005
                    Width =2031
                    Height =330
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtTotalEnc_Prin_mMonnaie"
                    ControlSource ="=[TxtTotalEnc_Prin_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =3435
                            Top =1005
                            Width =2145
                            Height =330
                            FontSize =12
                            Name ="EtiTotalEnc_Prin_mMonnaie"
                            Caption ="soit"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PiedPage"
        End
    End
End
CodeBehindForm
' See "EtaBrouillardReglement.cls"
