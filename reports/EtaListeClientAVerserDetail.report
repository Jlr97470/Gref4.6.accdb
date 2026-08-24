Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    FilterOn = NotDefault
    DateGrouping =1
    GridX =5
    GridY =5
    Width =9060
    ItemSuffix =67
    Left =3270
    Top =1755
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xbaa8039eefece240
    End
    RecordSource ="SELECT DISTINCTROW CLIENT.Cli_Code, CLIENT.Cli_Nom, DEBITEUR.Déb_Code, ENCAISSEM"
        "ENT.Enc_Prin_m1, ENCAISSEMENT.Enc_Prin_m2, ENCAISSEMENT.Enc_Date, ENCAISSEMENT.E"
        "nc_Montant_mo FROM ENCAISSEMENT_TYPE INNER JOIN (CLIENT INNER JOIN (DEBITEUR INN"
        "ER JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code=ENCAISSEMENT.Enc_Code_Débiteur) ON CLI"
        "ENT.Cli_Code=DEBITEUR.Déb_Code_Client) ON ENCAISSEMENT_TYPE.ETy_Code=ENCAISSEMEN"
        "T.Enc_Type WHERE (((ENCAISSEMENT.Enc_Prin_m1)<>0) AND ((ENCAISSEMENT_TYPE.ETy_Mo"
        "de)<>2) AND ((ENCAISSEMENT.Enc_Situation) Is Null)); "
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
            ControlSource ="Cli_Code"
        End
        Begin BreakLevel
            ControlSource ="Déb_Code"
        End
        Begin BreakLevel
            ControlSource ="Enc_Date"
        End
        Begin FormHeader
            Height =1700
            Name ="EntêteEtat0"
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Top =491
                    Width =5943
                    Height =480
                    FontSize =14
                    FontWeight =700
                    Name ="EtiListeClientAVerserDetail"
                    Caption ="Liste des Clients à verser :"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =6066
                    Top =545
                    Width =2955
                    Height =285
                    FontSize =10
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                End
                Begin Line
                    BorderWidth =2
                    Top =435
                    Width =9018
                    Name ="TrtListeClientAVerserDetail"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =680
            Name ="EntêteGroupe0"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =7410
                    Top =300
                    Width =1341
                    Height =285
                    FontSize =10
                    Name ="TxtCli_Date_der_reg"
                    ControlSource ="=DLast(\"Reg_Date\",\"REGLEMENT\",\"Reg_Cli_Code='\" & [Cli_Code] & \"'\")"
                    Format ="Short Date"
                End
                Begin Label
                    TextFontFamily =34
                    Left =5805
                    Top =300
                    Width =1560
                    Height =285
                    FontSize =10
                    Name ="EtiCli_Date_der_reg"
                    Caption ="Der. versement :"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =345
                    Top =300
                    Width =1131
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="TxtCli_Code"
                    ControlSource ="Cli_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =1530
                    Top =300
                    Width =4206
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =225
                    Top =240
                    Width =8595
                    Height =0
                    Name ="Boîte46"
                End
            End
        End
        Begin Section
            Height =226
            Name ="Détail1"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =1695
                    Width =1101
                    Height =225
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =3862
                    Width =1476
                    Height =225
                    TabIndex =1
                    Name ="TxtEnc_Prin_m"
                    ControlSource ="Enc_Prin_m1"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =7370
                    Width =1341
                    Height =225
                    TabIndex =2
                    Name ="TxtEnc_Date"
                    ControlSource ="Enc_Date"
                    Format ="Short Date"
                End
                Begin Label
                    TextFontFamily =34
                    Left =915
                    Width =735
                    Height =225
                    Name ="EtiDéb_Code"
                    Caption ="Débiteur :"
                End
                Begin Label
                    TextFontFamily =34
                    Left =3225
                    Width =630
                    Height =225
                    Name ="EtiEnc_Prin_m"
                    Caption ="a payé :"
                End
                Begin Label
                    TextFontFamily =34
                    Left =7029
                    Width =270
                    Height =225
                    Name ="EtiEnc_Date"
                    Caption ="le :"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =5790
                    Width =1191
                    Height =225
                    TabIndex =3
                    Name ="TxtEnc_Prin_mMonnaie"
                    ControlSource ="Enc_Prin_m2"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
                Begin Label
                    TextFontFamily =34
                    Left =5475
                    Width =315
                    Height =225
                    Name ="EtiEnc_Prin_mMonnaie"
                    Caption ="soit"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =375
            Name ="PiedGroupe0"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =3877
                    Top =90
                    Width =1476
                    Height =285
                    FontSize =10
                    Name ="TxtTotalEnc_Prin_m"
                    ControlSource ="=Sum([Enc_Prin_m1])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =2160
                    Top =90
                    Width =1710
                    Height =285
                    FontSize =10
                    Name ="EtiTotalEnc_Prin_m"
                    Caption ="à verser au client :"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =5430
                    Top =90
                    Width =1551
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="TxtTotalEnc_Prin_mMonnaie"
                    ControlSource ="=[TxtTotalEnc_Prin_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =990
            Name ="PiedEtat1"
            Begin
                Begin Line
                    Left =180
                    Top =578
                    Width =8610
                    Name ="Ligne47"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =2190
                    Top =645
                    Width =2136
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="TxtTotal"
                    ControlSource ="=Sum([Enc_Prin_m1])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Top =645
                            Width =2145
                            Height =330
                            FontSize =12
                            FontWeight =700
                            Name ="EtiTotal"
                            Caption ="Total Général :"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =4455
                    Top =645
                    Width =510
                    Height =345
                    FontSize =12
                    FontWeight =700
                    Name ="Étiquette66"
                    Caption ="soit"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =5070
                    Top =645
                    Width =2196
                    Height =330
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtTotalMonnaie"
                    ControlSource ="=[TxtTotal]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
            End
        End
    End
End
CodeBehindForm
' See "EtaListeClientAVerserDetail.cls"
