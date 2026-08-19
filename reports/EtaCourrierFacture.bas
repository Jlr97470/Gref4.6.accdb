Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DateGrouping =1
    GridX =5
    GridY =5
    Width =9420
    ItemSuffix =92
    Left =-27106
    Top =2700
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xd028ac6e2eefe240
    End
    RecordSource ="SELECT DISTINCTROW FACTURE.*, CLIENT.Cli_Nom, CLIENT.Cli_Groupe, ADRESSE_CLIENT."
        "AC_Adr1, ADRESSE_CLIENT.AC_Adr2, ADRESSE_CLIENT.AC_Adr3, ADRESSE_CLIENT.AC_Adr4,"
        " ADRESSE_CLIENT.AC_CodePostal, ADRESSE_CLIENT.AC_Ville, ADRESSE_CLIENT.AC_Pays, "
        "FACTURE_DETAIL.*, DEBITEUR.Déb_Code, DEBITEUR.Déb_Nom, DEBITEUR.Déb_Réf1_Client,"
        " DEBITEUR.Déb_Réf2_Client FROM ((CLIENT INNER JOIN ADRESSE_CLIENT ON CLIENT.Cli_"
        "Code=ADRESSE_CLIENT.AC_Code_Client) INNER JOIN FACTURE ON CLIENT.Cli_Code=FACTUR"
        "E.Fac_CodeCli) INNER JOIN (DEBITEUR INNER JOIN FACTURE_DETAIL ON DEBITEUR.Déb_Co"
        "de=FACTURE_DETAIL.FacD_Déb_Code) ON FACTURE.Fac_Id=FACTURE_DETAIL.FacD_Fac_Id WH"
        "ERE (((ADRESSE_CLIENT.AC_TypeAdresse_ID)=DLookUp(\"Adresse_Courrier\",\"PARAMETR"
        "E_GREF\")) AND ((FACTURE.Fac_Editer)=True)); "
    OnOpen ="[Event Procedure]"
    FilterOnLoad =255
    AllowLayoutView =0
    Begin
        Begin Label
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Rectangle
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin BreakLevel
            ControlSource ="Fac_Date"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Fac_Num"
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="FacD_ETy_Mode"
        End
        Begin BreakLevel
            ControlSource ="Déb_Nom"
        End
        Begin PageHeader
            Height =1133
            Name ="EntêtePage2"
        End
        Begin BreakHeader
            CanGrow = NotDefault
            Height =3975
            BreakLevel =1
            Name ="EntêteGroupe"
            Begin
                Begin Label
                    TextFontFamily =18
                    Left =2955
                    Top =2445
                    Width =540
                    Height =330
                    FontSize =12
                    Name ="Texte5"
                    Caption ="Ref :"
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextFontFamily =18
                    Top =2445
                    Width =840
                    Height =330
                    FontSize =12
                    Name ="EtiFac_Date"
                    Caption ="Paris le "
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =3510
                    Top =2445
                    Width =1680
                    Height =315
                    FontSize =12
                    Name ="TxtFac_CodeCli"
                    ControlSource ="Fac_CodeCli"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =855
                    Top =2445
                    Width =1935
                    Height =315
                    FontSize =12
                    TabIndex =1
                    Name ="TxtFac_Date"
                    ControlSource ="Fac_Date"
                    Format ="Short Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4590
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =2
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4590
                    Top =270
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =3
                    Name ="TxtAC_Adr1"
                    ControlSource ="AC_Adr1"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4590
                    Top =540
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =4
                    Name ="TxtAC_Adr2"
                    ControlSource ="AC_Adr2"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4590
                    Top =1353
                    Width =960
                    Height =270
                    FontSize =10
                    TabIndex =5
                    Name ="TxtAC_CodePostal"
                    ControlSource ="AC_CodePostal"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =5565
                    Top =1350
                    Width =3825
                    Height =270
                    FontSize =10
                    TabIndex =6
                    Name ="TxtAC_Ville"
                    ControlSource ="AC_Ville"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4590
                    Top =810
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =7
                    Name ="TxtAC_Adr3"
                    ControlSource ="AC_Adr3"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4590
                    Top =1080
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =8
                    Name ="TxtAC_Adr4"
                    ControlSource ="AC_Adr4"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4590
                    Top =1620
                    Width =3195
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    Name ="TxtAC_Pays"
                    ControlSource ="AC_Pays"

                End
                Begin Rectangle
                    Top =2940
                    Width =9071
                    Height =355
                    Name ="Boîte83"
                End
                Begin Label
                    TextFontFamily =18
                    Left =3000
                    Top =2955
                    Width =1560
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Étiquette84"
                    Caption ="FACTURE N°"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =4530
                    Top =2955
                    Width =1680
                    Height =330
                    FontSize =12
                    TabIndex =10
                    Name ="TxtFac_Num"
                    ControlSource ="Fac_Num"
                    FontName ="Times New Roman"

                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =18
                    Left =7110
                    Top =3405
                    Width =1770
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="EtiFac_Montant_ho"
                    Caption ="HONORAIRES HT"
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =18
                    Left =5100
                    Top =3405
                    Width =1770
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="EtiFac_Montant_m"
                    Caption ="REGLEMENTS"
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =1770
                    Top =3405
                    Width =1770
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="EtiDéb_Nom"
                    Caption ="NOM"
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =255
                    Top =3405
                    Width =1410
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="EtiDéb_Code"
                    Caption ="REFERENCES"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =525
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =105
                    Top =15
                    Width =1131
                    Height =255
                    FontSize =10
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =1470
                    Top =15
                    Width =3741
                    Height =255
                    FontSize =10
                    TabIndex =1
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =18
                    Left =5400
                    Width =1326
                    Height =255
                    FontSize =10
                    TabIndex =2
                    Name ="TxtFacD_Montant_m"
                    ControlSource ="FacD_Montant_m1"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =600
                    Top =300
                    Width =2880
                    Height =225
                    TabIndex =3
                    Name ="TxtDéb_Réf1_Client"
                    ControlSource ="Déb_Réf1_Client"

                End
                Begin Label
                    TextFontFamily =34
                    Left =105
                    Top =300
                    Width =495
                    Height =225
                    Name ="EtiDéb_Réf1_Client"
                    Caption ="Réf1 :"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4020
                    Top =300
                    Width =2325
                    Height =225
                    TabIndex =4
                    Name ="TxtDéb_Réf2_Client"
                    ControlSource ="Déb_Réf2_Client"

                End
                Begin Label
                    TextFontFamily =34
                    Left =3525
                    Top =300
                    Width =495
                    Height =225
                    Name ="EtiDéb_Réf2_Client"
                    Caption ="Réf2 :"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =18
                    Left =7380
                    Width =1326
                    Height =255
                    FontSize =10
                    TabIndex =5
                    Name ="TxtFacD_Montant_ho"
                    ControlSource ="=Round([FacD_Montant_m1],2)*[Fac_Honoraire]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =855
            BreakLevel =2
            Name ="PiedGroupe"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =5010
                    Top =30
                    Width =1710
                    Height =285
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtTotalFacD_Montant_m"
                    ControlSource ="=nz(Sum(Round([FacD_Montant_m1],2)),0)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =6983
                    Top =300
                    Width =1710
                    Height =285
                    ColumnOrder =1
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtTotalFacD_Montant_ho"
                    ControlSource ="=nz(Sum(Round([FacD_Montant_m1],2)),0)*[Fac_Honoraire]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =1140
                            Top =300
                            Width =3390
                            Height =285
                            FontSize =10
                            FontWeight =700
                            Name ="EtiTotalFacD_Montant_ho"
                            Caption ="TOTAL HONORAIRES HT :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =8797
                    Name ="Trait62"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =1142
                    Top =30
                    Width =3390
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="EtiTotalFacD_Montant_m"
                    ControlSource ="=IIf([FacD_Ety_Mode]=1,\"TOTAL ENCAISSE :\",\"TOTAL REGLEMENTS DIRECTS :\")"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =1905
            BreakLevel =1
            Name ="PiedGroupe1"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =1305
                    Top =450
                    Width =1341
                    Height =330
                    ColumnOrder =14
                    FontSize =12
                    Name ="TxtTotalFacD_Montant_mEncaisserTTC"
                    ControlSource ="=[TxtTotalFacD_Montant_mEncaisser]*[Fac_Honoraire]/100*(1+[Fac_TVA])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =15
                            Top =450
                            Width =1245
                            Height =330
                            FontSize =12
                            Name ="EtiTxtTotalFacD_Montant_mEncaisserTTC"
                            Caption ="Débit TTC :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =6885
                    Top =450
                    Width =1461
                    Height =330
                    ColumnOrder =13
                    FontSize =12
                    TabIndex =1
                    Name ="TxtTotalFacD_Montant_mEncaisser"
                    ControlSource ="=Nz(Sum(IIf([FacD_ETy_Mode]=1,Round([FacD_Montant_m1],2),0)),0)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =4650
                            Top =450
                            Width =2190
                            Height =330
                            FontSize =12
                            Name ="EtiTotalFacD_Montant_mEncaisser"
                            Caption ="du montant encaissé :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =3585
                    Top =450
                    Width =846
                    Height =330
                    ColumnOrder =12
                    FontSize =12
                    TabIndex =2
                    Name ="TxtFac_Honoraire"
                    ControlSource ="=[Fac_Honoraire]/100"
                    Format ="Percent"
                    FontName ="Times New Roman"

                End
                Begin Rectangle
                    BackStyle =0
                    Top =375
                    Width =9420
                    Height =1530
                    Name ="Boîte39"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =1305
                    Top =1485
                    Width =1341
                    Height =315
                    ColumnOrder =11
                    FontSize =12
                    TabIndex =3
                    Name ="TxtTotalFacD_Montant_mTVA"
                    ControlSource ="=[TxtTotalFacD_Montant_mTTC]*[Fac_TVA]/(1+[Fac_TVA])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =15
                            Top =1485
                            Width =1245
                            Height =330
                            FontSize =12
                            Name ="EtiTotalFacD_Montant_mTVA"
                            Caption ="Dont TVA :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =6885
                    Top =1485
                    Width =1461
                    Height =330
                    ColumnOrder =10
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    Name ="TxtMontantReverser"
                    ControlSource ="=[TxtTotalFacD_Montant_mTTC]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =4920
                            Top =1485
                            Width =1920
                            Height =330
                            FontSize =12
                            FontWeight =700
                            Name ="Texte43"
                            Caption ="Montant à régler :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =1305
                    Top =780
                    Width =1341
                    Height =330
                    ColumnOrder =9
                    FontSize =12
                    TabIndex =5
                    Name ="TxtTotalFacD_Montant_mDirectTTC"
                    ControlSource ="=[TxtTotalFacD_Montant_mDirect]*[Fac_Honoraire]/100*(1+[Fac_TVA])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =15
                            Top =780
                            Width =1245
                            Height =330
                            FontSize =12
                            Name ="EtiTotalFacD_Montant_mDirectTTC"
                            Caption ="Débit TTC :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =6885
                    Top =780
                    Width =1461
                    Height =330
                    ColumnOrder =8
                    FontSize =12
                    TabIndex =6
                    Name ="TxtTotalFacD_Montant_mDirect"
                    ControlSource ="=Nz(Sum(IIf([FacD_ETy_Mode]=2,Round([FacD_Montant_m1],2),0)),0)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =4470
                            Top =780
                            Width =2370
                            Height =330
                            FontSize =12
                            Name ="EtiTotalFacD_Montant_mDirect"
                            Caption ="des règlements directs :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =3585
                    Top =780
                    Width =846
                    Height =330
                    ColumnOrder =7
                    FontSize =12
                    TabIndex =7
                    Name ="TxtFac_Honoraire1"
                    ControlSource ="=[Fac_Honoraire]/100"
                    Format ="Percent"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =1305
                    Top =1110
                    Width =1341
                    Height =330
                    ColumnOrder =6
                    FontSize =12
                    TabIndex =8
                    Name ="TxtTotalFacD_Montant_mTTC"
                    ControlSource ="=Round([TxtTotalFacD_Montant_mDirectTTC],2)+Round([TxtTotalFacD_Montant_mEncaiss"
                        "erTTC],2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =15
                            Top =1110
                            Width =1245
                            Height =330
                            FontSize =12
                            Name ="EtiTotalFacD_Montant_mTTC"
                            Caption ="Total TTC :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =8385
                    Top =525
                    Width =1011
                    Height =225
                    ColumnOrder =5
                    TabIndex =9
                    Name ="TxtTotalFacD_Montant_mEncaisserMonnaie"
                    ControlSource ="=[TxtTotalFacD_Montant_mEncaisser]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =8355
                    Top =1560
                    Width =1041
                    Height =225
                    ColumnOrder =4
                    FontWeight =700
                    TabIndex =10
                    Name ="TxtMontantReverserMonnaie"
                    ControlSource ="=[TxtMontantReverser]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =8385
                    Top =855
                    Width =1011
                    Height =225
                    ColumnOrder =3
                    TabIndex =11
                    Name ="TxtTotalFacD_Montant_mDirectMonnaie"
                    ControlSource ="=[TxtTotalFacD_Montant_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =2655
                    Top =525
                    Width =801
                    Height =225
                    ColumnOrder =2
                    TabIndex =12
                    Name ="TxtTotalFacD_Montant_mEncaisserTTCMonnaie"
                    ControlSource ="=([TxtTotalFacD_Montant_mEncaisserTTC])*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =2655
                    Top =1560
                    Width =801
                    Height =225
                    ColumnOrder =1
                    TabIndex =13
                    Name ="TxtTotalFacD_Montant_mTVAMonnaie"
                    ControlSource ="=([TxtTotalFacD_Montant_mTVA])*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =2655
                    Top =855
                    Width =801
                    Height =225
                    ColumnOrder =0
                    TabIndex =14
                    Name ="TxtTotalFacD_Montant_mDirectTTCMonnaie"
                    ControlSource ="=([TxtTotalFacD_Montant_mDirectTTC])*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =2655
                    Top =1185
                    Width =801
                    Height =225
                    ColumnOrder =15
                    TabIndex =15
                    Name ="TxtTotalFacD_Montant_mTTCMonnaie"
                    ControlSource ="=[TxtTotalFacD_Montant_mTTC]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"

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
' See "EtaCourrierFacture.cls"
