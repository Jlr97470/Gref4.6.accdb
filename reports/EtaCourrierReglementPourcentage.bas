Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DateGrouping =1
    GridX =5
    GridY =5
    Width =9045
    ItemSuffix =65
    Left =75
    Top =1560
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xe7b234f155efe240
    End
    RecordSource ="SELECT CLIENT.Cli_Code, CLIENT.Cli_Nom, ADRESSE_CLIENT.AC_Adr1, ADRESSE_CLIENT.A"
        "C_Adr2, ADRESSE_CLIENT.AC_Adr3, ADRESSE_CLIENT.AC_Adr4, ADRESSE_CLIENT.AC_Ville,"
        " ADRESSE_CLIENT.AC_CodePostal, ADRESSE_CLIENT.AC_Pays, REGLEMENT.Reg_Date, REGLE"
        "MENT.Reg_Cli_Code, REGLEMENT.Reg_Cli_Pourcentage, REGLEMENT.Reg_TVA, REGLEMENT.R"
        "eg_Montant_mo, REGLEMENT_DETAIL.RegD_Id, REGLEMENT_DETAIL.RegD_Déb_Code, REGLEME"
        "NT_DETAIL.RegD_ETy_Mode, REGLEMENT_DETAIL.RegD_Type, REGLEMENT_DETAIL.RegD_Monta"
        "nt_m1, REGLEMENT_DETAIL.RegD_Montant_m2, DEBITEUR.Déb_Nom, DEBITEUR.Déb_Réf1_Cli"
        "ent, DEBITEUR.Déb_Réf2_Client, DEBITEUR.Déb_Date_Cré, DEBITEUR.Déb_Mont_Prin_m1,"
        " DEBITEUR.Déb_Mont_Prin_m2, DEBITEUR.Déb_Date_Chèque FROM ((CLIENT INNER JOIN AD"
        "RESSE_CLIENT ON CLIENT.Cli_Code=ADRESSE_CLIENT.AC_Code_Client) INNER JOIN REGLEM"
        "ENT ON CLIENT.Cli_Code=REGLEMENT.Reg_Cli_Code) INNER JOIN (DEBITEUR INNER JOIN R"
        "EGLEMENT_DETAIL ON DEBITEUR.Déb_Code=REGLEMENT_DETAIL.RegD_Déb_Code) ON REGLEMEN"
        "T.Reg_Id=REGLEMENT_DETAIL.RegD_Reg_Id WHERE (((REGLEMENT.Reg_Cli_Pourcentage)>0)"
        " AND ((ADRESSE_CLIENT.AC_TypeAdresse_ID)=DLookUp(\"Adresse_Courrier\",\"PARAMETR"
        "E_GREF\")) AND ((REGLEMENT.Reg_Editer)=True)); "
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
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Reg_Date"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Cli_Code"
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="RegD_ETy_Mode"
        End
        Begin BreakLevel
            ControlSource ="Déb_Nom"
        End
        Begin PageHeader
            Height =1133
            Name ="EntêtePage"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="EntêteGroupe"
        End
        Begin BreakHeader
            CanGrow = NotDefault
            ForceNewPage =1
            Height =2790
            BreakLevel =1
            Name ="EntêteGroupe1"
            Begin
                Begin Label
                    TextFontFamily =18
                    Left =2955
                    Top =2445
                    Width =540
                    Height =330
                    FontSize =12
                    Name ="EtiReg_Cli_Code"
                    Caption ="Ref :"
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextFontFamily =18
                    Top =2445
                    Width =840
                    Height =330
                    FontSize =12
                    Name ="EtiParis"
                    Caption ="Paris le "
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =3495
                    Top =2445
                    Width =1695
                    Height =330
                    FontSize =12
                    Name ="TxtReg_Cli_Code"
                    ControlSource ="Reg_Cli_Code"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =840
                    Top =2445
                    Width =1950
                    Height =330
                    FontSize =12
                    TabIndex =1
                    Name ="EtiReg_Date"
                    ControlSource ="Reg_Date"
                    Format ="Short Date"
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =7935
                    Top =2265
                    Width =765
                    Height =510
                    FontSize =10
                    Name ="EtiResteAVerser"
                    Caption ="Reste à verser"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4230
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =2
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4230
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
                    Left =4230
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
                    Left =4230
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
                    Left =5205
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
                    Left =4230
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
                    Left =4230
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
                    Left =4230
                    Top =1620
                    Width =3195
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    Name ="TxtAC_Pays"
                    ControlSource ="AC_Pays"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =690
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Width =1146
                    Height =255
                    FontSize =10
                    Name ="TxtRegD_Déb_Code"
                    ControlSource ="RegD_Déb_Code"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =1140
                    Width =3051
                    Height =255
                    FontSize =10
                    TabIndex =1
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =5310
                    Width =1026
                    Height =255
                    FontSize =10
                    TabIndex =2
                    Name ="TxtDéb_Date_Cré"
                    ControlSource ="Déb_Date_Cré"
                    Format ="Short Date"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =18
                    Left =6330
                    Width =1341
                    Height =255
                    FontSize =10
                    TabIndex =3
                    Name ="TxtRegD_Montant_m"
                    ControlSource ="RegD_Montant_m1"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =18
                    Left =4185
                    Width =1131
                    Height =255
                    FontSize =10
                    TabIndex =4
                    Name ="TxtRegD_Type"
                    ControlSource ="RegD_Type"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =18
                    Left =7665
                    Width =1326
                    Height =255
                    FontSize =10
                    TabIndex =5
                    Name ="TxtResteAVerser"
                    ControlSource ="=IIf([RegD_Ety_Mode]=1,Round([Déb_Mont_Prin_m1],2)-Round(Nz(DSum(\"Enc_Prin_m1\""
                        ",\"ENCAISSEMENT\",\"Enc_RegD_Id<=\" & [RegD_Id] & \" AND Enc_Code_Débiteur='\" &"
                        " [RegD_Déb_Code] & \"'\"),0),2),Null)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =18
                    Left =6330
                    Top =270
                    Width =1341
                    Height =255
                    TabIndex =6
                    Name ="TxtRegD_Montant_mMonnaie"
                    ControlSource ="=[TxtRegD_Montant_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =18
                    Left =7665
                    Top =270
                    Width =1326
                    Height =255
                    TabIndex =7
                    Name ="TxtResteAVerserMonnaie"
                    ControlSource ="=IIf([RegD_ETy_Mode]=1,[TxtResteAVerser]*DonnerTauxConvertion(1),Null)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =495
                    Top =270
                    Width =2865
                    Height =255
                    TabIndex =8
                    Name ="TxtDéb_Réf1_Client"
                    ControlSource ="Déb_Réf1_Client"
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Top =270
                            Width =495
                            Height =255
                            Name ="EtiDéb_Réf1_Client"
                            Caption ="Réf1 :"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =3915
                    Top =270
                    Width =1695
                    Height =255
                    TabIndex =9
                    Name ="TxtDéb_Réf2_Client"
                    ControlSource ="Déb_Réf2_Client"
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =3360
                            Top =270
                            Width =555
                            Height =255
                            Name ="EtiDéb_Réf2_Client"
                            Caption ="Réf2 :"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5610
                    Top =270
                    Width =726
                    Height =255
                    TabIndex =10
                    Name ="TxtDéb_Date_Chèque"
                    ControlSource ="Déb_Date_Chèque"
                    StatusBarText ="Debiteur Date Cheque"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =690
            BreakLevel =2
            Name ="PiedGroupe"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Left =6360
                    Top =225
                    Width =1290
                    Height =285
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =1
                    Name ="TxtTotalRegD_Montant_m"
                    ControlSource ="=Nz(Sum(Round([RegD_Montant_m1],2)),0)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =8145
                    Top =225
                    Width =840
                    Height =285
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =2
                    Name ="TxtTotalRegD_Montant_mMonnaie"
                    ControlSource ="=[TxtTotalRegD_Montant_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =7710
                            Top =225
                            Width =360
                            Height =285
                            FontSize =10
                            Name ="EtiSoit"
                            Caption ="soit"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3195
                    Top =225
                    Width =3105
                    Height =285
                    FontSize =10
                    Name ="TxtTotalLibelle"
                    ControlSource ="=IIf([RegD_ETy_Mode]=1,\"TOTAL ENCAISSE\",\"TOTAL REGLEMENTS DIRECTS\")"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =1815
            BreakLevel =1
            Name ="PiedGroupe1"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Left =1335
                    Top =255
                    Width =1281
                    Height =330
                    ColumnOrder =14
                    FontSize =12
                    Name ="TxtTotalRegD_Montant_mEncaisserTTC"
                    ControlSource ="=[TxtTotalRegD_Montant_mEncaisser]*[Reg_Cli_Pourcentage]/100*(1+[Reg_TVA])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =45
                            Top =255
                            Width =1290
                            Height =330
                            FontSize =12
                            Name ="EtiTotalRegD_Montant_mEncaisserTTC"
                            Caption ="Débit TTC :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =7680
                    Top =240
                    Width =1236
                    Height =330
                    ColumnOrder =13
                    FontSize =12
                    TabIndex =1
                    Name ="TxtTotalRegD_Montant_mEncaisser"
                    ControlSource ="=Nz(Sum(IIf([RegD_ETy_Mode]=1,Round([RegD_Montant_m1],2),0)),0)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =4455
                            Top =240
                            Width =2325
                            Height =330
                            FontSize =12
                            Name ="EtiTotalRegD_Montant_mEncaisser"
                            Caption ="du montant encaissé :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =3555
                    Top =240
                    Width =816
                    Height =330
                    ColumnOrder =12
                    FontSize =12
                    TabIndex =2
                    Name ="TxtReg_Cli_Pourcentage"
                    ControlSource ="=[Reg_Cli_Pourcentage]/100"
                    Format ="Percent"
                    FontName ="Times New Roman"
                End
                Begin Rectangle
                    BackStyle =0
                    Top =90
                    Width =9000
                    Height =1635
                    Name ="RecMontant"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =1335
                    Top =1290
                    Width =1281
                    Height =330
                    ColumnOrder =11
                    FontSize =12
                    TabIndex =3
                    Name ="TxtTotalRegD_Montant_mTVA"
                    ControlSource ="=[TxtTotalRegD_Montant_mTTC]*[Reg_TVA]/(1+[Reg_TVA])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =45
                            Top =1290
                            Width =1290
                            Height =330
                            FontSize =12
                            Name ="EtiTotalRegD_Montant_mTVA"
                            Caption ="Dont TVA :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =7665
                    Top =1260
                    Width =1251
                    Height =330
                    ColumnOrder =10
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    Name ="TxtMontantReverser"
                    ControlSource ="=[TxtTotalRegD_Montant_mEncaisser]-[TxtTotalRegD_Montant_mTTC]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =4455
                            Top =1275
                            Width =2325
                            Height =330
                            FontSize =12
                            FontWeight =700
                            Name ="EtiMontantReverser"
                            Caption ="Montant reversé :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =1335
                    Top =585
                    Width =1281
                    Height =330
                    ColumnOrder =9
                    FontSize =12
                    TabIndex =5
                    Name ="TxtTotalRegD_Montant_mDirectTTC"
                    ControlSource ="=[TxtTotalRegD_Montant_mDirect]*[Reg_Cli_Pourcentage]/100*(1+[Reg_TVA])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =45
                            Top =585
                            Width =1290
                            Height =330
                            FontSize =12
                            Name ="EtiTxtTotalRegD_Montant_mDirectTTC"
                            Caption ="Débit TTC :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =7680
                    Top =570
                    Width =1236
                    Height =330
                    ColumnOrder =8
                    FontSize =12
                    TabIndex =6
                    Name ="TxtTotalRegD_Montant_mDirect"
                    ControlSource ="=Nz(Sum(IIf([RegD_ETy_Mode]=2,Round([RegD_Montant_m1],2),0)),0)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =4455
                            Top =570
                            Width =2325
                            Height =330
                            FontSize =12
                            Name ="EtiTotalRegD_Montant_mDirect"
                            Caption ="des règlements directs :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =3555
                    Top =570
                    Width =816
                    Height =330
                    ColumnOrder =7
                    FontSize =12
                    TabIndex =7
                    Name ="TxtReg_Cli_Pourcentage1"
                    ControlSource ="=[Reg_Cli_Pourcentage]/100"
                    Format ="Percent"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =1335
                    Top =915
                    Width =1281
                    Height =330
                    ColumnOrder =6
                    FontSize =12
                    TabIndex =8
                    Name ="TxtTotalRegD_Montant_mTTC"
                    ControlSource ="=[TxtTotalRegD_Montant_mDirectTTC]+[TxtTotalRegD_Montant_mEncaisserTTC]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =45
                            Top =915
                            Width =1290
                            Height =330
                            FontSize =12
                            Name ="EtiTotalRegD_Montant_mTTC"
                            Caption ="Total TTC :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =6810
                    Top =300
                    Width =816
                    Height =225
                    ColumnOrder =5
                    TabIndex =9
                    Name ="TxtTotalRegD_Montant_mEncaisserMonnaie"
                    ControlSource ="=[TxtTotalRegD_Montant_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =6810
                    Top =1320
                    Width =801
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
                    TextFontFamily =18
                    Left =6810
                    Top =645
                    Width =816
                    Height =225
                    ColumnOrder =3
                    TabIndex =11
                    Name ="TxtTotalRegD_Montant_mDirectMonnaie"
                    ControlSource ="=[TxtTotalRegD_Montant_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =2655
                    Top =300
                    Width =816
                    Height =225
                    ColumnOrder =2
                    TabIndex =12
                    Name ="TxtTotalRegD_Montant_mEncaisserTTCMonnaie"
                    ControlSource ="=[TxtTotalRegD_Montant_mEncaisserTTC]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =2655
                    Top =1335
                    Width =816
                    Height =225
                    ColumnOrder =1
                    TabIndex =13
                    Name ="TxtTotalRegD_Montant_mTVAMonnaie"
                    ControlSource ="=[TxtTotalRegD_Montant_mTVA]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =2655
                    Top =645
                    Width =816
                    Height =225
                    ColumnOrder =0
                    TabIndex =14
                    Name ="TxtTotalRegD_Montant_mDirectTTCMonnaie"
                    ControlSource ="=[TxtTotalRegD_Montant_mDirectTTC]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =2655
                    Top =975
                    Width =816
                    Height =225
                    ColumnOrder =15
                    TabIndex =15
                    Name ="TxtTotalRegD_Montant_mTTCMonnaie"
                    ControlSource ="=[TxtTotalRegD_Montant_mTTC]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            Name ="PiedGroupe2"
        End
        Begin PageFooter
            Height =567
            Name ="PiedPage"
        End
    End
End
CodeBehindForm
' See "EtaCourrierReglementPourcentage.cls"
