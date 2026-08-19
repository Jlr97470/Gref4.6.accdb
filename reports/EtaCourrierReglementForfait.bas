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
    ItemSuffix =54
    Left =-27106
    Top =2700
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x7c67031a56efe240
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
        "T.Reg_Id=REGLEMENT_DETAIL.RegD_Reg_Id WHERE (((REGLEMENT.Reg_Cli_Pourcentage)=0)"
        " AND ((REGLEMENT_DETAIL.RegD_ETy_Mode)=1) AND ((ADRESSE_CLIENT.AC_TypeAdresse_ID"
        ")=DLookUp(\"Adresse_Courrier\",\"PARAMETRE_GREF\")) AND ((REGLEMENT.Reg_Editer)="
        "True)); "
    OnOpen ="[Event Procedure]"
    FilterOnLoad =0
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
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Reg_Date"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="Cli_Code"
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
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =3000
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
                    Name ="EtiRef"
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
                    Caption ="Paris le"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =3510
                    Top =2445
                    Width =1680
                    Height =330
                    FontSize =12
                    Name ="TxtReg_Cli_Code"
                    ControlSource ="Reg_Cli_Code"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =855
                    Top =2445
                    Width =1950
                    Height =330
                    FontSize =12
                    TabIndex =1
                    Name ="Date du jour"
                    ControlSource ="Reg_Date"
                    Format ="Short Date"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Date_du_jour"

                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =8130
                    Top =2490
                    Width =705
                    Height =510
                    FontSize =10
                    Name ="EtiResteAVerser"
                    Caption ="Reste à verser"
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =6495
                    Top =2715
                    Width =1140
                    Height =285
                    FontSize =10
                    Name ="EtiReglement"
                    Caption ="Règlement"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4245
                    Top =15
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =2
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4245
                    Top =285
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =3
                    Name ="TxtAC_Adr1"
                    ControlSource ="AC_Adr1"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4245
                    Top =555
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =4
                    Name ="TxtAC_Adr2"
                    ControlSource ="AC_Adr2"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4245
                    Top =1368
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
                    Top =1365
                    Width =3840
                    Height =270
                    FontSize =10
                    TabIndex =6
                    Name ="TxtAC_Ville"
                    ControlSource ="AC_Ville"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4245
                    Top =825
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =7
                    Name ="TxtAC_Adr3"
                    ControlSource ="AC_Adr3"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4245
                    Top =1095
                    Width =4800
                    Height =270
                    FontSize =10
                    TabIndex =8
                    Name ="TxtAC_Adr4"
                    ControlSource ="AC_Adr4"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4245
                    Top =1635
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
            Height =690
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Top =15
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
                    Top =15
                    Width =2871
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
                    Left =5145
                    Top =15
                    Width =1221
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
                    Left =6360
                    Top =15
                    Width =1356
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
                    Left =4005
                    Top =15
                    Width =1146
                    Height =255
                    FontSize =10
                    TabIndex =4
                    Name ="TxtRegD_Type"
                    ControlSource ="RegD_Type"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =510
                    Top =270
                    Width =2595
                    Height =255
                    TabIndex =5
                    Name ="TxtDéb_Réf1_Client"
                    ControlSource ="Déb_Réf1_Client"

                End
                Begin Label
                    TextFontFamily =34
                    Left =15
                    Top =270
                    Width =495
                    Height =255
                    Name ="EtiDéb_Réf1_Client"
                    Caption ="Réf1 :"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =3615
                    Top =270
                    Width =1695
                    Height =255
                    TabIndex =6
                    Name ="TxtDéb_Réf2_Client"
                    ControlSource ="Déb_Réf2_Client"

                End
                Begin Label
                    TextFontFamily =34
                    Left =3105
                    Top =270
                    Width =510
                    Height =255
                    Name ="EtiDéb_Réf2_Client"
                    Caption ="Réf2 :"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =18
                    Left =7710
                    Top =15
                    Width =1356
                    Height =255
                    FontSize =10
                    TabIndex =7
                    Name ="TxtResteAVerser"
                    ControlSource ="=Round([Déb_Mont_Prin_m1],2)-Round(Nz(DSum(\"Enc_Prin_m1\",\"ENCAISSEMENT\",\"En"
                        "c_RegD_Id<=\" & [RegD_Id] & \" AND Enc_Code_Débiteur='\" & [RegD_Déb_Code] & \"'"
                        "\"),0),2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =18
                    Left =6360
                    Top =270
                    Width =1356
                    Height =255
                    FontSize =10
                    TabIndex =8
                    Name ="TxtRegD_Montant_mMonnaie"
                    ControlSource ="=[TxtRegD_Montant_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =18
                    Left =7710
                    Top =270
                    Width =1356
                    Height =255
                    FontSize =10
                    TabIndex =9
                    Name ="TxtResteAVerserMonnaie"
                    ControlSource ="=[TxtResteAVerser]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5310
                    Top =270
                    Width =1056
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
            Height =1650
            BreakLevel =1
            Name ="PiedGroupe"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =1335
                    Top =750
                    Width =1356
                    Height =330
                    FontSize =12
                    Name ="TxtMontantTTC"
                    ControlSource ="=[TxtTotalRegD_Montant_m]*[Reg_Cli_Pourcentage]/100*(1+[Reg_TVA])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =45
                            Top =750
                            Width =1245
                            Height =330
                            FontSize =12
                            Name ="EtiMontantTTC"
                            Caption ="Débit TTC :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =7695
                    Top =765
                    Width =1281
                    Height =330
                    FontSize =12
                    TabIndex =1
                    Name ="TxtTotalRegD_Montant_m"
                    ControlSource ="=Nz(Sum(Round([RegD_Montant_m1],2)),0)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =4530
                            Top =765
                            Width =2115
                            Height =330
                            FontSize =12
                            Name ="EtiRegD_Montant_m2"
                            Caption ="Du montant encaissé :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =3780
                    Top =765
                    Width =651
                    Height =330
                    FontSize =12
                    TabIndex =2
                    Name ="TxtReg_Cli_Pourcentage"
                    ControlSource ="=[Reg_Cli_Pourcentage]/100"
                    Format ="Percent"
                    FontName ="Times New Roman"

                End
                Begin Rectangle
                    BackStyle =0
                    Top =570
                    Width =9030
                    Height =1005
                    Name ="RecMontant"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =1335
                    Top =1080
                    Width =1356
                    Height =330
                    FontSize =12
                    TabIndex =3
                    Name ="TxtMontantTVA"
                    ControlSource ="=[TxtMontantTTC]*([Reg_TVA])/(1+[Reg_TVA])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =45
                            Top =1080
                            Width =1245
                            Height =330
                            FontSize =12
                            Name ="EtiMontantTVA"
                            Caption ="Dont TVA :"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =7695
                    Top =1095
                    Width =1281
                    Height =330
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    Name ="TxtMontantReverser"
                    ControlSource ="=[TxtTotalRegD_Montant_m]-[TxtMontantTTC]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =18
                            Left =4530
                            Top =1095
                            Width =2115
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
                    TextAlign =3
                    TextFontFamily =18
                    Left =6705
                    Top =825
                    Width =951
                    Height =225
                    TabIndex =5
                    Name ="TxtTotalRegD_Montant_mMonnaie"
                    ControlSource ="=[TxtTotalRegD_Montant_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =6705
                    Top =1155
                    Width =951
                    Height =225
                    FontWeight =700
                    TabIndex =6
                    Name ="TxtMontantReverserMonnaie"
                    ControlSource ="=[TxtMontantReverser]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =2775
                    Top =780
                    Width =876
                    Height =225
                    TabIndex =7
                    Name ="TxtMontantTTCMonnaie"
                    ControlSource ="=[TxtMontantTTC]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =2775
                    Top =1125
                    Width =876
                    Height =225
                    TabIndex =8
                    Name ="TxtMontantTVAMonnaie"
                    ControlSource ="=[TxtMontantTVA]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            Name ="PiedGroupe1"
        End
        Begin PageFooter
            Height =0
            Name ="PiedPage"
        End
    End
End
CodeBehindForm
' See "EtaCourrierReglementForfait.cls"
