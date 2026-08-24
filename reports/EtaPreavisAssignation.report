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
    Width =9612
    ItemSuffix =9
    Left =2100
    Top =1875
    OnNoData ="[Event Procedure]"
    Filter ="TRel_Editer = True AND TRel_Niv_relance In (2,4,6)"
    RecSrcDt = Begin
        0xb6e1d8cab7eae240
    End
    RecordSource ="SELECT DEBITEUR.Déb_Code, DEBITEUR.Déb_Qualité, DEBITEUR.Déb_Nom, DEBITEUR.Déb_T"
        "yp_Cré, ADRESSE_DEBITEUR.AD_Adr1, ADRESSE_DEBITEUR.AD_Adr2, ADRESSE_DEBITEUR.AD_"
        "Adr3, ADRESSE_DEBITEUR.AD_Adr4, ADRESSE_DEBITEUR.AD_CodePostal, ADRESSE_DEBITEUR"
        ".AD_Ville, ADRESSE_DEBITEUR.AD_Pays, CLIENT.Cli_Nom, DEBITEUR.Déb_Mont_Prin_m1, "
        "DEBITEUR.Déb_Frais_m1, DEBITEUR.Déb_Mont_Prin_m2, DEBITEUR.Déb_Frais_m2, DEBITEU"
        "R.Déb_MonnaieEnCours, TMP_RELANCE.TRel_Niv_relance, TMP_RELANCE.TRel_Editer FROM"
        " (CLIENT INNER JOIN (DEBITEUR INNER JOIN TMP_RELANCE ON DEBITEUR.Déb_Code=TMP_RE"
        "LANCE.TRel_Déb_Code) ON CLIENT.Cli_Code=DEBITEUR.Déb_Code_Client) INNER JOIN ADR"
        "ESSE_DEBITEUR ON DEBITEUR.Déb_Code=ADRESSE_DEBITEUR.AD_Code_Débiteur WHERE (((TM"
        "P_RELANCE.TRel_Niv_relance)=2 Or (TMP_RELANCE.TRel_Niv_relance)=4 Or (TMP_RELANC"
        "E.TRel_Niv_relance)=6) AND ((TMP_RELANCE.TRel_Editer)=True) AND ((ADRESSE_DEBITE"
        "UR.AD_TypeAdresse_ID)=DLookUp(\"Adresse_Courrier\",\"PARAMETRE_GREF\"))); "
    OnOpen ="[Event Procedure]"
    Begin
        Begin Label
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin BreakLevel
            ControlSource ="TRel_Niv_relance"
        End
        Begin BreakLevel
            ControlSource ="Déb_Code"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =9600
            Name ="Détail"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =3450
                    Top =3420
                    Width =1575
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =5880
                    Top =8205
                    Width =1935
                    Height =330
                    FontSize =11
                    TabIndex =1
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =5085
                    Top =3420
                    Width =3855
                    Height =330
                    FontSize =11
                    TabIndex =2
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =49
                    Left =945
                    Top =4905
                    Width =1680
                    Height =330
                    FontSize =11
                    TabIndex =3
                    Name ="TxtDéb_Mont_Prin_m"
                    ControlSource ="=Round([Déb_Mont_Prin_m1],2)-Round(nz(DSum(\"Enc_Prin_m1\",\"ENCAISSEMENT\",\"En"
                        "c_Code_Débiteur='\" & [Déb_Code] & \"'\"),0),2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Courier New"
                End
                Begin Label
                    TextFontFamily =49
                    Left =1560
                    Top =3420
                    Width =1875
                    Height =255
                    FontSize =10
                    Name ="EtiReferences"
                    Caption ="Vos références:"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =49
                    Left =945
                    Top =5910
                    Width =1680
                    Height =330
                    FontSize =11
                    TabIndex =4
                    Name ="TxtDéb_Frais_m"
                    ControlSource ="=Round([Déb_Frais_m1],2)-Round(nz(DSum(\"Enc_Prin_m1\",\"ENCAISSEMENT\",\"Enc_Co"
                        "de_Débiteur='\" & [Déb_Code] & \"'\"),0),2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =49
                    Left =945
                    Top =6330
                    Width =1680
                    Height =330
                    FontSize =11
                    TabIndex =5
                    Name ="TxtTotal"
                    ControlSource ="=[TxtDéb_Mont_Prin_m]+[TxtDéb_Frais_m]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =3120
                    Top =4905
                    Width =5205
                    Height =330
                    FontSize =11
                    TabIndex =6
                    Name ="Déb_Typ_Cré"
                    ControlSource ="Déb_Typ_Cré"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =1560
                    Top =3780
                    Width =4170
                    Height =585
                    FontSize =11
                    TabIndex =7
                    Name ="TxtRelance"
                    ControlSource ="=IIf([TRel_Niv_relance]=\"2\",\"Seconde Relance\",IIf([TRel_Niv_relance]=\"4\",\""
                        "\",IIf([TRel_Niv_relance]=\"6\",\"Ultime Relance\",\"\")))"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5872
                    Top =1455
                    Width =3480
                    Height =255
                    FontSize =10
                    TabIndex =8
                    Name ="Texte14"
                    ControlSource ="AD_Ville"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =6060
                    Top =750
                    Width =3540
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5520
                    Top =1005
                    Width =3855
                    Height =255
                    FontSize =10
                    TabIndex =10
                    Name ="TxtAD_Adr1"
                    ControlSource ="AD_Adr1"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5520
                    Top =750
                    Width =540
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    Name ="TxtDéb_Qualité"
                    ControlSource ="Déb_Qualité"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5520
                    Top =1260
                    Width =3855
                    Height =255
                    FontSize =10
                    TabIndex =12
                    Name ="TxtAD_Adr2"
                    ControlSource ="AD_Adr2"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5520
                    Top =2030
                    Width =960
                    Height =255
                    FontSize =10
                    TabIndex =13
                    Name ="TxtAD_CodePostal"
                    ControlSource ="AD_CodePostal"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =6502
                    Top =2025
                    Width =2880
                    Height =255
                    FontSize =10
                    TabIndex =14
                    Name ="TxtAD_Ville"
                    ControlSource ="AD_Ville"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5520
                    Top =1515
                    Width =3855
                    Height =255
                    FontSize =10
                    TabIndex =15
                    Name ="TxtAD_Adr3"
                    ControlSource ="AD_Adr3"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5520
                    Top =1770
                    Width =3855
                    Height =255
                    FontSize =10
                    TabIndex =16
                    Name ="TxtAD_Adr4"
                    ControlSource ="AD_Adr4"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5520
                    Top =2280
                    Width =3090
                    Height =255
                    FontSize =10
                    TabIndex =17
                    Name ="TxtAD_Pays"
                    ControlSource ="AD_Pays"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =660
                    Top =240
                    Width =2598
                    Height =1226
                    FontSize =10
                    Name ="EtiEnteteCaen"
                    Caption ="CABINET GREF / N.R.S.\015\0129, rue Pasteur\015\01214000 CAEN\015\012Tel  02.31."
                        "85.52.27\015\012Fax  02.31.86.07.50"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =49
                    Left =7620
                    Top =4965
                    Width =1935
                    Height =225
                    FontSize =11
                    TabIndex =18
                    Name ="TxtDéb_Mont_Prin_mMonnaie"
                    ControlSource ="=[TxtDéb_Mont_Prin_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =49
                    Left =7620
                    Top =5970
                    Width =1935
                    Height =225
                    FontSize =11
                    TabIndex =19
                    Name ="TxtDéb_Frais_mMonnaie"
                    ControlSource ="=[TxtDéb_Frais_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =49
                    Left =7620
                    Top =6390
                    Width =1935
                    Height =225
                    FontSize =11
                    TabIndex =20
                    Name ="TxtTotalMonnaie"
                    ControlSource ="=[TxtTotal]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Courier New"
                End
            End
        End
    End
End
CodeBehindForm
' See "EtaPreavisAssignation.cls"
