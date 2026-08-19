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
    Width =9600
    ItemSuffix =16
    Left =1440
    Top =1770
    OnNoData ="[Event Procedure]"
    Filter ="TRel_Editer = True AND TRel_Niv_relance=1"
    RecSrcDt = Begin
        0x1c282bb8b7eae240
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
        "P_RELANCE.TRel_Niv_relance)=1) AND ((TMP_RELANCE.TRel_Editer)=True) AND ((ADRESS"
        "E_DEBITEUR.AD_TypeAdresse_ID)=DLookUp(\"Adresse_Courrier\",\"PARAMETRE_GREF\")))"
        "; "
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
            Height =10890
            Name ="Détail1"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =1080
                    Top =2430
                    Width =1680
                    Height =315
                    FontSize =12
                    FontWeight =700
                    Name ="Déb_Code"
                    ControlSource ="Déb_Code"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =6240
                    Top =4980
                    Width =1935
                    Height =315
                    FontSize =11
                    TabIndex =1
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="dd/mm/yyyy"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =49
                    Left =6135
                    Top =7395
                    Width =1920
                    Height =315
                    FontSize =11
                    TabIndex =2
                    Name ="TxtDéb_Mont_Prin_m"
                    ControlSource ="=Round([Déb_Mont_Prin_m1],2)-Round(nz(DSum(\"Enc_Prin_m1\",\"ENCAISSEMENT\",\"En"
                        "c_Code_Débiteur='\" & [Déb_Code] & \"'\"),0),2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =49
                    Left =6135
                    Top =8115
                    Width =1920
                    Height =315
                    FontSize =11
                    TabIndex =3
                    Name ="TxtDéb_Frais_m"
                    ControlSource ="=Round([Déb_Frais_m1],2)-Round(nz(DSum(\"Enc_Frais_m1\",\"ENCAISSEMENT\",\"Enc_C"
                        "ode_Débiteur='\" & [Déb_Code] & \"'\"),0),2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =49
                    Left =6135
                    Top =8850
                    Width =1920
                    Height =315
                    FontSize =11
                    TabIndex =4
                    Name ="TxtTotal"
                    ControlSource ="=[TxtDéb_Mont_Prin_m]+[TxtDéb_Frais_m]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =3195
                    Top =6180
                    Width =5475
                    Height =495
                    FontSize =11
                    TabIndex =5
                    Name ="TxtDéb_Typ_Cré"
                    ControlSource ="Déb_Typ_Cré"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =3090
                    Top =5940
                    Width =4845
                    Height =255
                    FontSize =11
                    TabIndex =6
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =1245
                    Top =450
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
                    Left =8085
                    Top =7395
                    Width =1500
                    Height =315
                    FontSize =11
                    TabIndex =7
                    Name ="TxtDéb_Mont_Prin_mMonnaie"
                    ControlSource ="=[TxtDéb_Mont_Prin_m]*DonnerTaucConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =49
                    Left =8085
                    Top =8115
                    Width =1500
                    Height =315
                    FontSize =11
                    TabIndex =8
                    Name ="TxtDéb_Frais_mMonnaie"
                    ControlSource ="=[TxtDéb_Frais_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =49
                    Left =8085
                    Top =8850
                    Width =1500
                    Height =315
                    FontSize =11
                    TabIndex =9
                    Name ="TxtTotalMonnaie"
                    ControlSource ="=[TxtTotal]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5400
                    Top =1365
                    Width =4140
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4860
                    Top =1620
                    Width =4455
                    Height =255
                    FontSize =10
                    TabIndex =11
                    Name ="TxtAD_Adr1"
                    ControlSource ="AD_Adr1"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4860
                    Top =1365
                    Width =540
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    Name ="TxtDéb_Qualité"
                    ControlSource ="Déb_Qualité"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4860
                    Top =1875
                    Width =4455
                    Height =255
                    FontSize =10
                    TabIndex =13
                    Name ="TxtAD_Adr2"
                    ControlSource ="AD_Adr2"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4860
                    Top =2645
                    Width =960
                    Height =255
                    FontSize =10
                    TabIndex =14
                    Name ="TxtAD_CodePostal"
                    ControlSource ="AD_CodePostal"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5827
                    Top =2640
                    Width =3495
                    Height =255
                    FontSize =10
                    TabIndex =15
                    Name ="TxtAD_Ville"
                    ControlSource ="AD_Ville"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4860
                    Top =2130
                    Width =4455
                    Height =255
                    FontSize =10
                    TabIndex =16
                    Name ="TxtAD_Adr3"
                    ControlSource ="AD_Adr3"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4860
                    Top =2385
                    Width =4455
                    Height =255
                    FontSize =10
                    TabIndex =17
                    Name ="TxtAD_Adr4"
                    ControlSource ="AD_Adr4"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4860
                    Top =2895
                    Width =3165
                    Height =255
                    FontSize =10
                    TabIndex =18
                    Name ="TxtAD_Pays"
                    ControlSource ="AD_Pays"
                    FontName ="Courier New"
                End
            End
        End
    End
End
CodeBehindForm
' See "EtaAvisMiseRecouvrement.cls"
