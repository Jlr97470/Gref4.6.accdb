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
    Width =9615
    ItemSuffix =14
    Left =4395
    Top =2520
    OnNoData ="[Event Procedure]"
    Filter ="TRel_Editer = True AND TRel_Niv_relance In (3,5)"
    RecSrcDt = Begin
        0x37f594cfb7eae240
    End
    RecordSource ="SELECT DEBITEUR.Déb_Code, DEBITEUR.Déb_Qualité, DEBITEUR.Déb_Nom, ADRESSE_DEBITE"
        "UR.AD_Adr1, ADRESSE_DEBITEUR.AD_Adr2, ADRESSE_DEBITEUR.AD_Adr3, ADRESSE_DEBITEUR"
        ".AD_Adr4, ADRESSE_DEBITEUR.AD_CodePostal, ADRESSE_DEBITEUR.AD_Ville, ADRESSE_DEB"
        "ITEUR.AD_Pays, CLIENT.Cli_Nom, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb_Frais_m1,"
        " DEBITEUR.Déb_Mont_Prin_m2, DEBITEUR.Déb_Frais_m2, DEBITEUR.Déb_MonnaieEnCours, "
        "TMP_RELANCE.TRel_Niv_relance, TMP_RELANCE.TRel_Editer FROM (CLIENT INNER JOIN (D"
        "EBITEUR INNER JOIN TMP_RELANCE ON DEBITEUR.Déb_Code=TMP_RELANCE.TRel_Déb_Code) O"
        "N CLIENT.Cli_Code=DEBITEUR.Déb_Code_Client) INNER JOIN ADRESSE_DEBITEUR ON DEBIT"
        "EUR.Déb_Code=ADRESSE_DEBITEUR.AD_Code_Débiteur WHERE (((TMP_RELANCE.TRel_Niv_rel"
        "ance)=3 Or (TMP_RELANCE.TRel_Niv_relance)=5) AND ((TMP_RELANCE.TRel_Editer)=True"
        ") AND ((ADRESSE_DEBITEUR.AD_TypeAdresse_ID)=DLookUp(\"Adresse_Courrier\",\"PARAM"
        "ETRE_GREF\"))); "
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
            Height =9255
            Name ="Détail"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =450
                    Top =3870
                    Width =3885
                    Height =570
                    FontSize =10
                    TabIndex =4
                    Name ="TxtRelance5"
                    ControlSource ="=IIf([TRel_Niv_relance]=3,\"\",\"Dernier avis avant déplacement d'un agent de re"
                        "couvrement.\")"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =450
                    Top =3390
                    Width =1755
                    Height =270
                    FontSize =12
                    FontWeight =700
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =5385
                    Top =4635
                    Width =1935
                    Height =315
                    FontSize =10
                    TabIndex =1
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =450
                    Top =3120
                    Width =3735
                    Height =270
                    FontSize =10
                    TabIndex =2
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =3675
                    Top =6585
                    Width =1950
                    Height =270
                    FontSize =11
                    TabIndex =3
                    Name ="TxtMontant"
                    ControlSource ="=Round([Déb_Mont_Prin_m1],2)-Round(nz(DSum(\"Enc_Prin_m1\",\"ENCAISSEMENT\",\"En"
                        "c_Code_Débiteur='\" & [Déb_Code] & \"'\"),0),2)+Round([Déb_Frais_m1],2)-Round(nz"
                        "(DSum(\"Enc_Frais_m1\",\"ENCAISSEMENT\",\"Enc_Code_Débiteur='\" & [Déb_Code] & \""
                        "'\"),0),2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Top =1133
                    Width =2598
                    Height =1256
                    FontSize =10
                    Name ="EtiEnteteCaen"
                    Caption ="CABINET GREF / N.R.S.\015\0129, rue Pasteur\015\01214000 CAEN\015\012Tel  02.31."
                        "85.52.27\015\012Fax  02.31.86.07.50"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5400
                    Top =1140
                    Width =4125
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4845
                    Top =1395
                    Width =4455
                    Height =255
                    FontSize =10
                    TabIndex =6
                    Name ="TxtAD_Adr1"
                    ControlSource ="AD_Adr1"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4845
                    Top =1140
                    Width =540
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    Name ="TxtDéb_Qualité"
                    ControlSource ="Déb_Qualité"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4845
                    Top =1650
                    Width =4455
                    Height =255
                    FontSize =10
                    TabIndex =8
                    Name ="TxtAD_Adr2"
                    ControlSource ="AD_Adr2"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4845
                    Top =2420
                    Width =960
                    Height =255
                    FontSize =10
                    TabIndex =9
                    Name ="TxtAD_CodePostal"
                    ControlSource ="AD_CodePostal"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5827
                    Top =2415
                    Width =3480
                    Height =255
                    FontSize =10
                    TabIndex =10
                    Name ="TxtAD_Ville"
                    ControlSource ="AD_Ville"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4845
                    Top =1905
                    Width =4455
                    Height =255
                    FontSize =10
                    TabIndex =11
                    Name ="TxtAD_Adr3"
                    ControlSource ="AD_Adr3"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4845
                    Top =2160
                    Width =4455
                    Height =255
                    FontSize =10
                    TabIndex =12
                    Name ="TxtAD_Adr4"
                    ControlSource ="AD_Adr4"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =4845
                    Top =2670
                    Width =3165
                    Height =255
                    FontSize =10
                    TabIndex =13
                    Name ="TxtAD_Pays"
                    ControlSource ="AD_Pays"
                    FontName ="Courier New"
                End
            End
        End
    End
End
CodeBehindForm
' See "EtaAvisAssignation.cls"
