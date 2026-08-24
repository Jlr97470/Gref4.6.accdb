Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DateGrouping =1
    GridX =5
    GridY =5
    Width =9315
    ItemSuffix =49
    Left =210
    Top =2100
    RecSrcDt = Begin
        0x7ba881ee4de9e240
    End
    RecordSource ="SELECT DISTINCTROW CLIENT.Cli_Nom, DEBITEUR.Déb_Code, DEBITEUR.Déb_Qualité, DEBI"
        "TEUR.Déb_Nom, DEBITEUR.Déb_Typ_Cré, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb_Mont"
        "_Prin_m2, ADRESSE_DEBITEUR.AD_Adr1, ADRESSE_DEBITEUR.AD_Adr2, ADRESSE_DEBITEUR.A"
        "D_Adr3, ADRESSE_DEBITEUR.AD_Adr4, ADRESSE_DEBITEUR.AD_Ville, ADRESSE_DEBITEUR.AD"
        "_CodePostal, ADRESSE_DEBITEUR.AD_Pays FROM (CODE_GESTION INNER JOIN (CLIENT INNE"
        "R JOIN DEBITEUR ON CLIENT.Cli_Code=DEBITEUR.Déb_Code_Client) ON CODE_GESTION.CG_"
        "Code=DEBITEUR.Déb_Situation) INNER JOIN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Code=AD"
        "RESSE_DEBITEUR.AD_Code_Débiteur WHERE (((DEBITEUR.Déb_Date_der_rel) Is Null) AND"
        " ((ADRESSE_DEBITEUR.AD_TypeAdresse_ID)=DLookUp(\"Adresse_Courrier\",\"PARAMETRE_"
        "GREF\")) AND ((CODE_GESTION.CG_Code_Action)=5)); "
    Begin
        Begin Label
            TextFontFamily =2
            FontName ="Arial"
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
            ControlSource ="Déb_Code"
        End
        Begin PageHeader
            Height =570
            Name ="EntêtePage0"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =12433
            Name ="EntêteGroupe0"
            Begin
                Begin Label
                    TextFontFamily =18
                    Top =3285
                    Width =1050
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="EtiDéb_Code"
                    Caption ="Référence :"
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextFontFamily =18
                    Left =4590
                    Top =2460
                    Width =810
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="EtiDate"
                    Caption ="Paris, le "
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextFontFamily =18
                    Top =5134
                    Width =9315
                    Height =2445
                    FontSize =10
                    FontWeight =700
                    Name ="EtiCourrierNouveauxChequesVoles1"
                    Caption ="             Nous sommes chargés par les Etablissements\015\012du recouvrement d"
                        "e leurs créances.\015\012\015\012             Bénéficiaire d'un chèque émis sur "
                        "votre compte qui n'a pas été payé pour le motif de perte ou de vol, nous vous de"
                        "mandons, afin de pouvoir classer ce dossier en conformité avec l'article 272-1 d"
                        "u Code général des Impôts, de bien vouloir nous faire parvenir les pièces justif"
                        "icatives suivantes :\015\012\015\012             - Copie de la déclaration de pe"
                        "rte ou du dépôt de plainte pour vol auprès du commissariat de police ou de genda"
                        "rmerie.\015\012             - Copie de la confirmation écrite d'opposition adres"
                        "sée à votre banque."
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextFontFamily =18
                    Left =566
                    Top =4410
                    Width =1890
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="EtiCourrierNouveauxChequesVoles"
                    Caption ="Madame, Monsieur, "
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =5130
                    Top =210
                    Width =3945
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =4590
                    Top =210
                    Width =525
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtDéb_Qualité"
                    ControlSource ="Déb_Qualité"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =1095
                    Top =3291
                    Width =1680
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =5385
                    Top =2460
                    Width =1935
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =4590
                    Top =480
                    Width =4485
                    Height =270
                    FontSize =10
                    TabIndex =4
                    Name ="TxtAD_Adr1"
                    ControlSource ="AD_Adr1"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =4590
                    Top =750
                    Width =4485
                    Height =270
                    FontSize =10
                    TabIndex =5
                    Name ="TxtAD_Adr2"
                    ControlSource ="AD_Adr2"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =4590
                    Top =1560
                    Width =960
                    Height =270
                    FontSize =10
                    TabIndex =6
                    Name ="TxtAD_CodePostal"
                    ControlSource ="AD_CodePostal"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =5550
                    Top =1560
                    Width =3525
                    Height =270
                    FontSize =10
                    TabIndex =7
                    Name ="TxtAD_Ville"
                    ControlSource ="AD_Ville"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =4590
                    Top =1020
                    Width =4485
                    Height =270
                    FontSize =10
                    TabIndex =8
                    Name ="TxtAD_Adr3"
                    ControlSource ="AD_Adr3"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =4590
                    Top =1290
                    Width =4485
                    Height =270
                    FontSize =10
                    TabIndex =9
                    Name ="TxtAD_Adr4"
                    ControlSource ="AD_Adr4"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =4590
                    Top =1830
                    Width =3360
                    Height =270
                    FontSize =10
                    TabIndex =10
                    Name ="TxtAD_Pays"
                    ControlSource ="AD_Pays"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =4545
                    Top =5130
                    Width =4755
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextFontFamily =18
                    Left =2385
                    Top =12103
                    Width =4515
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="EtiCourrierNouveauxChequesVoles5"
                    Caption ="SERVICE ENQUETES ET POURSUITES"
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextFontFamily =18
                    Top =7579
                    Width =9300
                    Height =2925
                    FontSize =10
                    FontWeight =700
                    Name ="EtiCourrierNouveauxChequesVoles2"
                    Caption ="\015\012             L'article 66 alinéa 1 du décret-loi du 30 octobre 1935 modi"
                        "fié par la loi n° 91-1382 du 30 décembre 1991 dispose : \"punie d'un emprisonnem"
                        "ent d'un an à cinq ans et d'une amende de 3 600 Frs à 2 500 000 Frs ou l'une de "
                        "ces deux peines seulement toute personne qui, avec l'intention de porter atteint"
                        "e aux droits d'autrui, aura, après l'émission d'un chèque, retiré par quelque mo"
                        "yen que ce soit, dont le transfert ou le virement de tout ou partie de la provis"
                        "ion\015\012\015\012             A défaut des pièces justificatives demandées, no"
                        "us serons, à regret, contraints de mettre en cause votre responsabilité au regar"
                        "d de l'article 66 précité.\015\012\015\012             Toutefois, si le chèque a"
                        " bien été émis par vous, nous vous demandons et, au besoin, nous vous mettons en"
                        " demeure de nous faire parvenir la somme de :"
                    FontName ="Times New Roman"
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextFontFamily =18
                    Left =3660
                    Top =8779
                    Width =5430
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="EtiCourrierNouveauxChequesVoles3"
                    Caption ="ou  fait dans les mêmes conditions défense au tiré de payer\"."
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =18
                    Left =4980
                    Top =10219
                    Width =1710
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    Name ="TxtDéb_Mont_Prin_m1"
                    ControlSource ="Déb_Mont_Prin_m1"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =7140
                    Top =10219
                    Width =1710
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =13
                    Name ="TxtDéb_Mont_Prin_m2"
                    ControlSource ="Déb_Mont_Prin_m2"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextFontFamily =18
                    Top =10504
                    Width =9300
                    Height =1005
                    FontSize =10
                    FontWeight =700
                    Name ="EtiCourrierNouveauxChequesVoles4"
                    Caption ="\015\012             Vous souhaitant bonne réception de la présente,\015\012\015"
                        "\012             Veuillez agréer l'expression de nos salutations distinguées."
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextFontFamily =18
                    Top =3696
                    Width =675
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="EtiDéb_Typ_Cré"
                    Caption ="Objet :"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =705
                    Top =3690
                    Width =8595
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =14
                    Name ="TxtDéb_Typ_Cré"
                    ControlSource ="Déb_Typ_Cré"
                    FontName ="Times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =6690
                    Top =10215
                    Width =375
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Étiquette48"
                    Caption ="soit"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =0
            Name ="Détail1"
        End
        Begin PageFooter
            Height =0
            Name ="PiedPage2"
        End
    End
End
CodeBehindForm
' See "EtaCourrierNouveauCheque.cls"
