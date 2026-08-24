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
    ItemSuffix =126
    Left =450
    Top =1680
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x23ff438bf7e8e240
    End
    RecordSource ="SELECT DISTINCTROW TMP_IRRECOUVRABLE.TIrr_Déb_Code, DEBITEUR.Déb_Qualité, DEBITE"
        "UR.Déb_Nom, DEBITEUR.Déb_Nom_Banque, DEBITEUR.Déb_Date_Chèque, DEBITEUR.Déb_Num_"
        "Chèque, DEBITEUR.Déb_RIB_Banque, DEBITEUR.Déb_RIB_Guichet, DEBITEUR.Déb_RIB_Comp"
        "te, DEBITEUR.Déb_RIB_Clé, DEBITEUR.Déb_RIB_Dom, DEBITEUR.Déb_Date_Cré, DEBITEUR."
        "Déb_Mont_Prin_m1, DEBITEUR.Déb_Mont_Prin_m2, DEBITEUR.Déb_MonnaieEnCours, DEBITE"
        "UR.Déb_Réf1_Client, DEBITEUR.Déb_Réf2_Client, ADRESSE_DEBITEUR.AD_Adr1, ADRESSE_"
        "DEBITEUR.AD_Adr2, ADRESSE_DEBITEUR.AD_Adr3, ADRESSE_DEBITEUR.AD_Adr4, ADRESSE_DE"
        "BITEUR.AD_Ville, ADRESSE_DEBITEUR.AD_CodePostal, ADRESSE_DEBITEUR.AD_Pays FROM ("
        "DEBITEUR INNER JOIN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Code=ADRESSE_DEBITEUR.AD_Co"
        "de_Débiteur) INNER JOIN TMP_IRRECOUVRABLE ON DEBITEUR.Déb_Code=TMP_IRRECOUVRABLE"
        ".TIrr_Déb_Code WHERE (((ADRESSE_DEBITEUR.AD_TypeAdresse_ID)=DLookUp(\"Adresse_Co"
        "urrier\",\"PARAMETRE_GREF\")) AND ((TMP_IRRECOUVRABLE.TIrr_Editer)=True)); "
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
            ControlSource ="Déb_Nom"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            Name ="EntêteGroupe"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =13889
            Name ="Détail1"
            Begin
                Begin TextBox
                    TextFontFamily =49
                    Left =470
                    Top =9244
                    Width =8400
                    FontSize =9
                    Name ="TxtDéb_Nom"
                    ControlSource ="=Trim([Déb_Qualité]) & \" \" & [DEBITEUR.Déb_Nom]"
                    FontName ="Courier New"
                End
                Begin Line
                    Left =241
                    Top =1485
                    Width =8844
                    Name ="Trait52"
                End
                Begin Line
                    BorderLineStyle =1
                    Left =241
                    Top =2166
                    Width =8844
                    Name ="Trait53"
                End
                Begin Line
                    Left =241
                    Top =1485
                    Width =0
                    Height =3300
                    Name ="Trait54"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4988
                    Top =1485
                    Width =0
                    Height =3300
                    Name ="Trait55"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =345
                    Top =1559
                    Width =3675
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Étiquette56"
                    Caption ="IDENTITE COMPLETE DU PLAIGNANT"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =5235
                    Top =1559
                    Width =2355
                    Height =480
                    FontSize =10
                    FontWeight =700
                    Name ="Étiquette57"
                    Caption ="ENSEIGNE DU MAGASIN\015\012(Nom de la Société)"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =465
                    Top =4964
                    Width =4035
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Étiquette58"
                    Caption ="... DEPOSE PLAINTE CONTRE INCONNU"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =468
                    Top =2392
                    Width =615
                    Height =240
                    FontSize =9
                    Name ="Étiquette59"
                    Caption ="NOM :"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =465
                    Top =2654
                    Width =945
                    Height =240
                    FontSize =9
                    Name ="Étiquette60"
                    Caption ="Prénom :"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =465
                    Top =2909
                    Width =2130
                    Height =240
                    FontSize =9
                    Name ="Étiquette61"
                    Caption ="Date de naissance :"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =465
                    Top =3164
                    Width =2130
                    Height =240
                    FontSize =9
                    Name ="Étiquette62"
                    Caption ="Lieu de naissance :"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =465
                    Top =3681
                    Width =1365
                    Height =240
                    FontSize =9
                    Name ="Étiquette64"
                    Caption ="Profession :"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =465
                    Top =3929
                    Width =1155
                    Height =240
                    FontSize =9
                    Name ="Étiquette65"
                    Caption ="Domicile :"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =465
                    Top =4184
                    Width =615
                    Height =240
                    FontSize =9
                    Name ="Étiquette66"
                    Caption ="Tél :"
                    FontName ="Courier New"
                End
                Begin Line
                    BorderLineStyle =1
                    Left =240
                    Top =4784
                    Width =8844
                    Name ="Trait67"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =9085
                    Top =1485
                    Width =0
                    Height =3300
                    Name ="Trait68"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =5230
                    Top =4207
                    Width =2130
                    Height =240
                    FontSize =9
                    Name ="Étiquette69"
                    Caption ="(timbre du magasin)"
                    FontName ="Courier New"
                End
                Begin Line
                    BorderLineStyle =1
                    Left =241
                    Top =5385
                    Width =8844
                    Name ="Trait70"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =355
                    Top =5499
                    Width =1800
                    Height =240
                    FontSize =9
                    Name ="Étiquette71"
                    Caption ="DATE DES FAITS :"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =358
                    Top =5799
                    Width =1905
                    Height =240
                    FontSize =9
                    Name ="Étiquette72"
                    Caption ="HEURE DES FAITS :"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =355
                    Top =6099
                    Width =4065
                    Height =240
                    FontSize =9
                    Name ="Étiquette73"
                    Caption ="LIEU (rayon dans une grande surface):"
                    FontName ="Courier New"
                End
                Begin Line
                    Left =240
                    Top =5384
                    Width =0
                    Height =1740
                    Name ="Trait74"
                End
                Begin Line
                    Left =4995
                    Top =5384
                    Width =0
                    Height =1740
                    Name ="Trait75"
                End
                Begin Line
                    Left =9090
                    Top =5354
                    Width =0
                    Height =1740
                    Name ="Trait78"
                End
                Begin Line
                    BorderLineStyle =1
                    Left =240
                    Top =7109
                    Width =8844
                    Name ="Trait79"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =5117
                    Top =5499
                    Width =3960
                    Height =240
                    FontSize =9
                    Name ="Étiquette80"
                    Caption ="Marchandises achetées (description):"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =49
                    Left =2169
                    Top =5499
                    FontSize =9
                    TabIndex =1
                    Name ="TxtDéb_Date_Chèque"
                    ControlSource ="Déb_Date_Chèque"
                    Format ="Short Date"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =5117
                    Top =6292
                    Width =1260
                    Height =240
                    FontSize =9
                    Name ="Étiquette82"
                    Caption ="Préjudice :"
                    FontName ="Courier New"
                End
                Begin Line
                    Left =240
                    Top =7335
                    Width =8844
                    Name ="Trait84"
                End
                Begin Line
                    BorderLineStyle =1
                    Left =240
                    Top =7776
                    Width =8844
                    Name ="Trait85"
                End
                Begin Line
                    Left =240
                    Top =7335
                    Width =0
                    Height =4185
                    Name ="Trait86"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4995
                    Top =7769
                    Width =0
                    Height =900
                    Name ="Trait87"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =345
                    Top =7409
                    Width =5115
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="Étiquette88"
                    Caption ="IDENTIFICATION DU CHEQUE ANNEXE AU PRESENT"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =465
                    Top =7889
                    Width =3645
                    Height =240
                    FontSize =9
                    Name ="Étiquette90"
                    Caption ="NOM DE L'ETABLISSEMENT BANCAIRE :"
                    FontName ="Courier New"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =9084
                    Top =7335
                    Width =0
                    Height =4185
                    Name ="Trait91"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =5115
                    Top =7874
                    Width =4065
                    Height =240
                    FontSize =9
                    Name ="Étiquette92"
                    Caption ="ADRESSE DE L'ETABLISSEMENT BANCAIRE :"
                    FontName ="Courier New"
                End
                Begin Line
                    BorderLineStyle =1
                    Left =241
                    Top =8787
                    Width =8844
                    Name ="Trait93"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =465
                    Top =8894
                    Width =5370
                    Height =240
                    FontSize =9
                    Name ="Étiquette94"
                    Caption ="NOM, PRENOM ET ADRESSE DU TITULAIRE DU CHEQUIER :"
                    FontName ="Courier New"
                End
                Begin Line
                    BorderLineStyle =1
                    Left =241
                    Top =10034
                    Width =8844
                    Name ="Trait95"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =468
                    Top =8220
                    Width =4416
                    Height =450
                    FontSize =9
                    TabIndex =2
                    Name ="TxtDéb_Nom_Banque"
                    ControlSource ="Déb_Nom_Banque"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5120
                    Top =8220
                    Width =3846
                    Height =450
                    FontSize =9
                    TabIndex =3
                    Name ="TxtDéb_RIB_Dom"
                    ControlSource ="Déb_RIB_Dom"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =465
                    Top =9479
                    Width =8505
                    FontSize =9
                    TabIndex =4
                    Name ="TxtAD_Adr"
                    ControlSource ="=Trim([AD_Adr1]) & \" \" & Trim([AD_Adr2]) & \" \" & Trim([AD_Adr3]) & \" \" & T"
                        "rim([AD_Adr4])"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =465
                    Top =9719
                    Width =8505
                    FontSize =9
                    TabIndex =5
                    Name ="TxtAD_CodePostal"
                    ControlSource ="=Trim([AD_CodePostal]) & \" \" & Trim([AD_Ville]) & \"  \" & Trim([AD_Pays])"
                    FontName ="Courier New"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =465
                    Top =10469
                    Width =4425
                    FontSize =9
                    TabIndex =6
                    Name ="TxtDéb_RIB"
                    ControlSource ="=Trim([Déb_RIB_Compte]) & \" \" & Trim([Déb_RIB_Clé])"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =465
                    Top =10154
                    Width =1590
                    Height =240
                    FontSize =9
                    Name ="Étiquette108"
                    Caption ="N° DU COMPTE :"
                    FontName ="Courier New"
                End
                Begin Line
                    BorderLineStyle =1
                    Left =240
                    Top =10829
                    Width =8844
                    Name ="Trait109"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4995
                    Top =10049
                    Width =0
                    Height =1470
                    Name ="Trait115"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =5115
                    Top =10469
                    Width =3750
                    FontSize =9
                    TabIndex =7
                    Name ="TxtDéb_Num_Chèque"
                    ControlSource ="Déb_Num_Chèque"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =5115
                    Top =10154
                    Width =1590
                    Height =240
                    FontSize =9
                    Name ="Étiquette117"
                    Caption ="N° DU CHEQUE :"
                    FontName ="Courier New"
                End
                Begin Line
                    BorderLineStyle =1
                    Left =240
                    Top =11519
                    Width =8844
                    Name ="Trait118"
                End
                Begin TextBox
                    TextFontFamily =49
                    Left =465
                    Top =11204
                    Width =4425
                    FontSize =9
                    TabIndex =8
                    Name ="TxtDéb_Mont_Prin_m"
                    ControlSource ="=IIf([Déb_MonnaieEnCours]=1,Format([Déb_Mont_Prin_m1],\"#,##0.00\"\" F\"\"\"),Fo"
                        "rmat([Déb_Mont_Prin_m2],\"#,##0.00 €\"))"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =465
                    Top =10889
                    Width =2130
                    Height =240
                    FontSize =9
                    Name ="Étiquette120"
                    Caption ="MONTANT DU CHEQUE :"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =5115
                    Top =10889
                    Width =2775
                    Height =240
                    FontSize =9
                    Name ="Étiquette121"
                    Caption ="MENTION FIGURANT AU DOS :"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =468
                    Top =11962
                    Width =8400
                    Height =555
                    FontSize =9
                    Name ="Étiquette123"
                    Caption ="SIGNALEMENT EVENTUEL DE LA PERSONNE QUI A REMIS CE CHEQUE OU OBSERVATIONS PARTIC"
                        "ULIERES (complicité - identification d'une voiture - etc...)"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =6818
                    Top =13096
                    Width =285
                    Height =240
                    FontSize =9
                    Name ="Étiquette124"
                    Caption ="Le"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =49
                    Left =6825
                    Top =13454
                    Width =1365
                    Height =435
                    FontSize =9
                    Name ="Étiquette125"
                    Caption ="Le Plaignant\015\012(Signature)"
                    FontName ="Courier New"
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    TextFontFamily =49
                    Left =225
                    Width =8160
                    Height =450
                    FontSize =18
                    FontWeight =700
                    Name ="Étiquette49"
                    Caption ="LETTRE PLAINTE"
                    FontName ="Courier New"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =49
                    Left =247
                    Top =570
                    Width =8145
                    Height =555
                    FontSize =12
                    Name ="Étiquette50"
                    Caption ="A\015\012MONSIEUR LE PROCUREUR DE LA REPUBLIQUE"
                    FontName ="Courier New"
                End
            End
        End
    End
End
CodeBehindForm
' See "EtaLettrePlainte.cls"
