Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DateGrouping =1
    GridX =5
    GridY =5
    Width =9411
    ItemSuffix =70
    Left =180
    Top =2460
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x7514bffa16eae240
    End
    RecordSource ="SELECT DISTINCTROW DEBITEUR.Déb_Code_Enc, DEBITEUR.Déb_Code, DEBITEUR.Déb_Qualit"
        "é, DEBITEUR.Déb_Nom, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb_Mont_Prin_m2, ADRES"
        "SE_DEBITEUR.AD_CodePostal FROM DEBITEUR INNER JOIN ADRESSE_DEBITEUR ON DEBITEUR."
        "Déb_Code=ADRESSE_DEBITEUR.AD_Code_Débiteur WHERE (((DEBITEUR.Déb_Code_Enc) Like "
        "[Veuillez saisir un code encaisseur :])); "
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
            ControlSource ="Déb_Code_Enc"
        End
        Begin BreakLevel
            ControlSource ="Déb_Code"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1937
            Name ="EntêteEtat"
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Top =90
                    Width =5550
                    Height =705
                    FontSize =14
                    FontWeight =700
                    Name ="EtiListeDebiteurEncaisseur"
                    Caption ="Liste des Débiteurs par Encaisseur"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =6066
                    Top =140
                    Width =2955
                    Height =285
                    FontSize =10
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                End
                Begin Line
                    BorderWidth =2
                    Top =30
                    Width =9018
                    Name ="TrtListeDebiteurEncaisseur"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =793
            Name ="EntêteGroupe"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =1700
                    Width =1791
                    Height =345
                    FontSize =12
                    FontWeight =700
                    Name ="TxtDéb_Code_Enc"
                    ControlSource ="Déb_Code_Enc"
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Width =1590
                            Height =345
                            FontSize =12
                            FontWeight =700
                            Name ="EtiDéb_Code_Enc"
                            Caption ="Encaisseur :"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =7513
                    Top =360
                    Width =855
                    Height =270
                    FontSize =10
                    Name ="EtiDéb_Mont_Prin_m"
                    Caption ="Principal"
                End
                Begin Rectangle
                    BackStyle =0
                    Top =726
                    Width =9360
                    Height =0
                    Name ="Boîte46"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =5212
                    Top =371
                    Width =1170
                    Height =270
                    FontSize =10
                    Name ="EtiAD_CodePostal"
                    Caption ="Code Postal"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =340
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =34
                    Left =6570
                    Width =1326
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="TxtDéb_Mont_Prin_m1"
                    ControlSource ="Déb_Mont_Prin_m1"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Top =15
                    Width =1251
                    Height =285
                    FontSize =10
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =1920
                    Top =15
                    Width =3186
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =5137
                    Top =15
                    Width =1341
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="TxtAD_CodePostal"
                    ControlSource ="AD_CodePostal"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =1290
                    Width =531
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="TxtDéb_Qualité"
                    ControlSource ="Déb_Qualité"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =34
                    Left =7935
                    Width =1371
                    Height =285
                    FontSize =10
                    TabIndex =5
                    Name ="TxtDéb_Mont_Prin_m2"
                    ControlSource ="Déb_Mont_Prin_m2"
                    Format ="#,##0.00 €"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =680
            Name ="PiedGroupe"
            Begin
                Begin Rectangle
                    BackStyle =0
                    BorderLineStyle =3
                    Top =15
                    Width =9360
                    Height =0
                    Name ="Boîte56"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =34
                    Left =6576
                    Top =113
                    Width =1326
                    Height =270
                    FontSize =10
                    Name ="TxtTotalDéb_Mont_Prin_m1"
                    ControlSource ="=Sum([Déb_Mont_Prin_m1])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Locked = NotDefault
                    TextFontFamily =34
                    Left =7941
                    Top =113
                    Width =1371
                    Height =270
                    FontSize =10
                    TabIndex =1
                    Name ="TxtTotalDéb_Mont_Prin_m2"
                    ControlSource ="=Sum([Déb_Mont_Prin_m2])"
                    Format ="#,##0.00 €"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =1025
            Name ="PiedEtat"
        End
    End
End
CodeBehindForm
' See "EtaListeDebiteurEncaisseur.cls"
