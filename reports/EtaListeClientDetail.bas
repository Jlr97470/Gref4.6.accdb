Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DateGrouping =1
    GridX =5
    GridY =5
    Width =9060
    ItemSuffix =83
    Left =165
    Top =2640
    RecSrcDt = Begin
        0x1f033dddb8e9e240
    End
    RecordSource ="SELECT CLIENT.* FROM CLIENT; "
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
            ControlSource ="Cli_Code"
        End
        Begin FormHeader
            Height =1700
            Name ="EntêteEtat"
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Top =491
                    Width =5943
                    Height =480
                    FontSize =14
                    FontWeight =700
                    Name ="EtiListeClientDetail"
                    Caption ="Liste des Clients"
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
                    Name ="Ligne26"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =113
                    Top =1587
                    Width =8835
                    Height =0
                    Name ="RecListeClientDetail"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =5625
                    Top =1245
                    Width =1560
                    Height =285
                    FontSize =10
                    Name ="EtiCli_Date_der_cré"
                    Caption ="Dernière Créance"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =7256
                    Top =1247
                    Width =1710
                    Height =285
                    FontSize =10
                    Name ="EtiCli_Date_der_reg"
                    Caption ="Dernière Règlt"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =1133
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Width =1131
                    Height =285
                    FontSize =10
                    Name ="TxtCli_Code"
                    ControlSource ="Cli_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =1247
                    Width =4206
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =7258
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="TxtCli_Date_der_reg"
                    ControlSource ="=DLast(\"Reg_Date\",\"REGLEMENT\",\"Reg_Cli_Code='\" & [Cli_Code] & \"'\")"
                    Format ="Short Date"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =5518
                    Top =315
                    Width =1161
                    Height =225
                    TabIndex =3
                    Name ="TxtCli_Int_Tel"
                    ControlSource ="Cli_Int_Tel"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =2325
                    Top =315
                    Width =2721
                    Height =225
                    TabIndex =4
                    Name ="TxtCli_Int_Nom"
                    ControlSource ="Cli_Int_Nom"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =2325
                    Top =541
                    Width =1341
                    Height =225
                    TabIndex =5
                    Name ="TxtCli_Date_contrat"
                    ControlSource ="Cli_Date_contrat"
                    Format ="Short Date"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =1245
                    Top =315
                    Width =1065
                    Height =225
                    Name ="Texte67"
                    Caption ="Interlocuteur :"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =5098
                    Top =315
                    Width =390
                    Height =225
                    Name ="EtiCli_Int_Tel"
                    Caption ="Tel :"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =6675
                    Top =315
                    Width =585
                    Height =225
                    Name ="EtiCli_Int_Poste"
                    Caption ="Poste :"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =7290
                    Top =315
                    Width =921
                    Height =225
                    TabIndex =6
                    Name ="TxtCli_Int_Poste"
                    ControlSource ="Cli_Int_Poste"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =1245
                    Top =540
                    Width =1065
                    Height =225
                    Name ="EtiCli_Date_contrat"
                    Caption ="Contrat :"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =4863
                    Top =544
                    Width =1251
                    Height =225
                    TabIndex =7
                    Name ="TxtCli_Mont_Hon"
                    ControlSource ="=IIf([Cli_MonnaieEnCours]=1,Format([Cli_Mont_Hon_m1],\"#,##0.00\"\" F\"\"\"),For"
                        "mat([Cli_Mont_Hon_m2],\"#,##0.00 €\"))"
                    Format ="Standard"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =3765
                    Top =540
                    Width =1065
                    Height =225
                    Name ="EtiCli_Mont_Hon"
                    Caption ="Honoraires :"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =7291
                    Top =540
                    Width =876
                    Height =225
                    TabIndex =8
                    Name ="TxtCli_Pourcentage"
                    ControlSource ="Cli_Pourcentage"
                    Format ="Standard"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =6183
                    Top =544
                    Width =1080
                    Height =225
                    Name ="EtiCli_Pourcentage"
                    Caption ="Pourcentage :"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =2325
                    Top =766
                    Width =516
                    Height =225
                    TabIndex =9
                    Name ="TxtCli_Représentant"
                    ControlSource ="Cli_Représentant"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =1185
                    Top =765
                    Width =1125
                    Height =225
                    Name ="EtiCli_Représentant"
                    Caption ="Représentant :"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =4860
                    Top =765
                    Width =771
                    Height =225
                    TabIndex =10
                    Name ="TxtCli_APE"
                    ControlSource ="Cli_APE"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =3960
                    Top =765
                    Width =870
                    Height =225
                    Name ="EtiCli_APE"
                    Caption ="APE :"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =5555
                    Height =285
                    FontSize =10
                    TabIndex =11
                    Name ="TxtCli_Date_der_cré"
                    ControlSource ="Cli_Date_der_cré"
                    Format ="Short Date"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =566
            Name ="PiedEtat1"
        End
    End
End
