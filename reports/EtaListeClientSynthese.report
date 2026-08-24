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
    ItemSuffix =63
    Left =-465
    Top =2895
    RecSrcDt = Begin
        0xfe9186e9b8e9e240
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
            Name ="EntêteEtat0"
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Top =491
                    Width =5943
                    Height =480
                    FontSize =14
                    FontWeight =700
                    Name ="Texte23"
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
                    Name ="Boîte46"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =5669
                    Top =1247
                    Width =1560
                    Height =285
                    FontSize =10
                    Name ="EtiCli_Date_contrat"
                    Caption ="Contrat"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =7256
                    Top =1247
                    Width =1710
                    Height =285
                    FontSize =10
                    Name ="EtiCli_Date_der_cré"
                    Caption ="Dernière Créance"
                End
            End
        End
        Begin Section
            Height =285
            Name ="Détail1"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =113
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
                    Left =1360
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
                    Left =5782
                    Width =1341
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="TxtCli_Date_contrat"
                    ControlSource ="Cli_Date_contrat"
                    Format ="Short Date"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =7258
                    Height =285
                    FontSize =10
                    TabIndex =3
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
