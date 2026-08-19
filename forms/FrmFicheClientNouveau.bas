Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowEdits = NotDefault
    DataEntry = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =6240
    ItemSuffix =32
    Left =4155
    Right =10755
    Bottom =6390
    TimerInterval =10
    RecSrcDt = Begin
        0xcb770377f3ece240
    End
    RecordSource ="SELECT CLIENT.* FROM CLIENT; "
    Caption ="Saisir un nouveau client"
    BeforeUpdate ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Rectangle
            Width =850
            Height =850
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            Width =1701
            LabelX =-1701
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin ComboBox
            Width =1701
            LabelX =-1701
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin FormHeader
            Height =690
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =113
                    Top =113
                    Width =471
                    Height =471
                    Name ="CmdFermer"
                    Caption ="Fermer"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000019000000180000000100040000000000800100000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888888888888888888888888880888888 ,
                        0x8888888888888888888888880088888888888888888888888888888030888888 ,
                        0x8888888888888888888888033088888888888888888888888888803330888888 ,
                        0x8888888888888888888803333088888888888888888888888000333330000000 ,
                        0x00000888888888888880333330ffff0888888888888888888880333330ffff08 ,
                        0x88888888888888888880333300ffff0888888888888888888880333330ffff08 ,
                        0x88888488888888888880333330ffff0888884488888888888880333330ffff08 ,
                        0x88844488888888888880333330ffff0888444444488888888880333330ffff08 ,
                        0x84444444488888888880333330ffff0884444444488888888880333330ffff08 ,
                        0x8844444448888888888033330fffff08888444888888800088803330ffffff08 ,
                        0x88884488888880ff8880330fffffff0888888488888880ff888030ffffffff08 ,
                        0x888888888888800088800fffffffff0888888888888888888880000000000008 ,
                        0x8888888888888888000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                End
            End
        End
        Begin Section
            Height =3975
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =85
                    ColumnCount =2
                    ListRows =3
                    ListWidth =4536
                    Left =2258
                    Top =885
                    Width =2279
                    Height =255
                    TabIndex =2
                    Name ="CmbCli_Représentant"
                    ControlSource ="Cli_Représentant"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT REPRESENTANT.* FROM REPRESENTANT; "
                    ColumnWidths ="0;4536"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =225
                            Top =885
                            Width =1980
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiCli_Représentant"
                            Caption ="Représentant           :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =0
                    OverlapFlags =93
                    Left =225
                    Top =3300
                    Width =5775
                    Height =616
                    Name ="BteContrat"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    Left =435
                    Top =3165
                    Width =705
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="EtiContrat"
                    Caption ="Contrat"
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =215
                    Left =1538
                    Top =3525
                    Width =1260
                    Height =255
                    TabIndex =11
                    Name ="TxtCli_Mont_Hon_m"
                    ControlSource ="Cli_Mont_Hon_m2"
                    Format ="Standard"
                    StatusBarText ="Montant des honoraires"
                    BeforeUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =330
                            Top =3525
                            Width =1125
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiCli_Mont_Hon"
                            Caption ="Honoraires :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =215
                    Left =4853
                    Top =3529
                    Width =960
                    Height =255
                    TabIndex =12
                    Name ="TxtCli_Pourcentage"
                    ControlSource ="Cli_Pourcentage"
                    Format ="Standard"
                    BeforeUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =2970
                            Top =3525
                            Width =1800
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiCli_Pourcentage"
                            Caption ="ou     Pourcentage :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    Left =2258
                    Top =510
                    Width =3675
                    Height =255
                    TabIndex =1
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"

                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =2
                    Left =2258
                    Top =150
                    Width =675
                    Height =255
                    Name ="TxtCli_Code"
                    ControlSource ="Cli_Code"
                    InputMask =">AAAAA;;_"

                End
                Begin Label
                    OverlapFlags =85
                    Left =225
                    Top =150
                    Width =1980
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiCli_Code"
                    Caption ="Code                       :"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    OverlapFlags =85
                    Left =225
                    Top =510
                    Width =1980
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiCli_Nom"
                    Caption ="Nom                       :"
                    FontName ="MS Sans Serif"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    ColumnCount =2
                    ListRows =2
                    Left =2834
                    Top =1247
                    Width =1131
                    Height =255
                    TabIndex =3
                    Name ="CmbCli_Mont_Hon_mo"
                    ControlSource ="Cli_Mont_Hon_mo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Mon_Id, Mon_Nom FROM MONNAIE; "
                    ColumnWidths ="0;1134"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="2"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =225
                            Top =1258
                            Width =2550
                            Height =240
                            FontSize =7
                            FontWeight =700
                            BackColor =8421504
                            Name ="EtiCli_Mont_Hon_mo"
                            Caption ="Monnaie de traitement        :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4536
                    Left =2261
                    Top =1620
                    Width =2271
                    Height =255
                    TabIndex =4
                    Name ="CmbCli_Type"
                    ControlSource ="Cli_Type"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CLIENT_TYPE.CT_Id, CLIENT_TYPE.CT_Nom FROM CLIENT_TYPE ORDER BY [CT_Nom];"
                        " "
                    ColumnWidths ="0;4536"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =225
                            Top =1620
                            Width =1965
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiCli_Type"
                            Caption ="Type                       :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =15
                    ListWidth =4536
                    Left =2261
                    Top =2385
                    Width =2271
                    Height =255
                    TabIndex =6
                    Name ="CmbCli_Groupe"
                    ControlSource ="Cli_Groupe"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CLIENT_GROUPE.CG_Id, CLIENT_GROUPE.CG_Nom FROM CLIENT_GROUPE INNER JOIN C"
                        "LIENT_GROUPE AS CLIENT_GROUPE_1 ON CLIENT_GROUPE.CG_Parent_Id=CLIENT_GROUPE_1.CG"
                        "_Id WHERE (((CLIENT_GROUPE.CG_Parent_Id)=Formulaires!FrmFicheClientNouveau!CmbCG"
                        "_Parent_Id)) ORDER BY CLIENT_GROUPE.CG_Nom; "
                    ColumnWidths ="0;4536"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =225
                            Top =2385
                            Width =1965
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiCli_Groupe"
                            Caption ="Sous Groupe           :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4650
                    Top =1635
                    Width =260
                    Height =245
                    TabIndex =5
                    Name ="CmdTypeClient"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000000a000000040000000100040000000000200000000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888800000080080080080000008008008008000000 ,
                        0x8888888888000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4650
                    Top =2400
                    Width =260
                    Height =245
                    TabIndex =7
                    Name ="CmdCli_Groupe"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000000a000000040000000100040000000000200000000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888800000080080080080000008008008008000000 ,
                        0x8888888888000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =15
                    ListWidth =4536
                    Left =2261
                    Top =1980
                    Width =2271
                    Height =255
                    TabIndex =8
                    Name ="CmbCG_Parent_Id"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CLIENT_GROUPE.CG_Id, CLIENT_GROUPE.CG_Nom FROM CLIENT_GROUPE WHERE (((CLI"
                        "ENT_GROUPE.CG_Parent_Id) Is Null)) ORDER BY CLIENT_GROUPE.CG_Nom; "
                    ColumnWidths ="0;4536"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =225
                            Top =1980
                            Width =1965
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiCG_Parent_Id"
                            Caption ="Groupe                    :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4650
                    Top =1995
                    Width =260
                    Height =245
                    TabIndex =9
                    Name ="CmdCG_Parent_Id"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000000a000000040000000100040000000000200000000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888800000080080080080000008008008008000000 ,
                        0x8888888888000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2271
                    Top =2820
                    Width =996
                    TabIndex =10
                    Name ="TxtCli_Date_contrat"
                    ControlSource ="Cli_Date_contrat"
                    Format ="Short Date"
                    StatusBarText ="Client Date Du Contrat"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =225
                            Top =2820
                            Width =1965
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiCli_Date_contrat"
                            Caption ="Date contrat            :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =570
            BackColor =12632256
            Name ="PiedFormulaire"
            Begin
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =2595
                    Top =60
                    Width =1125
                    Height =450
                    FontWeight =300
                    Name ="CmdOk"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"

                End
            End
        End
    End
End
CodeBehindForm
' See "FrmFicheClientNouveau.cls"
