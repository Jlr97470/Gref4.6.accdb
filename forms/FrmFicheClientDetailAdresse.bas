Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    Cycle =1
    GridX =5
    GridY =5
    Width =5670
    ItemSuffix =71
    Left =4890
    Right =10920
    Bottom =7755
    RecSrcDt = Begin
        0x3e1f515170e5e240
    End
    RecordSource ="SELECT ADRESSE_CLIENT.* FROM ADRESSE_CLIENT; "
    Caption ="Détail d'une adresse"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            TextAlign =3
            FontWeight =700
            BackColor =12632256
        End
        Begin Rectangle
            SpecialEffect =2
            Width =850
            Height =850
        End
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
        End
        Begin OptionButton
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            Width =187
            Height =187
            LabelX =-236
        End
        Begin CheckBox
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            Width =187
            Height =187
            LabelX =-236
        End
        Begin TextBox
            SpecialEffect =2
            LabelAlign =3
            Width =1701
            Height =255
            LabelX =-236
        End
        Begin ListBox
            SpecialEffect =2
            LabelAlign =3
            Width =1701
            Height =1417
            LabelX =-236
        End
        Begin ComboBox
            SpecialEffect =2
            LabelAlign =3
            Width =1701
            Height =255
            LabelX =-236
        End
        Begin ToggleButton
            TextFontFamily =2
            Width =283
            Height =283
        End
        Begin FormHeader
            Height =680
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
                        0x00000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =1474
                    Top =113
                    Width =2716
                    Height =452
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="TxtAT_TypeAdresse_Libellé"
                    ControlSource ="=DLookUp(\"AT_TypeAdresse_Libellé\",\"ADRESSE_TYPE\",\"AT_TypeAdresse_ID=\" & [T"
                        "xtAC_TypeAdresse_ID])"
                End
            End
        End
        Begin Section
            Height =5235
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =285
                    Top =105
                    Width =5100
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16711680
                    Name ="TxtCli_Nom"
                    ControlSource ="=DLookUp(\"Cli_Nom\",\"CLIENT\",\"Cli_Code='\" & [TxtAC_Code_Client] & \"'\")"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =1702
                    Top =940
                    Width =3420
                    TabIndex =2
                    Name ="TxtAC_NomComplémentaire"
                    ControlSource ="AC_NomComplémentaire"
                    StatusBarText ="Adresse home"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextAlign =1
                            Left =795
                            Top =645
                            Width =2040
                            Height =240
                            FontSize =7
                            Name ="EtiAC_NomComplémentaire"
                            Caption ="Nom complémentaire :"
                            FontName ="Small Fonts"
                        End
                    End
                End
                Begin ComboBox
                    FontItalic = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    Left =795
                    Top =940
                    Width =787
                    Height =227
                    TabIndex =1
                    Name ="CmbAC_GenreComplémentaire"
                    ControlSource ="AC_GenreComplémentaire"
                    RowSourceType ="Value List"
                    RowSource ="ASS;ENF;ENT;M;M/M;MLE;MME;STE"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1693
                    Top =1350
                    Width =3420
                    TabIndex =3
                    Name ="TxtAC_Adr1"
                    ControlSource ="AC_Adr1"
                    StatusBarText ="Adresse home"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =360
                            Top =1350
                            Width =1290
                            Height =240
                            Name ="EtiAC_Adr1"
                            Caption ="Adresse 1     :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1694
                    Top =1645
                    Width =3420
                    TabIndex =4
                    Name ="TxtAC_Adr2"
                    ControlSource ="AC_Adr2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =361
                            Top =1645
                            Width =1290
                            Height =240
                            Name ="Eti_Adr2"
                            Caption ="Adresse 2     :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1693
                    Top =2520
                    Width =855
                    TabIndex =7
                    Name ="TxtAC_CodePostal"
                    ControlSource ="AC_CodePostal"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =360
                            Top =2520
                            Width =1290
                            Height =240
                            Name ="EtiAC_CodePostal"
                            Caption ="Code Postal  :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1693
                    Top =2805
                    Width =3420
                    TabIndex =8
                    Name ="TxtAC_Ville"
                    ControlSource ="AC_Ville"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =360
                            Top =2805
                            Width =1290
                            Height =240
                            Name ="EtiAC_Ville"
                            Caption ="Ville             :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1693
                    Top =1935
                    Width =3420
                    TabIndex =5
                    Name ="TxtAC_Adr3"
                    ControlSource ="AC_Adr3"
                    StatusBarText ="Adresse home"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =360
                            Top =1935
                            Width =1290
                            Height =240
                            Name ="EtiAC_Adr3"
                            Caption ="Adresse 3     :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1694
                    Top =2230
                    Width =3420
                    TabIndex =6
                    Name ="TxtAC_Adr4"
                    ControlSource ="AC_Adr4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =361
                            Top =2230
                            Width =1290
                            Height =240
                            Name ="EtiAC_Adr4"
                            Caption ="Adresse 4     :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1693
                    Top =3093
                    Width =2280
                    TabIndex =9
                    Name ="TxtAC_Tel"
                    ControlSource ="AC_Tel"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =360
                            Top =3090
                            Width =1275
                            Height =240
                            Name ="EtiAC_Tel"
                            Caption ="Téléphone 1 :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1693
                    Top =3678
                    Width =2280
                    TabIndex =11
                    Name ="TxtAC_Fax"
                    ControlSource ="AC_Fax"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =360
                            Top =3682
                            Width =1290
                            Height =240
                            Name ="EtiAC_Fax"
                            Caption ="Fax              :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1693
                    Top =3378
                    Width =2280
                    TabIndex =10
                    Name ="TxtAC_Tel2"
                    ControlSource ="AC_Tel2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =360
                            Top =3375
                            Width =1275
                            Height =240
                            Name ="EtiAC_Tel2"
                            Caption ="Téléphone 2 :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1693
                    Top =3978
                    Width =2280
                    TabIndex =12
                    Name ="TxtAC_email"
                    ControlSource ="AC_email"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =360
                            Top =3975
                            Width =1305
                            Height =240
                            Name ="EtiAC_email"
                            Caption ="e-mail           :"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    Left =1693
                    Top =4263
                    Width =3405
                    Height =885
                    TabIndex =13
                    Name ="TxtAC_Note"
                    ControlSource ="AC_Note"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =360
                            Top =4260
                            Width =1305
                            Height =240
                            Name ="EtiAC_Note"
                            Caption ="Note            :"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =119
                    IMESentenceMode =3
                    Left =285
                    Top =345
                    ColumnWidth =1470
                    TabIndex =14
                    Name ="TxtAC_Code_Client"
                    ControlSource ="AC_Code_Client"
                    StatusBarText ="Adresse Client Code Client"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =119
                    IMESentenceMode =3
                    Left =3690
                    Top =345
                    ColumnWidth =855
                    TabIndex =15
                    Name ="TxtAC_TypeAdresse_ID"
                    ControlSource ="AC_TypeAdresse_ID"
                    StatusBarText ="Adresse Client Type"
                End
            End
        End
        Begin FormFooter
            Height =690
            BackColor =12632256
            Name ="PiedFormulaire"
            Begin
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =2280
                    Top =120
                    Width =1125
                    Height =450
                    FontWeight =300
                    Name ="CmdOK"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "FrmFicheClientDetailAdresse.cls"
