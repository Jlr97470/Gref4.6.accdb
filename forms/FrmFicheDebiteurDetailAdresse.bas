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
    ItemSuffix =88
    Left =4755
    Right =10770
    Bottom =8145
    Filter ="[AD_Code_Débiteur]= 'CA00100001' AND [AD_TypeAdresse_ID]=8"
    RecSrcDt = Begin
        0xcc050689eee5e240
    End
    RecordSource ="SELECT ADRESSE_DEBITEUR.* FROM ADRESSE_DEBITEUR; "
    Caption ="Détail d'une Adresse"
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
            Name ="EntêteFormulaire1"
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
                        "xtAD_TypeAdresse_ID])"
                End
            End
        End
        Begin Section
            Height =5625
            BackColor =12632256
            Name ="Détail0"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =120
                    Top =105
                    Width =450
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16711680
                    Name ="TxtDéb_Qualité"
                    ControlSource ="=DLookUp(\"Déb_Qualité\",\"DEBITEUR\",\"Déb_Code='\" & [TxtAD_Code_Débiteur] & \""
                        "'\")"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =600
                    Top =105
                    Width =4950
                    Height =240
                    FontWeight =700
                    TabIndex =1
                    BackColor =12632256
                    ForeColor =16711680
                    Name ="TxtDéb_Nom"
                    ControlSource ="=DLookUp(\"Déb_Nom\",\"DEBITEUR\",\"Déb_Code='\" & [TxtAD_Code_Débiteur] & \"'\""
                        ")"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1708
                    Top =1375
                    Width =3390
                    TabIndex =4
                    Name ="TxtAD_Adr1"
                    ControlSource ="AD_Adr1"
                    StatusBarText ="Adresse home"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =375
                            Top =1375
                            Width =1290
                            Height =240
                            Name ="EtiAD_Adr1"
                            Caption ="Adresse 1     :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1709
                    Top =1670
                    Width =3390
                    TabIndex =5
                    Name ="TxtAD_Adr2"
                    ControlSource ="AD_Adr2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =375
                            Top =1675
                            Width =1290
                            Height =240
                            Name ="EtiAD_Adr2"
                            Caption ="Adresse 2     :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1708
                    Top =2545
                    Width =855
                    TabIndex =8
                    Name ="TxtAD_CodePostal"
                    ControlSource ="AD_CodePostal"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =375
                            Top =2545
                            Width =1290
                            Height =240
                            Name ="EtiAD_CodePostal"
                            Caption ="Code Postal  :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1708
                    Top =2830
                    Width =3390
                    TabIndex =9
                    Name ="TxtAD_Ville"
                    ControlSource ="AD_Ville"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =375
                            Top =2830
                            Width =1290
                            Height =240
                            Name ="EtiAD_Ville"
                            Caption ="Ville             :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1708
                    Top =3418
                    Width =2280
                    TabIndex =11
                    Name ="TxtAD_Tel"
                    ControlSource ="AD_Tel"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =375
                            Top =3415
                            Width =1275
                            Height =240
                            Name ="EtiAD_Tel"
                            Caption ="Téléphone 1 :"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =1707
                    Top =1015
                    Width =3390
                    TabIndex =3
                    Name ="TxtAD_NomComplémentaire"
                    ControlSource ="AD_NomComplémentaire"
                    StatusBarText ="Adresse home"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextAlign =1
                            Left =800
                            Top =675
                            Width =2040
                            Height =240
                            FontSize =7
                            Name ="EtiAD_NomComplémentaire"
                            Caption ="Nom complémentaire :"
                            FontName ="Small Fonts"
                        End
                    End
                End
                Begin ComboBox
                    FontItalic = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    Left =800
                    Top =1015
                    Width =787
                    Height =227
                    TabIndex =2
                    Name ="CmbAD_GenreComplémentaire"
                    ControlSource ="AD_GenreComplémentaire"
                    RowSourceType ="Value List"
                    RowSource ="ASS;ENF;ENT;M;M/M;MLE;MME;STE"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1708
                    Top =3988
                    Width =2280
                    TabIndex =13
                    Name ="TxtAD_Fax"
                    ControlSource ="AD_Fax"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =375
                            Top =3985
                            Width =1275
                            Height =240
                            Name ="EtiAD_Fax"
                            Caption ="Fax              :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1708
                    Top =1960
                    Width =3390
                    TabIndex =6
                    Name ="TxtAD_Adr3"
                    ControlSource ="AD_Adr3"
                    StatusBarText ="Adresse home"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =375
                            Top =1960
                            Width =1290
                            Height =240
                            Name ="EtiAD_Adr3"
                            Caption ="Adresse 3     :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1709
                    Top =2255
                    Width =3390
                    TabIndex =7
                    Name ="TxtAD_Adr4"
                    ControlSource ="AD_Adr4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =375
                            Top =2260
                            Width =1290
                            Height =240
                            Name ="EtiAD_Adr4"
                            Caption ="Adresse 4     :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1708
                    Top =3115
                    Width =1260
                    TabIndex =10
                    Name ="TxtAD_Pays"
                    ControlSource ="AD_Pays"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =375
                            Top =3115
                            Width =1275
                            Height =240
                            Name ="EtiAD_Pays"
                            Caption ="Pays            :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1708
                    Top =3703
                    Width =2280
                    TabIndex =12
                    Name ="TxtAD_Tel2"
                    ControlSource ="AD_Tel2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =375
                            Top =3700
                            Width =1275
                            Height =240
                            Name ="EtiAD_Tel2"
                            Caption ="Téléphone 2 :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1708
                    Top =4288
                    Width =2280
                    TabIndex =14
                    Name ="TxtAD_email"
                    ControlSource ="AD_email"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =375
                            Top =4285
                            Width =1275
                            Height =240
                            Name ="EtiAD_email"
                            Caption ="e-mail           :"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    Left =1708
                    Top =4588
                    Width =3390
                    Height =885
                    TabIndex =15
                    Name ="TxtAD_Note"
                    ControlSource ="AD_Note"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =390
                            Top =4585
                            Width =1260
                            Height =240
                            Name ="EtiAD_Note"
                            Caption ="Note            :"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =119
                    IMESentenceMode =3
                    Left =120
                    Top =345
                    ColumnWidth =1725
                    TabIndex =16
                    Name ="TxtAD_Code_Débiteur"
                    ControlSource ="AD_Code_Débiteur"
                    StatusBarText ="Adresse Debiteur Code Debiteur\015\012Adresse Debiteur Code Debiteur"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =119
                    IMESentenceMode =3
                    Left =3975
                    Top =345
                    Width =1581
                    ColumnWidth =855
                    TabIndex =17
                    Name ="TxtAD_TypeAdresse_ID"
                    ControlSource ="AD_TypeAdresse_ID"
                    StatusBarText ="Adresse Debiteur Type Adresse"
                End
            End
        End
        Begin FormFooter
            Height =690
            BackColor =12632256
            Name ="PiedFormulaire2"
            Begin
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =2265
                    Top =120
                    Width =1155
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
' See "FrmFicheDebiteurDetailAdresse.cls"
