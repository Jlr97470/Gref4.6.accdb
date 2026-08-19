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
    GridX =5
    GridY =5
    Width =4530
    ItemSuffix =37
    Left =6750
    Top =990
    Right =11640
    Bottom =5220
    TimerInterval =10
    RecSrcDt = Begin
        0xdbc63df970e5e240
    End
    Caption ="Saisir un nouveau contrat"
    OnOpen ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
            Width =850
            Height =850
        End
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
        End
        Begin TextBox
            Width =1701
            LabelX =-1701
        End
        Begin ComboBox
            Width =1701
            LabelX =-1701
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
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
            End
        End
        Begin Section
            Height =1700
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =0
                    OverlapFlags =93
                    Left =180
                    Top =135
                    Width =4245
                    Height =1456
                    Name ="RecContrat"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =3
                    Left =285
                    Top =30
                    Width =705
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="EtiContrat"
                    Caption ="Contrat"
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =215
                    Left =1700
                    Top =675
                    Width =1140
                    Height =255
                    TabIndex =1
                    Name ="TxtCli_Mont_Hon_m"
                    Format ="Standard"
                    StatusBarText ="Montant des honoraires"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =285
                            Top =690
                            Width =1305
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiCli_Mont_Hon_m"
                            Caption ="Honoraires    :"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =215
                    Left =1700
                    Top =1247
                    Width =1140
                    Height =255
                    TabIndex =3
                    Name ="TxtCli_Pourcentage"
                    Format ="Standard"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =285
                            Top =1247
                            Width =1290
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiCli_Pourcentage"
                            Caption ="Pourcentage :"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =675
                    Top =975
                    Width =300
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiOu"
                    Caption ="ou"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2925
                    Top =675
                    Width =1382
                    TabIndex =2
                    ForeColor =8388608
                    Name ="TxtMon_Libelle"
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1695
                    Top =315
                    Width =1146
                    Name ="TxtCli_Date_contrat"
                    Format ="Short Date"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =285
                            Top =315
                            Width =1335
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiCli_Date_contrat"
                            Caption ="Date             :"
                        End
                    End
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
                    Left =1700
                    Top =113
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
' See "FrmFicheClientContratNouveau.cls"
