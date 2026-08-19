Version =20
VersionRequired =20
Begin Form
    AllowEditing = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =7590
    ItemSuffix =86
    Left =5925
    Top =105
    Right =13875
    Bottom =6510
    RecSrcDt = Begin
        0x8cef283c1beae240
    End
    RecordSource ="SELECT FACTURE.Fac_Num, FACTURE.Fac_CodeCli, CLIENT.Cli_Nom, CLIENT.Cli_Groupe, "
        "FACTURE.Fac_Date, FACTURE.Fac_Editer, FACTURE.Fac_Deja_Editer FROM CLIENT INNER "
        "JOIN FACTURE ON CLIENT.Cli_Code=FACTURE.Fac_CodeCli; "
    Caption ="Reédition des factures Clients"
    OnOpen ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
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
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =187
            Height =187
        End
        Begin CheckBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =187
            Height =187
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =1701
            Height =255
            LabelX =-1701
        End
        Begin ListBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin ComboBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =1701
            Height =255
            LabelX =-1701
        End
        Begin FormHeader
            Height =2265
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =690
                    Top =15
                    Width =6801
                    Height =1940
                    BackColor =8421504
                    Name ="RecRech"
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    Left =690
                    Top =30
                    Width =6797
                    Height =1880
                    TabIndex =2
                    Name ="CadClass"
                    DefaultValue ="1"
                    Begin
                        Begin OptionButton
                            OverlapFlags =215
                            Left =4335
                            Top =345
                            OptionValue =1
                            Name ="OptFac_Num"
                        End
                        Begin OptionButton
                            OverlapFlags =215
                            Left =6945
                            Top =390
                            OptionValue =4
                            Name ="OptFac_Date"
                        End
                        Begin OptionButton
                            OverlapFlags =215
                            Left =4335
                            Top =628
                            OptionValue =2
                            Name ="OptFac_CodeCli"
                        End
                        Begin OptionButton
                            OverlapFlags =215
                            Left =4335
                            Top =930
                            OptionValue =3
                            Name ="OptCli_Nom"
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =1
                    OverlapFlags =223
                    Left =5400
                    Top =990
                    Width =1665
                    Height =450
                    BackColor =12632256
                    Name ="RecToutRien"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    Left =1140
                    Top =2025
                    Width =1140
                    Height =240
                    Name ="EtiFac_CodeCli"
                    Caption ="Code"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =95
                    Left =2280
                    Top =2025
                    Width =3600
                    Height =240
                    Name ="EtiCli_Nom"
                    Caption ="Nom"
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =120
                    Top =615
                    Width =455
                    Height =455
                    FontSize =7
                    FontWeight =400
                    TabIndex =1
                    Name ="CmdImprimer"
                    Caption ="Imprimer"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000019000000180000000100040000000000800100000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888888888888888888888000008888000000000000 ,
                        0x000088888880ffff8880888888888888888008888880ffff8880888888888888 ,
                        0x888080888880ffff8000000000000000000008088880ff00808888888888bbbb ,
                        0x888800808880ffff808888888888bbbb888808008880ffff8088888888887777 ,
                        0x888808808880ff008088888888888888888808880880ffff8000000000000000 ,
                        0x000008880880ffff8088888888888888888080880880ffff8808888888888888 ,
                        0x888088080880ffff8880000000000000000008800880ffff888880ffffffffff ,
                        0xff0880880880fff0888880ffffffffffff0088070880fff08888880ff0000000 ,
                        0xfff000008880ffff8888880ffffffffffff008888880000088888880ffffffff ,
                        0xffff088888888888888888880f00000000ff088888888888888888880fffffff ,
                        0xfffff088888888888888888880fffffffffff088888888888888888880000000 ,
                        0x0000008888888888000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000
                    End
                    FontName ="Small Fonts"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =215
                            Top =2025
                            Width =1155
                            Height =240
                            Name ="EtiFac_Num"
                            Caption ="Numero"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =3041
                    Top =585
                    Width =1131
                    Height =240
                    TabIndex =4
                    Name ="TxtRech_Fac_CodeCli"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =1732
                            Top =585
                            Width =1230
                            Height =240
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiRech_Fac_Cli_Code"
                            Caption ="Code  :"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =1808
                    Top =870
                    Width =2361
                    Height =240
                    TabIndex =5
                    Name ="TxtRech_Cli_Nom"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =1230
                            Top =870
                            Width =555
                            Height =240
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiRech_Cli_Nom"
                            Caption ="Nom :"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =4215
                    Top =105
                    Width =525
                    Height =195
                    FontWeight =400
                    BackColor =8421504
                    ForeColor =65535
                    Name ="EtiClass"
                    Caption ="Class."
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =5657
                    Top =315
                    Width =1131
                    Height =240
                    TabIndex =6
                    Name ="TxtRech_Fac_DateDebut"
                    Format ="Short Date"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4800
                            Top =315
                            Width =780
                            Height =240
                            FontSize =7
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiRech_Fac_DateDebut"
                            Caption ="Entre le :"
                        End
                    End
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =215
                    AccessKey =82
                    Left =5160
                    Top =1545
                    Width =2055
                    Height =315
                    TabIndex =8
                    Name ="CmdOK"
                    Caption ="&Rechercher"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5657
                    Top =585
                    Width =1131
                    TabIndex =7
                    Name ="TxtRech_Fac_DateFin"
                    Format ="Short Date"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4680
                            Top =585
                            Width =900
                            Height =240
                            FontSize =7
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiRech_Fac_DateFin"
                            Caption ="Et le :"
                        End
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =87
                    Left =5880
                    Top =2025
                    Width =1155
                    Height =240
                    Name ="EtiFac_Date"
                    Caption ="Date"
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =105
                    Top =90
                    Width =471
                    Height =456
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
                        0x8888888888888888
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =3034
                    Top =300
                    Width =1131
                    Height =240
                    TabIndex =3
                    Name ="TxtRech_Fac_Num"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =1725
                            Top =300
                            Width =1230
                            Height =240
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiRech_Fac_Num"
                            Caption ="Numero  :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4536
                    Left =2003
                    Top =1470
                    Width =2166
                    TabIndex =9
                    Name ="CmbCli_Groupe"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CLIENT_GROUPE.CG_Id, CLIENT_GROUPE.CG_Nom FROM CLIENT_GROUPE INNER JOIN C"
                        "LIENT_GROUPE AS CLIENT_GROUPE_1 ON CLIENT_GROUPE.CG_Parent_Id=CLIENT_GROUPE_1.CG"
                        "_Id WHERE (((CLIENT_GROUPE.CG_Parent_Id)=Formulaires!FrmClientFactureDefinitifSe"
                        "lectionListe!CmbGroupe)) ORDER BY CLIENT_GROUPE.CG_Nom; "
                    ColumnWidths ="0;4536"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =765
                            Top =1470
                            Width =1230
                            Height =240
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiCli_Groupe"
                            Caption ="Sous Groupe     :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4536
                    Left =2003
                    Top =1170
                    Width =2166
                    TabIndex =10
                    Name ="CmbGroupe"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CLIENT_GROUPE.CG_Id, CLIENT_GROUPE.CG_Nom FROM CLIENT_GROUPE WHERE (((CLI"
                        "ENT_GROUPE.CG_Parent_Id) Is Null)) ORDER BY CLIENT_GROUPE.CG_Nom; "
                    ColumnWidths ="0;4536"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =765
                            Top =1170
                            Width =1215
                            Height =240
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiGroupe"
                            Caption ="Groupe             :"
                        End
                    End
                End
                Begin Label
                    SpecialEffect =1
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    Left =6255
                    Top =1095
                    Width =720
                    Height =255
                    Name ="EtiRien"
                    Caption ="Rien"
                    OnClick ="=ToutRien(False)"
                End
                Begin Label
                    SpecialEffect =1
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    Left =5505
                    Top =1095
                    Width =720
                    Height =255
                    Name ="EtiTout"
                    Caption ="Tout"
                    OnClick ="=ToutRien(True)"
                End
            End
        End
        Begin Section
            Height =255
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =1170
                    Top =15
                    Width =1101
                    Height =240
                    Name ="TxtFac_CodeCli"
                    ControlSource ="Fac_CodeCli"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =5895
                    Top =15
                    Width =1131
                    Height =240
                    TabIndex =1
                    Name ="TxtFac_Date"
                    ControlSource ="Fac_Date"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =2295
                    Top =15
                    Width =3576
                    Height =240
                    TabIndex =2
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =7080
                    Top =30
                    TabIndex =3
                    Name ="ChkFac_Editer"
                    ControlSource ="Fac_Editer"
                    StatusBarText ="Reglement Editer"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =7365
                    Top =30
                    TabIndex =4
                    Name ="ChkFac_Deja_Editer"
                    ControlSource ="Fac_Deja_Editer"
                    StatusBarText ="Reglement Deja Editer"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =15
                    Top =15
                    Width =1131
                    Height =240
                    TabIndex =5
                    Name ="TxtFac_Num"
                    ControlSource ="Fac_Num"
                End
            End
        End
        Begin FormFooter
            Height =226
            BackColor =12632256
            Name ="PiedFormulaire"
        End
    End
End
CodeBehindForm
' See "FrmClientFactureDefinitifSelectionListe.cls"
