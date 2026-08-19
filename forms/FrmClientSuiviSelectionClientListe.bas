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
    Width =5100
    ItemSuffix =45
    Left =4485
    Top =285
    Right =9945
    Bottom =5925
    RecSrcDt = Begin
        0xba6b095910e8e240
    End
    RecordSource ="SELECT CLIENT.* FROM CLIENT; "
    Caption ="Liste des Clients"
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
            Height =1080
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =45
                    Top =840
                    Width =660
                    Height =240
                    Name ="EtiCli_Code"
                    Caption ="Code"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =780
                    Top =840
                    Width =3900
                    Height =240
                    Name ="EtiCli_Nom"
                    Caption ="Nom"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =113
                    Top =113
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
                        0x8888888888888888000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    AccessKey =82
                    Left =735
                    Top =120
                    Width =441
                    Height =456
                    TabIndex =1
                    Name ="CmdRechercher"
                    Caption ="&Rechercher"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000019000000180000000100040000000000800100000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888800000000888888800000000888888888000000008888888 ,
                        0x0000000088888888800ff0000888888800ff000088888888800ff00008888888 ,
                        0x00ff000088888888800ff0000888888800ff0000888888888000000000008000 ,
                        0x0000000088888888800000000000800000000000888888888000ff000000000f ,
                        0xf0000000888888888000ff000088000ff0000000888888888000ff000088000f ,
                        0xf0000000888888888800ff000088000ff0000008888888888880000000000000 ,
                        0x000000888888888788800000000000000000008888888874888800f00000800f ,
                        0x0000088888888847888800000000800000000888888888488888000000008000 ,
                        0x0000088888888848888888000088888000088888888888478888880ff0888880 ,
                        0xff088888888888748888880ff0888880ff088888888888878888880000888880 ,
                        0x0008888888888888888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin Label
                    SpecialEffect =1
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =4155
                    Top =225
                    Width =720
                    Height =255
                    Name ="EtiRien"
                    Caption ="Rien"
                    OnClick ="=ToutRien(False)"
                End
                Begin Label
                    SpecialEffect =1
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =3405
                    Top =225
                    Width =720
                    Height =255
                    Name ="EtiTout"
                    Caption ="Tout"
                    OnClick ="=ToutRien(True)"
                End
                Begin Rectangle
                    SpecialEffect =1
                    BackStyle =0
                    OverlapFlags =215
                    Left =3300
                    Top =120
                    Width =1665
                    Height =450
                    Name ="BteToutRien"
                End
            End
        End
        Begin Section
            Height =270
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =30
                    Top =15
                    Width =681
                    Height =240
                    Name ="TxtCli_Code"
                    ControlSource ="Cli_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =780
                    Top =15
                    Width =3906
                    Height =240
                    TabIndex =1
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4785
                    Top =45
                    Width =202
                    TabIndex =2
                    Name ="ChkCli_Flag"
                    ControlSource ="=Not IsNull(DLookUp(\"TCli_Code\",\"TMP_CLIENT\",\"TCli_Code='\" & [Cli_Code] & "
                        "\"'\"))"
                    OnMouseDown ="[Event Procedure]"
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
                    Left =2055
                    Top =75
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
' See "FrmClientSuiviSelectionClientListe.cls"
