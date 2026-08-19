Version =20
VersionRequired =20
Begin Form
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =4110
    ItemSuffix =38
    Left =2115
    Top =330
    Right =6945
    Bottom =5100
    RecSrcDt = Begin
        0xc9e6f9d5d0e9e240
    End
    RecordSource ="SELECT DISTINCTROW FRAIS_GRILLE.Fra_Grille, FRAIS_GRILLE.Fra_Grille_Libellé FROM"
        " FRAIS_GRILLE ORDER BY FRAIS_GRILLE.Fra_Grille; "
    Caption ="Listes des grilles de Frais de Recouvrement"
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
        Begin FormHeader
            Height =907
            BackColor =12632256
            Name ="EntêteFormulaire1"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =113
                    Top =680
                    Width =534
                    Height =227
                    FontSize =7
                    Name ="EtiFra_Grille"
                    Caption ="Code"
                    FontName ="Small Fonts"
                End
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
                        0x8888888888888888000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =755
                    Top =675
                    Width =3255
                    Height =225
                    FontSize =7
                    Name ="EtiFra_Grille_Libellé"
                    Caption ="Libellé"
                    FontName ="Small Fonts"
                End
            End
        End
        Begin Section
            Height =255
            BackColor =12632256
            Name ="Détail0"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =113
                    Width =546
                    Name ="TxtFra_Grille"
                    ControlSource ="Fra_Grille"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    Left =750
                    Width =3261
                    TabIndex =1
                    Name ="TxtFra_Grille_Libellé"
                    ControlSource ="Fra_Grille_Libellé"
                End
            End
        End
        Begin FormFooter
            Height =226
            BackColor =12632256
            Name ="PiedFormulaire2"
        End
    End
End
CodeBehindForm
' See "FrmParam_FraisGrille.cls"
