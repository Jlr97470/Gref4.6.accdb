Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =3975
    ItemSuffix =19
    Left =4260
    Top =255
    Right =8595
    Bottom =3795
    RecSrcDt = Begin
        0xe5674b4d0de7e240
    End
    Caption ="Relancer"
    AllowDatasheetView =0
    Begin
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
            Height =680
            BackColor =12632256
            Name ="EntêteFormulaire0"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =113
                    Top =113
                    Width =456
                    Height =426
                    Name ="CmdFermer"
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
                        0x0000000000000000
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
            Height =1133
            BackColor =12632256
            Name ="Détail0"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    Left =2749
                    Top =574
                    Width =1071
                    TabIndex =1
                    Name ="TxtDateFin"
                    Format ="Short Date"
                    DefaultValue ="Date()"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextAlign =1
                            Left =1980
                            Top =570
                            Width =570
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiDateFin"
                            Caption ="et le"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    Left =707
                    Top =574
                    Width =1116
                    Name ="TxtDateDebut"
                    Format ="Short Date"
                    DefaultValue ="Date()"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextAlign =1
                            Left =135
                            Top =570
                            Width =435
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiDateDebut"
                            Caption ="le"
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =135
                    Top =210
                    Width =3690
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiRelanceSelection"
                    Caption ="Editer les Relances à envoyer entre"
                End
            End
        End
        Begin FormFooter
            Height =566
            BackColor =12632256
            Name ="PiedFormulaire1"
            Begin
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =1425
                    Top =75
                    Width =1125
                    Height =450
                    FontWeight =300
                    Name ="CmdRelancer"
                    Caption ="Relancer"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "FrmDebiteurRelanceSelectionEdition.cls"
