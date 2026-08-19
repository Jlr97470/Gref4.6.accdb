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
    Width =3405
    ItemSuffix =19
    Left =5895
    Top =1275
    Right =9645
    Bottom =5475
    RecSrcDt = Begin
        0x2283b42f70e9e240
    End
    Caption ="Chèques volés"
    OnCurrent ="[Event Procedure]"
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
            Name ="EntêteFormulaire"
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
            Height =2267
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =1140
                    Top =810
                    Width =1131
                    Name ="TxtQuantite"
                    ControlSource ="=Nz(DCount(\"Déb_Code\",\"SelDebiteurCodeGestion\",\"IsNull(Déb_Date_der_rel) AN"
                        "D CG_Code_Action = 5\"),0)"
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Top =180
                    Width =3405
                    Height =435
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiDebiteurNouveauCheque"
                    Caption ="Edition des courriers concernant les nouveaux chèques volés"
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =105
                    Top =810
                    Width =1020
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiQuantite"
                    Caption ="Quantité :"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =825
                    Top =1385
                    Width =453
                    Height =450
                    FontSize =8
                    TabIndex =1
                    Name ="CmdVisualiser"
                    OnClick ="=ImpVisu(2)"
                    PictureData = Begin
                        0x28000000100000000f0000000100040000000000780000000000000000000000 ,
                        0x1000000010000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff0000000000000088000ffffffffff080000fffffff00007008 ,
                        0x0ffffff0788707880fffff0788e770880fffff08888780880fffff08e8878088 ,
                        0x0fffff07ee8770880ffffff0788708880fffffff000088880ffffffffff08888 ,
                        0x0fffffff000088880fffffff080888880fffffff008888880000000008888888
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =1506
                    Top =1385
                    Width =453
                    Height =450
                    FontSize =8
                    TabIndex =2
                    Name ="CmdImprimer"
                    StatusBarText ="Imprimer le rapport"
                    OnClick ="=ImpVisu(0)"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x1000000010000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888000000000008888088888888808088 ,
                        0x00000000000008080888888bbb88000808888887778808080000000000000880 ,
                        0x08888888888080808000000000080800880ffffffff080808880f00000f00008 ,
                        0x8880ffffffff088888880f00000f088888880ffffffff0888888800000000088 ,
                        0x8888888888888888
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =73
                    TextFontFamily =34
                    Left =2190
                    Top =1380
                    Width =486
                    Height =456
                    FontSize =8
                    TabIndex =3
                    Name ="CmdExporter"
                    Caption ="&i"
                    StatusBarText ="Imprimer le rapport"
                    OnClick ="=ImpVisu(5)"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaad000000000000add0330880000330da ,
                        0xa0330880000330add0330880000330daa0330000000330add0333333333330da ,
                        0xa0330000000330add030fffffff030daa030fffcfff030add030ffcccff030da ,
                        0xa030fffcccfdadadd030ffffcccacadaa000000daccccdaddadadadadacccada ,
                        0xadadadadaccccdad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="PiedFormulaire"
        End
    End
End
CodeBehindForm
' See "FrmDebiteurNouveauChequeFiltreEdition.cls"
