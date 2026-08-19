Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =7095
    ItemSuffix =51
    Left =3570
    Top =225
    Right =11040
    Bottom =4965
    RecSrcDt = Begin
        0xc1eb966ed0e9e240
    End
    RecordSource ="SELECT FRAIS.* FROM FRAIS WHERE (((FRAIS.Fra_Grille)=Formulaires!FrmParam_Frais!"
        "CmbFra_Grille)); "
    Caption ="Frais de Recouvrement"
    OnLoad ="[Event Procedure]"
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
        Begin Line
            Width =1701
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
        Begin Subform
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =750
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
                    Caption ="FermerSimple"
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
                        0x00000000000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    ColumnCount =2
                    ListRows =3
                    ListWidth =2041
                    Left =3060
                    Top =240
                    Width =2040
                    TabIndex =3
                    BoundColumn =1
                    Name ="CmbFra_Grille"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT FRAIS_GRILLE.Fra_Grille_Libellé, FRAIS_GRILLE.Fra_Grille FROM FRAIS_GRILL"
                        "E ORDER BY FRAIS_GRILLE.Fra_Grille; "
                    ColumnWidths ="1814;227"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2295
                            Top =240
                            Width =690
                            Height =240
                            Name ="EtiFra_Grille"
                            Caption ="Grille :"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =5205
                    Top =225
                    Width =260
                    Height =275
                    TabIndex =4
                    Name ="CmdFraisGrille"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000000a000000040000000100040000000000200000000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888800000080080080080000008008008008000000 ,
                        0x8888888888000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
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
                    Left =720
                    Top =120
                    Width =472
                    Height =472
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
                        0x0fffffff000088880fffffff080888880fffffff008888880000000008888888 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000
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
                    Left =1335
                    Top =120
                    Width =472
                    Height =472
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
                        0x8888888888888888000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
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
                    Left =5730
                    Top =180
                    Width =1256
                    Height =347
                    TabIndex =5
                    Name ="CmdInverseur"
                    Caption ="Inverseur"
                    OnClick ="[Event Procedure]"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =2850
            BackColor =12632256
            Name ="Détail0"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    Left =2310
                    Top =225
                    Width =936
                    Height =225
                    Name ="Fra_Tranche1"
                    ControlSource ="Fra_Tranche1_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =300
                    Top =225
                    Width =1935
                    Height =225
                    FontWeight =400
                    Name ="TexteT1"
                    Caption ="De 0 à 100 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =2310
                    Top =465
                    Width =936
                    Height =225
                    TabIndex =1
                    Name ="Fra_Tranche2"
                    ControlSource ="Fra_Tranche2_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =300
                    Top =465
                    Width =1935
                    Height =225
                    FontWeight =400
                    Name ="TexteT2"
                    Caption ="De 100.01 à 300 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =2310
                    Top =705
                    Width =936
                    Height =225
                    TabIndex =2
                    Name ="Fra_Tranche3"
                    ControlSource ="Fra_Tranche3_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =300
                    Top =705
                    Width =1935
                    Height =225
                    FontWeight =400
                    Name ="TexteT3"
                    Caption ="De 300.01 à 500 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =2310
                    Top =945
                    Width =936
                    Height =225
                    TabIndex =3
                    Name ="Fra_Tranche4"
                    ControlSource ="Fra_Tranche4_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =300
                    Top =945
                    Width =1935
                    Height =225
                    FontWeight =400
                    Name ="TexteT4"
                    Caption ="De 500.01 à 1000 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =2310
                    Top =1185
                    Width =936
                    Height =225
                    TabIndex =4
                    Name ="Fra_Tranche5"
                    ControlSource ="Fra_Tranche5_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =300
                    Top =1185
                    Width =1935
                    Height =225
                    FontWeight =400
                    Name ="TexteT5"
                    Caption ="De 1000.01 à 1500 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =2310
                    Top =1425
                    Width =936
                    Height =225
                    TabIndex =5
                    Name ="Fra_Tranche6"
                    ControlSource ="Fra_Tranche6_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =300
                    Top =1425
                    Width =1935
                    Height =225
                    FontWeight =400
                    Name ="TexteT6"
                    Caption ="De 1500.01 à 2000 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =2310
                    Top =1665
                    Width =936
                    Height =225
                    TabIndex =6
                    Name ="Fra_Tranche7"
                    ControlSource ="Fra_Tranche7_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =300
                    Top =1665
                    Width =1935
                    Height =225
                    FontWeight =400
                    Name ="TexteT7"
                    Caption ="De 2000.01 à 3000 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =2310
                    Top =1905
                    Width =936
                    Height =225
                    TabIndex =7
                    Name ="Fra_Tranche8"
                    ControlSource ="Fra_Tranche8_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =300
                    Top =1905
                    Width =1935
                    Height =225
                    FontWeight =400
                    Name ="TexteT8"
                    Caption ="De 3000.01 à 4000 F"
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =2
                    Left =2310
                    Top =2145
                    Width =936
                    Height =225
                    TabIndex =8
                    Name ="Fra_Tranche9"
                    ControlSource ="Fra_Tranche9_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =300
                    Top =2145
                    Width =1935
                    Height =225
                    FontWeight =400
                    Name ="TexteT9"
                    Caption ="De 4000.01 à 5000 F"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    Left =5693
                    Top =225
                    Width =936
                    Height =225
                    TabIndex =9
                    Name ="Fra_Tranche10"
                    ControlSource ="Fra_Tranche10_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3495
                    Top =225
                    Width =2130
                    Height =225
                    FontWeight =400
                    Name ="TexteT10"
                    Caption ="De 5000.01 à 7000 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =5693
                    Top =465
                    Width =936
                    Height =225
                    TabIndex =10
                    Name ="Fra_Tranche11"
                    ControlSource ="Fra_Tranche11_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3495
                    Top =465
                    Width =2130
                    Height =225
                    FontWeight =400
                    Name ="TexteT11"
                    Caption ="De 7000.01 à 10000 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =5693
                    Top =705
                    Width =936
                    Height =225
                    TabIndex =11
                    Name ="Fra_Tranche12"
                    ControlSource ="Fra_Tranche12_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3495
                    Top =705
                    Width =2130
                    Height =225
                    FontWeight =400
                    Name ="TexteT12"
                    Caption ="De 10000.01 à 15000 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =5693
                    Top =945
                    Width =936
                    Height =225
                    TabIndex =12
                    Name ="Fra_Tranche13"
                    ControlSource ="Fra_Tranche13_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3495
                    Top =945
                    Width =2130
                    Height =225
                    FontWeight =400
                    Name ="TexteT13"
                    Caption ="De 15000.01 à 20000 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =5693
                    Top =1185
                    Width =936
                    Height =225
                    TabIndex =13
                    Name ="Fra_Tranche14"
                    ControlSource ="Fra_Tranche14_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3495
                    Top =1185
                    Width =2130
                    Height =225
                    FontWeight =400
                    Name ="TexteT14"
                    Caption ="De 20000.01 à 30000 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =5693
                    Top =1425
                    Width =936
                    Height =225
                    TabIndex =14
                    Name ="Fra_Tranche15"
                    ControlSource ="Fra_Tranche15_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3495
                    Top =1425
                    Width =2130
                    Height =225
                    FontWeight =400
                    Name ="TexteT15"
                    Caption ="De 30000.01 à 40000 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =5693
                    Top =1665
                    Width =936
                    Height =225
                    TabIndex =15
                    Name ="Fra_Tranche16"
                    ControlSource ="Fra_Tranche16_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3495
                    Top =1665
                    Width =2130
                    Height =225
                    FontWeight =400
                    Name ="TexteT16"
                    Caption ="De 40000.01 à 50000 F"
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =5693
                    Top =1905
                    Width =936
                    Height =225
                    TabIndex =16
                    Name ="Fra_Tranche17"
                    ControlSource ="Fra_Tranche17_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3495
                    Top =1905
                    Width =2130
                    Height =225
                    FontWeight =400
                    Name ="TexteT17"
                    Caption ="De 50000.01 à 70000 F"
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =2
                    Left =5693
                    Top =2145
                    Width =936
                    Height =225
                    TabIndex =17
                    Name ="Fra_Tranche18"
                    ControlSource ="Fra_Tranche18_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3495
                    Top =2145
                    Width =2130
                    Height =225
                    FontWeight =400
                    Name ="TexteT18"
                    Caption ="Suppérieur à 70000 F"
                End
                Begin Line
                    OverlapFlags =85
                    Left =3375
                    Top =225
                    Width =0
                    Height =2175
                    Name ="Ligne80"
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="PiedFormulaire2"
        End
    End
End
CodeBehindForm
' See "FrmParam_Frais.cls"
