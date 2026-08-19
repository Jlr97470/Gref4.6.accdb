Version =20
VersionRequired =20
Begin Form
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =7935
    ItemSuffix =217
    Left =3360
    Top =30
    Right =11640
    Bottom =8295
    TimerInterval =1000
    RecSrcDt = Begin
        0x8d7a8876b4ece240
    End
    RecordSource ="SELECT TMP_REPRESENTANT_CLIENT.* FROM TMP_REPRESENTANT_CLIENT WHERE (((TMP_REPRE"
        "SENTANT_CLIENT.TRCRep_Code)=Formulaires!FrmClientSuiviSelection!CmbTRCRep_Code))"
        " ORDER BY TMP_REPRESENTANT_CLIENT.TRCRep_Code, TMP_REPRESENTANT_CLIENT.TRCCli_Co"
        "de; "
    Caption ="Suivi des Clients"
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
            Height =2310
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
                        0x00000000000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =1620
                    Top =735
                    Width =5310
                    Height =405
                    FontSize =14
                    ForeColor =8355711
                    Name ="EtiFrmClientSelectionTitreOmbre"
                    Caption ="Selection des clients"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    Left =1575
                    Top =690
                    Width =5325
                    Height =405
                    FontSize =14
                    ForeColor =16777215
                    Name ="EtiFrmClientSuiviSelectionTitre"
                    Caption ="Selection des clients"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =90
                    Top =2070
                    Width =1065
                    Height =240
                    Name ="EtiTRCCli_Code"
                    Caption ="Code"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1200
                    Top =2070
                    Width =2715
                    Height =240
                    Name ="EtiCli_Nom"
                    Caption ="Nom"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    Left =1545
                    Top =1530
                    Width =3696
                    TabIndex =1
                    Name ="CmbTRCRep_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT REPRESENTANT.Rep_Code, REPRESENTANT.Rep_Nom FROM REPRESENTANT ORDER BY RE"
                        "PRESENTANT.Rep_Nom; "
                    ColumnWidths ="0;2268"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =105
                            Top =1530
                            Width =1335
                            Height =240
                            FontSize =7
                            BackColor =8421504
                            Name ="EtiTRCRep_Code"
                            Caption ="Représentant :"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =73
                    TextFontFamily =34
                    Left =735
                    Top =120
                    Width =472
                    Height =472
                    FontSize =8
                    TabIndex =2
                    Name ="CmdImprimer"
                    Caption ="&i"
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
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =86
                    TextFontFamily =34
                    Left =1365
                    Top =120
                    Width =472
                    Height =472
                    FontSize =8
                    TabIndex =3
                    Name ="CmdVisualiser"
                    Caption ="&v"
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
                    AccessKey =73
                    TextFontFamily =34
                    Left =2010
                    Top =120
                    Width =486
                    Height =471
                    FontSize =8
                    TabIndex =4
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
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =80
                    TextFontFamily =34
                    Left =7365
                    Top =105
                    Width =471
                    Height =471
                    FontSize =7
                    TabIndex =5
                    Name ="CmdSupprimer"
                    Caption ="&p"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000019000000180000000100040000000000800100000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888888888888888888888888888888887000000000 ,
                        0x00788888888888888888880fbbbbb3b333088888888888888888880fbbbb3b33 ,
                        0x33088888888888888888880fb0bb03b033088888888888888888880fb0bb0b30 ,
                        0x33088888888888888888880fb0bb03b033088888888888888888880fb0bb0b30 ,
                        0x33088888888888808888880fb0bb03b033088888888888088888880fb0bb0b30 ,
                        0x330888888888808f8888880fb0bb03b033088888888887008888880fb0bb0b30 ,
                        0x33088888888888888888880fb0bb03b033088888888888888888880fb0bb0b30 ,
                        0x33088888888888888888880fbbbbb3b333088888888888888888880fbbbb3b33 ,
                        0x330888888888888888888700000000000007888888888888888880bbbbbb3b33 ,
                        0x3330888888888888888880bbbbb3b3b333308888888888808888870000000000 ,
                        0x0007888888888880888888888700000788888888888888888888888888888888 ,
                        0x8888888888888888
                    End
                    FontName ="Small Fonts"
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
                    Left =6000
                    Top =225
                    Width =720
                    Height =270
                    Name ="EtiRien"
                    Caption ="Rien"
                    OnClick ="=ToutRien(False)"
                End
                Begin Label
                    SpecialEffect =1
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =5250
                    Top =225
                    Width =720
                    Height =270
                    Name ="EtiTout"
                    Caption ="Tout"
                    OnClick ="=ToutRien(True)"
                End
                Begin Rectangle
                    SpecialEffect =1
                    BackStyle =0
                    OverlapFlags =215
                    Left =5145
                    Top =120
                    Width =1665
                    Height =450
                    Name ="RecToutRien"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3975
                    Top =2070
                    Width =1245
                    Height =240
                    Name ="EtiCli_Date_contrat"
                    Caption ="Date Contrat"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =285
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =90
                    Top =15
                    Width =1086
                    ColumnWidth =1815
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtTRCCli_Code"
                    ControlSource ="TRCCli_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    Left =1185
                    Top =15
                    Width =2751
                    TabIndex =1
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtCli_Nom"
                    ControlSource ="=DLookUp(\"Cli_Nom\",\"CLIENT\",\"Cli_Code='\" & [TRCCli_Code] & \"'\")"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =7695
                    Top =60
                    TabIndex =2
                    Name ="ChkTRC_Flag"
                    ControlSource ="TRC_Flag"
                    StatusBarText ="Representant Client Selectionner"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =3960
                    Top =15
                    Width =1281
                    TabIndex =3
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtCli_Date_contrat"
                    ControlSource ="=DLookUp(\"Cli_Date_contrat\",\"CLIENT\",\"Cli_Code='\" & [TRCCli_Code] & \"'\")"
                End
            End
        End
        Begin FormFooter
            Height =585
            BackColor =12632256
            Name ="PiedFormulaire"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =3120
                    Top =75
                    Width =2266
                    Height =405
                    Name ="CmdClientListe"
                    Caption ="&Sélection Client"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "FrmClientSuiviSelection.cls"
