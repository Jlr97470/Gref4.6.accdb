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
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =6960
    ItemSuffix =53
    Left =4080
    Top =120
    Right =11400
    Bottom =6405
    RecSrcDt = Begin
        0xc7cc04764dece240
    End
    RecordSource ="SELECT TAUX.*, TAUX.Tx_Grille FROM TAUX WHERE (((TAUX.Tx_Grille)=Formulaires!Frm"
        "Param_Taux!CmbTxG_Grille)) ORDER BY TAUX.Tx_N°; "
    Caption ="Taux de Relance"
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
        Begin Subform
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =1200
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
                    Name ="CmdFemer"
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
                        0x0000000000000000
                    End
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
                    Left =1290
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
                        0x8888888888888888000000000000000000000000000000000000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Top =942
                            Width =450
                            Height =210
                            FontSize =7
                            Name ="EtiRel_Nom"
                            Caption ="Nom"
                            FontName ="Small Fonts"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    ColumnCount =2
                    ListRows =3
                    ListWidth =2041
                    Left =4410
                    Top =195
                    Width =2040
                    TabIndex =3
                    BoundColumn =1
                    Name ="CmbTxG_Grille"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TAUX_GRILLE.TxG_Grille_Libellé, TAUX_GRILLE.TxG_Grille FROM TAUX_GRILLE O"
                        "RDER BY TAUX_GRILLE.TxG_Grille; "
                    ColumnWidths ="1814;228"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3645
                            Top =195
                            Width =690
                            Height =240
                            Name ="EtiTxG_Grille"
                            Caption ="Grille :"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6555
                    Top =180
                    Width =260
                    Height =275
                    TabIndex =4
                    Name ="CmdTauxGrille"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000000a000000040000000100040000000000200000000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888800000080080080080000008008008008000000 ,
                        0x8888888888000000000000000000000000000000000000000000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4875
                    Top =922
                    Width =360
                    Height =210
                    FontSize =7
                    Name ="Texte32"
                    Caption ="N°"
                    FontName ="Small Fonts"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6180
                    Top =922
                    Width =690
                    Height =210
                    FontSize =7
                    Name ="EtiTx_Délai"
                    Caption ="Délai (j)"
                    FontName ="Small Fonts"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5325
                    Top =915
                    Width =795
                    Height =210
                    FontSize =7
                    Name ="EtiTx_Frais"
                    Caption ="Frais (%)"
                    FontName ="Small Fonts"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3105
                    Top =930
                    Width =1710
                    Height =210
                    FontSize =7
                    Name ="EtiTx_Code_Action"
                    Caption ="Action:"
                    FontName ="Small Fonts"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =255
            BackColor =12632256
            Name ="Détail0"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Width =3036
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtRel_Nom"
                    ControlSource ="=dlookup(\"Rel_Nom\",\"RELANCE\",\"Rel_Code_Action=\" & [Tx_Code_Action] & \" AN"
                        "D Rel_N°=\" & [Tx_N°])"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =4890
                    Width =336
                    TabIndex =1
                    Name ="Tx_N°"
                    ControlSource ="Tx_N°"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =6165
                    Width =696
                    TabIndex =2
                    Name ="TxtTx_Délai"
                    ControlSource ="Tx_Délai"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =5340
                    Width =711
                    TabIndex =3
                    Name ="TxtTx_Frais"
                    ControlSource ="Tx_Frais"
                End
                Begin ComboBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3105
                    ColumnWidth =1605
                    TabIndex =4
                    Name ="CmbTx_Code_Action"
                    ControlSource ="Tx_Code_Action"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CODE_ACTION.CA_Code, CODE_ACTION.CA_Libellé FROM CODE_ACTION ORDER BY COD"
                        "E_ACTION.CA_Libellé; "
                    ColumnWidths ="0;4536"
                    StatusBarText ="Grille Taux Code Action"
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
' See "FrmParam_Taux.cls"
