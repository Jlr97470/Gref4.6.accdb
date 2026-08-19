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
    Width =9078
    ItemSuffix =213
    Left =2640
    Top =150
    Right =12060
    Bottom =4905
    TimerInterval =1000
    Filter ="Déb_MonnaieEnCours = 2"
    RecSrcDt = Begin
        0x011a6a3235e8e240
    End
    RecordSource ="SELECT TMP_RELANCE.*, DLookUp(\"Déb_MonnaieEnCours\",\"DEBITEUR\",\"Déb_Code='\""
        " & [TRel_Déb_Code] & \"'\") AS Déb_MonnaieEnCours FROM TMP_RELANCE ORDER BY TMP_"
        "RELANCE.TRel_Niv_relance, TMP_RELANCE.TRel_Déb_Code; "
    Caption ="Edition des Relances"
    OnOpen ="[Event Procedure]"
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
            Height =1247
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
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =1844
                    Top =158
                    Width =3900
                    Height =405
                    FontSize =14
                    ForeColor =8355711
                    Name ="EtiRelanceTitreOmbre"
                    Caption ="Liste des Relances à Editer"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =1814
                    Top =113
                    Width =3900
                    Height =405
                    FontSize =14
                    ForeColor =16777215
                    Name ="EtiRelanceTitre"
                    Caption ="Liste des Relances à Editer"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =570
                    Top =1005
                    Width =1125
                    Height =240
                    Name ="EtiTRel_Déb_Code"
                    Caption ="Débiteur"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1740
                    Top =1005
                    Width =3870
                    Height =240
                    Name ="EtiDéb_Nom"
                    Caption ="Nom"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =7110
                    Top =1005
                    Width =1395
                    Height =240
                    Name ="EtiDéb_Frais_m"
                    Caption ="Frais"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5670
                    Top =1005
                    Width =1380
                    Height =240
                    Name ="EtiResteAverser_m"
                    Caption ="Reste à verser"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =73
                    TextFontFamily =34
                    Left =690
                    Top =120
                    Width =472
                    Height =472
                    FontSize =8
                    TabIndex =1
                    Name ="CmdImprimer"
                    Caption ="&i"
                    StatusBarText ="Imprimer le rapport"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x1000000010000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888000000000008888088888888808088 ,
                        0x00000000000008080888888bbb88000808888887778808080000000000000880 ,
                        0x08888888888080808000000000080800880ffffffff080808880f00000f00008 ,
                        0x8880ffffffff088888880f00000f088888880ffffffff0888888800000000088 ,
                        0x8888888888888888000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000
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
                    AccessKey =80
                    TextFontFamily =34
                    Left =8055
                    Top =120
                    Width =471
                    Height =471
                    FontSize =7
                    TabIndex =2
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
                        0x8888888888888888000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000
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
                    Left =6884
                    Top =233
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
                    Left =6134
                    Top =233
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
                    Left =6029
                    Top =128
                    Width =1665
                    Height =450
                    Name ="RecToutRien"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    ListRows =2
                    Left =7110
                    Top =682
                    Width =1401
                    TabIndex =3
                    Name ="CmbMonnaie"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Monnaie.Mon_Id, Monnaie.Mon_Nom FROM MONNAIE; "
                    ColumnWidths ="0;1134"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =6030
                            Top =682
                            Width =1020
                            Height =240
                            FontSize =7
                            BackColor =8421504
                            Name ="EtiMonnaie"
                            Caption ="Monnaie :"
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =307
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =570
                    Top =15
                    Width =1131
                    ColumnWidth =1815
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtTRel_Déb_Code"
                    ControlSource ="TRel_Déb_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =1755
                    Top =15
                    Width =3861
                    TabIndex =1
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Nom"
                    ControlSource ="=Dlookup(\"Déb_Nom\",\"DEBITEUR\",\"Déb_Code='\" & [TRel_Déb_Code] & \"'\")"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =5670
                    Top =15
                    Width =1371
                    ColumnWidth =840
                    TabIndex =2
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtResteAverser_m"
                    ControlSource ="=Round(DLookUp(\"Déb_Mont_Prin_m1\",\"DEBITEUR\",\"Déb_Code='\" & [TRel_Déb_Code"
                        "] & \"'\"),2)-Round(nz(DSum(\"Enc_Prin_m1\",\"ENCAISSEMENT\",\"Enc_Code_Débiteur"
                        "='\" & [TRel_Déb_Code] & \"'\")),2)"
                    Format ="Standard"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7095
                    Top =15
                    Width =1416
                    ColumnWidth =825
                    TabIndex =3
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Frais_m"
                    ControlSource ="=DLookUp(\"Déb_Frais_m1\",\"DEBITEUR\",\"Déb_Code='\" & [TRel_Déb_Code] & \"'\")"
                    Format ="Standard"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =75
                    Top =15
                    Width =426
                    TabIndex =4
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtTRel_Niv_relance"
                    ControlSource ="TRel_Niv_relance"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8610
                    Top =60
                    TabIndex =5
                    Name ="ChkTRel_Editer"
                    ControlSource ="TRel_Editer"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="Yes"
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    BorderWidth =3
                    Left =8835
                    Top =60
                    TabIndex =6
                    Name ="ChkTRel_Deja_Editer"
                    ControlSource ="TRel_Deja_Editer"
                    StatusBarText ="Relance Deja Editer"
                End
            End
        End
        Begin FormFooter
            Height =225
            BackColor =12632256
            Name ="PiedFormulaire"
        End
    End
End
CodeBehindForm
' See "FrmDebiteurRelanceSelection.cls"
