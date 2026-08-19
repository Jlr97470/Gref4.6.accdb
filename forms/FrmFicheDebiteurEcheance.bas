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
    Width =4762
    ItemSuffix =81
    Left =4980
    Top =15
    Right =10095
    Bottom =3675
    RecSrcDt = Begin
        0xd471ac8431e8e240
    End
    RecordSource ="SELECT DEBITEUR_ECHEANCE.* FROM DEBITEUR_ECHEANCE; "
    Caption ="Les echéances"
    BeforeInsert ="[Event Procedure]"
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
            Height =915
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =1431
                    Top =190
                    Width =2190
                    Height =405
                    FontSize =14
                    ForeColor =8355711
                    Name ="EtiAtermoiementsTitreOmbre"
                    Caption ="Echéances"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =1380
                    Top =150
                    Width =2190
                    Height =405
                    FontSize =14
                    ForeColor =16777215
                    Name ="EtiAtermoiementsTitre"
                    Caption ="Echéances"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =30
                    Top =120
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
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =0
                    Top =705
                    Width =915
                    Height =210
                    FontSize =7
                    Name ="EtiDe_Ope_Code"
                    Caption ="Opérateur :"
                    FontName ="Small Fonts"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =0
                    Left =900
                    Top =705
                    Width =1185
                    Height =210
                    FontSize =7
                    Name ="EtiDe_Date_éch"
                    Caption ="Date"
                    FontName ="Small Fonts"
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =615
                    Top =120
                    Width =472
                    Height =472
                    FontSize =8
                    TabIndex =1
                    Name ="CmdImprimer"
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
                        0x00000000
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
                    TextAlign =0
                    Left =2265
                    Top =705
                    Width =1185
                    Height =210
                    FontSize =7
                    Name ="EtiDe_Mont_éch_m"
                    Caption ="Montant"
                    FontName ="Small Fonts"
                End
            End
        End
        Begin Section
            Height =345
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    Left =2265
                    Top =30
                    Width =1416
                    TabIndex =2
                    Name ="TxtDe_Mont_éch_m"
                    ControlSource ="De_Mont_éch_m1"
                    Format ="Standard"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =915
                    Top =30
                    Width =1311
                    TabIndex =1
                    Name ="TxtDe_Date_éch"
                    ControlSource ="De_Date_éch"
                    Format ="Short Date"
                    BeforeUpdate ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3945
                    Top =30
                    Width =786
                    Height =225
                    TabIndex =4
                    Name ="TxtDe_Mont_éch_mo"
                    ControlSource ="De_Mont_éch_mo"
                    StatusBarText ="Atermoiement Montant Monnaie"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3232
                    Left =30
                    Top =30
                    Width =846
                    Name ="CmbDe_Ope_Code"
                    ControlSource ="De_Ope_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT OPERATEUR.Ope_Code, OPERATEUR.Ope_Nom FROM OPERATEUR ORDER BY OPERATEUR.O"
                        "pe_Code; "
                    ColumnWidths ="398;2835"
                    StatusBarText ="Debiteur Echeance Operateur Code"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3705
                    Top =60
                    Width =1037
                    Height =226
                    TabIndex =3
                    ForeColor =8388608
                    Name ="TxtMon_Nom"
                    ControlSource ="=DLookUp(\"Mon_Nom\",\"MONNAIE\",\"Mon_Id=\" & [TxtDe_Mont_éch_mo])"
                End
            End
        End
        Begin FormFooter
            Height =1245
            BackColor =12632256
            Name ="PiedFormulaire"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    Left =2330
                    Top =226
                    Width =1386
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtTotal"
                    ControlSource ="=Round(Nz(Sum([De_Mont_éch_m1])),2)"
                    Format ="Standard"
                    DefaultValue ="0"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =525
                    Top =225
                    Width =396
                    TabIndex =1
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtNb"
                    ControlSource ="=Count([De_Date_éch])"
                End
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =255
                    Left =30
                    Top =90
                    Width =4500
                    Height =495
                    Name ="RecAtermoiements"
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =0
                    Left =90
                    Top =240
                    Width =390
                    Height =210
                    FontSize =7
                    Name ="EtiNb"
                    Caption ="Nb :"
                    FontName ="Small Fonts"
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =0
                    Left =1700
                    Top =241
                    Width =585
                    Height =210
                    FontSize =7
                    Name ="EtiTotal"
                    Caption ="Total :"
                    FontName ="Small Fonts"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =1814
                    Top =680
                    Width =1125
                    Height =450
                    FontWeight =300
                    TabIndex =2
                    Name ="CmdOK"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "FrmFicheDebiteurEcheance.cls"
