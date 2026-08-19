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
    Width =5340
    ItemSuffix =41
    Left =4650
    Top =1695
    Right =10365
    Bottom =6915
    RecSrcDt = Begin
        0xd9c84d75f3ece240
    End
    RecordSource ="SELECT DISTINCTROW MOTIF_IMPAYE.MI_Code, MOTIF_IMPAYE.MI_Libellé, MOTIF_IMPAYE.M"
        "I_Code_Gestion, MOTIF_IMPAYE.MI_Grille FROM MOTIF_IMPAYE WHERE (((MOTIF_IMPAYE.M"
        "I_Grille)=Formulaires!FrmParam_Motif_Impaye!CmbMIG_Grille)); "
    Caption ="Grille de Motifs Impayés Clients"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    Begin
        Begin Label
            TextAlign =3
            FontWeight =700
            BackColor =12632256
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Rectangle
            SpecialEffect =2
            Width =850
            Height =850
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin OptionButton
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            Width =187
            Height =187
            LabelX =-236
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CheckBox
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            Width =187
            Height =187
            LabelX =-236
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            SpecialEffect =2
            LabelAlign =3
            Width =1701
            Height =255
            LabelX =-236
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin ListBox
            SpecialEffect =2
            LabelAlign =3
            Width =1701
            Height =1417
            LabelX =-236
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin ComboBox
            SpecialEffect =2
            LabelAlign =3
            Width =1701
            Height =255
            LabelX =-236
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin FormHeader
            Height =907
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =113
                    Top =680
                    Width =534
                    Height =227
                    FontSize =7
                    Name ="EtiMI_Code"
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
                        0x8888888888888888000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =755
                    Top =675
                    Width =3165
                    Height =225
                    FontSize =7
                    Name ="EtiMI_Libellé"
                    Caption ="Motif impayé propre au Client"
                    FontName ="Small Fonts"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3975
                    Top =675
                    Width =1275
                    Height =225
                    FontSize =7
                    Name ="EtiMI_Code_Gestion"
                    Caption ="Situation GREF"
                    FontName ="Small Fonts"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =735
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
                        0x0000000000000000000000000000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =1350
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
                        0x8888888888888888000000000000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    ListRows =3
                    ListWidth =2041
                    Left =2850
                    Top =270
                    Width =2040
                    TabIndex =3
                    BoundColumn =1
                    Name ="CmbMIG_Grille"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MOTIF_IMPAYE_GRILLE.MIG_Grille_Libellé, MOTIF_IMPAYE_GRILLE.MIG_Grille FR"
                        "OM MOTIF_IMPAYE_GRILLE ORDER BY MOTIF_IMPAYE_GRILLE.MIG_Grille; "
                    ColumnWidths ="1814;229"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"01\""

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2085
                            Top =270
                            Width =690
                            Height =240
                            Name ="EtiMIG_Grille"
                            Caption ="Grille :"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4995
                    Top =255
                    Width =260
                    Height =275
                    TabIndex =4
                    Name ="CmdMotif_ImpayeGrille"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000000a000000040000000100040000000000200000000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888800000080080080080000008008008008000000 ,
                        0x88888888880000000000000000000000000000000000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

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
                    Name ="TxtMI_Code"
                    ControlSource ="MI_Code"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    Left =750
                    Width =3471
                    TabIndex =1
                    Name ="TxtMI_Libellé"
                    ControlSource ="MI_Libellé"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    ColumnCount =2
                    ListRows =5
                    ListWidth =2268
                    Left =4305
                    Width =722
                    TabIndex =2
                    Name ="CmbMI_Code_Gestion"
                    ControlSource ="MI_Code_Gestion"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CODE_GESTION.CG_Code, CODE_GESTION.CG_Libellé FROM CODE_GESTION ORDER BY "
                        "CODE_GESTION.CG_Code; "
                    ColumnWidths ="284;1985"
                    DefaultValue ="1"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =5040
                    Width =186
                    TabIndex =3
                    Name ="TxtMI_Grille"
                    ControlSource ="MI_Grille"
                    DefaultValue ="=[CmbMIG_Grille]"

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
' See "FrmParam_Motif_Impaye.cls"
