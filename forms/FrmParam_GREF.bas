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
    Width =6615
    ItemSuffix =68
    Left =270
    Top =120
    Right =6885
    Bottom =5205
    RecSrcDt = Begin
        0xda7ea9f0d6e9e240
    End
    RecordSource ="PARAMETRE_GREF"
    Caption ="Paramétrage divers"
    OnOpen ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    Begin
        Begin Rectangle
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
        Begin OptionGroup
            Width =1701
            Height =1701
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
            Width =1701
            LabelX =-1701
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
        Begin ComboBox
            Width =1701
            LabelX =-1701
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
            Height =567
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =165
                    Top =45
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
            End
        End
        Begin Section
            SpecialEffect =2
            Height =4530
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =2
                    Left =2767
                    Top =1294
                    Width =1038
                    Height =255
                    TabIndex =5
                    Name ="TxtFrais_Déplacement_m1"
                    ControlSource ="Frais_Déplacement_m1"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    StatusBarText ="Pourcentage"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =0
                    OverlapFlags =93
                    Left =150
                    Top =255
                    Width =6285
                    Height =616
                    Name ="RecParam_GREF"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =270
                    Top =120
                    Width =1905
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="EtiParam_GREF"
                    Caption ="Paramètres Généraux"
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
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =285
                    Top =1290
                    Width =2415
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiFrais_Déplacement_m1"
                    Caption ="Frais de Déplacement       :"
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
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =285
                    Top =1620
                    Width =2415
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiAdresse_Courrier"
                    Caption ="Adresse pour le courrier    :"
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
                    OverlapFlags =215
                    TextAlign =2
                    Left =1208
                    Top =499
                    Width =1038
                    Height =255
                    TabIndex =3
                    Name ="TxtTaux_TVA"
                    ControlSource ="Taux_TVA"
                    Format ="Percent"
                    StatusBarText ="Pourcentage"

                End
                Begin Label
                    OverlapFlags =215
                    Left =450
                    Top =495
                    Width =690
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiTaux_TVA"
                    Caption ="TVA   :"
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
                    OverlapFlags =93
                    TextAlign =1
                    ColumnCount =2
                    Left =2774
                    Top =1620
                    Width =2259
                    Height =255
                    FontSize =4
                    TabIndex =6
                    Name ="CmbAdresse_Courrier"
                    ControlSource ="Adresse_Courrier"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCTROW ADRESSE_TYPE.* FROM ADRESSE_TYPE; "
                    ColumnWidths ="0;1134"

                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =0
                    OverlapFlags =255
                    Left =150
                    Top =1065
                    Width =6285
                    Height =1921
                    Name ="Boîte24"
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =270
                    Top =945
                    Width =1650
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="Texte25"
                    Caption ="Paramètres GREF"
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
                Begin OptionGroup
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =150
                    Top =3195
                    Width =6270
                    Height =1180
                    Name ="CadMode_Calcul_Frais"
                    ControlSource ="Mode_Calcul_Frais"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =270
                            Top =3075
                            Width =2220
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            ForeColor =16777215
                            Name ="EtiMode_Calcul_Frais"
                            Caption ="Mode de calcul des Frais"
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
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =570
                            Top =3390
                            Width =1110
                            Height =195
                            OptionValue =1
                            Name ="OptAncienMode"
                            OnKeyUp ="[Event Procedure]"
                            OnMouseUp ="[Event Procedure]"
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
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =247
                                    Left =855
                                    Top =3360
                                    Width =1260
                                    Height =240
                                    FontWeight =700
                                    BackColor =12632256
                                    Name ="Texte38"
                                    Caption ="Grille de Frais"
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
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =570
                            Top =3716
                            Width =2295
                            OptionValue =2
                            Name ="OptNouveauMode"
                            OnKeyUp ="[Event Procedure]"
                            OnMouseUp ="[Event Procedure]"
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
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =247
                                    Left =855
                                    Top =3686
                                    Width =2040
                                    Height =240
                                    FontWeight =700
                                    Name ="Texte40"
                                    Caption ="Intérêt légal"
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
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =570
                            Top =4025
                            Width =2160
                            OptionValue =3
                            Name ="OptFonctionClient"
                            OnKeyUp ="[Event Procedure]"
                            OnMouseUp ="[Event procedure]"
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
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =247
                                    Left =855
                                    Top =3995
                                    Width =2040
                                    Height =240
                                    FontWeight =700
                                    Name ="Texte50"
                                    Caption ="En fonction du Client"
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
                            End
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =223
                    TextAlign =2
                    Left =5108
                    Top =3645
                    Width =888
                    Height =255
                    TabIndex =1
                    Name ="TxtTx_Interet_legal"
                    ControlSource ="Tx_Interet_legal"
                    Format ="Percent"
                    StatusBarText ="Pourcentage"

                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =1
                    Left =3180
                    Top =3645
                    Width =1905
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiTx_Interet_legal"
                    Caption ="Taux d'intérêt légal  :"
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
                    OverlapFlags =223
                    TextAlign =2
                    Left =5122
                    Top =3990
                    Width =873
                    Height =255
                    TabIndex =2
                    Name ="TxtFrais_Accessoires"
                    ControlSource ="Frais_Accessoires"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    StatusBarText ="Pourcentage"

                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =1
                    Left =3180
                    Top =3990
                    Width =1875
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiFraisAccessoires"
                    Caption ="Frais Accessoires    :"
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
                Begin OptionGroup
                    TabStop = NotDefault
                    SpecialEffect =1
                    OverlapFlags =247
                    Left =3120
                    Top =3585
                    Width =2970
                    Height =720
                    TabIndex =4
                    Name ="Champ43"

                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =2
                    OverlapFlags =247
                    TextAlign =1
                    Left =2775
                    Top =2280
                    Width =1143
                    Height =255
                    TabIndex =7
                    Name ="TxtMot_de_passe"
                    ControlSource ="Mot_de_passe"
                    StatusBarText ="Pourcentage"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    InputMask ="Password"

                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =285
                    Top =2280
                    Width =1725
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiMot_de_passe"
                    Caption ="Mot de passe"
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
                    OverlapFlags =215
                    TextAlign =2
                    Left =3968
                    Top =499
                    Width =1038
                    Height =255
                    TabIndex =8
                    Name ="TxtTauxConvertion_m1m2"
                    ControlSource ="TauxConvertion_m1m2"
                    Format ="General Number"
                    StatusBarText ="Taux de convertion"

                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =3135
                    Top =495
                    Width =810
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiTauxConvertion_m1m2"
                    Caption ="1 Euro ="
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
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =5025
                    Top =495
                    Width =810
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="Étiquette59"
                    Caption ="Francs"
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
                    OverlapFlags =247
                    TextAlign =2
                    Left =4268
                    Top =1294
                    Width =1038
                    Height =255
                    TabIndex =9
                    Name ="TxtFrais_Déplacement_m2"
                    ControlSource ="Frais_Déplacement_m2"
                    Format ="#,##0.00 €;-#,##0.00 €"
                    StatusBarText ="Pourcentage"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =3825
                            Top =1290
                            Width =390
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiFrais_Déplacement_m2"
                            Caption ="soit"
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
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =247
                    TextAlign =1
                    Left =2775
                    Top =2610
                    Width =1143
                    Height =255
                    TabIndex =10
                    Name ="TxtCode_Sous_Traitant"
                    ControlSource ="Code_Sous_Traitant"
                    StatusBarText ="Pourcentage"

                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =285
                    Top =2610
                    Width =1725
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiCode_Sous_Traitant"
                    Caption ="Code BLITZ"
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
                    OverlapFlags =247
                    TextFontFamily =34
                    Left =5100
                    Top =1635
                    Width =260
                    Height =245
                    TabIndex =11
                    Name ="CmdTypeAdresse"
                    Caption ="..."
                    StatusBarText ="Ajouter un nouveau type d'adresse"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000000a000000040000000100040000000000200000000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888800000080080080080000008008008008000000 ,
                        0x8888888888000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
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
' See "FrmParam_GREF.cls"
