Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DateGrouping =1
    GridY =10
    Width =9070
    ItemSuffix =39
    Left =270
    Top =2160
    RecSrcDt = Begin
        0xb31b41437524e440
    End
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    Begin
        Begin Label
            OldBorderStyle =1
            FontWeight =700
            FontName ="Arial"
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
            BackStyle =0
            Width =850
            Height =280
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
        Begin Line
            Width =1701
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
        Begin BoundObjectFrame
            AutoLabel = NotDefault
            Width =1998
            Height =280
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
        Begin TextBox
            AutoLabel = NotDefault
            Width =1701
            Height =280
            LabelX =-1701
            FontName ="Arial"
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
            ShowDatePicker =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="=Date()"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="=[Rel_Code_Action]<2"
        End
        Begin BreakLevel
            ControlSource ="Rel_Code_Action"
        End
        Begin BreakLevel
            ControlSource ="Rel_N°"
        End
        Begin PageHeader
            Height =2280
            Name ="EntêtePage0"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Top =56
                    Width =6228
                    Height =480
                    FontSize =14
                    Name ="EtiAnalyseSituationDebiteur"
                    Caption ="Analyse de la situation des Débiteurs"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    Left =6360
                    Top =105
                    Width =2655
                    Height =285
                    FontSize =10
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Long Date"

                End
                Begin Line
                    BorderWidth =2
                    Width =9018
                    Name ="TrtAnalyseSituationDebiteur"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="EntêteGroupe1"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =780
            BreakLevel =1
            Name ="EntêteGroupe0"
            Begin
                Begin Label
                    Left =1695
                    Top =495
                    Width =3900
                    Height =285
                    FontSize =10
                    Name ="EtiRel_Nom"
                    Caption ="Nom"
                End
                Begin Label
                    TextAlign =2
                    Left =5595
                    Top =495
                    Width =645
                    Height =285
                    FontSize =10
                    Name ="EtiRel_N°"
                    Caption ="N°"
                End
                Begin Label
                    TextAlign =2
                    Left =6240
                    Top =495
                    Width =1290
                    Height =285
                    FontSize =10
                    Name ="EtiNbDebiteur"
                    Caption ="Nombre"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =295
            Name ="Détail1"
            Begin
                Begin TextBox
                    Left =1695
                    Width =3900
                    Height =295
                    FontSize =10
                    Name ="TxtRel_Nom"
                    ControlSource ="Rel_Nom"

                End
                Begin TextBox
                    TextAlign =2
                    Left =5595
                    Width =645
                    Height =295
                    FontSize =10
                    TabIndex =1
                    Name ="TxtRel_N°"
                    ControlSource ="Rel_N°"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    Left =6240
                    Width =1290
                    Height =295
                    FontSize =10
                    TabIndex =2
                    Name ="TxtNbDebiteur"
                    ControlSource ="NbDebiteur"
                    Format ="Standard"
                    StatusBarText ="Type ou Taille du champ"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =295
            BreakLevel =1
            Name ="PiedGroupe1"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    Left =6240
                    Width =1290
                    Height =295
                    FontSize =10
                    FontWeight =700
                    Name ="TxtTotalNbDebiteur"
                    ControlSource ="=Sum([NbDebiteur])"
                    Format ="Standard"
                    StatusBarText ="Type ou Taille du champ"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =585
            Name ="PiedGroupe0"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6240
                    Top =150
                    Width =1281
                    Height =255
                    FontSize =10
                    FontWeight =700
                    Name ="TxtNbDebiteurEcheance"
                    ControlSource ="=DCount(\"De_Ope_Code\",\"DEBITEUR_ECHEANCE\",\"De_Date_éch<=Date()\")"

                    Begin
                        Begin Label
                            Left =1695
                            Top =165
                            Width =1800
                            Height =285
                            FontSize =10
                            Name ="EtiNbDebiteurEcheance"
                            Caption ="Nb Echeance CB :"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Height =420
            Name ="PiedPage2"
        End
    End
End
