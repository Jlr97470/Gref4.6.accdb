Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =0
    GridX =5
    GridY =5
    Width =3060
    ItemSuffix =41
    Left =7725
    Top =2160
    Right =10785
    Bottom =4380
    RecSrcDt = Begin
        0xe978f3d6d5e9e240
    End
    RecordSource ="SELECT DISTINCTROW FRAIS_TRANS.* FROM FRAIS_TRANS WHERE (((FRAIS_TRANS.Frt_Début"
        "_m1)>1500)) ORDER BY FRAIS_TRANS.Frt_Début_m1; "
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
            Height =0
            BackColor =12632256
            Name ="EntêteFormulaire1"
        End
        Begin Section
            Height =240
            BackColor =12632256
            Name ="Détail0"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    Left =113
                    Width =2031
                    Height =240
                    TabIndex =2
                    BackColor =8388608
                    Name ="Tranche"
                    ControlSource ="=IIf([Frt_Fin_m1]=0,\"Suppérieur à \" & [Frt_Début_m1],\"De \" & [Frt_Début_m1] "
                        "& \" à \" & [Frt_Fin_m1]) & \" F\""
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =2
                    Left =2154
                    Width =906
                    Height =240
                    Name ="Frt_Frais"
                    ControlSource ="Frt_Frais_m1"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =113
                    Width =111
                    Height =210
                    TabIndex =1
                    Name ="Frt_Grille"
                    ControlSource ="Frt_Grille"
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
' See "SFrmParam_FraisTrans2.cls"
