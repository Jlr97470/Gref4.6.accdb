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
    BorderStyle =0
    GridX =5
    GridY =5
    Width =6600
    ItemSuffix =52
    Left =4545
    Top =3030
    Right =11055
    Bottom =6870
    RecSrcDt = Begin
        0x8d8059fb93ece240
    End
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
            Height =232
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =60
                    Width =495
                    Height =210
                    FontSize =7
                    Name ="EtiCG_Code"
                    Caption ="Code"
                    FontName ="Small Fonts"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3644
                    Width =765
                    Height =225
                    FontSize =7
                    Name ="EtiNbDebiteur"
                    Caption ="Nombre"
                    FontName ="Small Fonts"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =555
                    Width =585
                    Height =210
                    FontSize =7
                    Name ="EtiCG_Libellé"
                    Caption ="Libellé"
                    FontName ="Small Fonts"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4477
                    Width =1260
                    Height =225
                    FontSize =7
                    Name ="EtiSolde"
                    Caption ="Solde"
                    FontName ="Small Fonts"
                End
            End
        End
        Begin Section
            Height =270
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =105
                    Top =30
                    Width =411
                    Height =225
                    FontSize =7
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtCG_Code"
                    ControlSource ="CG_Code"
                    FontName ="Small Fonts"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =607
                    Top =30
                    Width =2961
                    Height =225
                    FontSize =7
                    TabIndex =1
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtCG_Libellé"
                    ControlSource ="CG_Libellé"
                    FontName ="Small Fonts"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =3637
                    Top =30
                    Width =771
                    Height =225
                    FontSize =7
                    TabIndex =2
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtNbDebiteur"
                    ControlSource ="NbDebiteur"
                    Format ="Standard"
                    FontName ="Small Fonts"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =4470
                    Top =30
                    Width =1266
                    Height =225
                    FontSize =7
                    TabIndex =3
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtSolde"
                    ControlSource ="=[Déb_Mont_Prin_m2]-[Déb_Prin_Enc_m2]"
                    Format ="Standard"
                    FontName ="Small Fonts"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    Left =5790
                    Top =30
                    Width =771
                    Height =225
                    FontSize =7
                    TabIndex =4
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtPourcentage"
                    ControlSource ="=IIf([TxtTotal]=0,0,[TxtSolde]/[TxtTotal])"
                    Format ="Percent"
                    FontName ="Small Fonts"
                End
            End
        End
        Begin FormFooter
            Height =240
            BackColor =12632256
            Name ="PiedFormulaire"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =525
                    Width =1266
                    Height =225
                    FontSize =7
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtTotalSolde"
                    ControlSource ="=Sum([Déb_Mont_Prin_m2]-[Déb_Prin_Enc_m2])"
                    Format ="Standard"
                    FontName ="Small Fonts"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =1845
                    Width =1266
                    Height =225
                    FontSize =7
                    TabIndex =1
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtTotalDéb_Mont_Reversé_m"
                    ControlSource ="=Sum([Déb_Mont_Reversé_m2])+Round(Sum([Déb_Mont_Reversé_m1])*DonnerTauxConvertio"
                        "n(1),2)"
                    Format ="Standard"
                    FontName ="Small Fonts"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =3150
                    Width =1266
                    Height =225
                    FontSize =7
                    TabIndex =2
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtTotalDéb_Prin_Enc_m"
                    ControlSource ="=Sum([Déb_Prin_Enc_m2])"
                    Format ="Standard"
                    FontName ="Small Fonts"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =4440
                    Width =1266
                    Height =225
                    FontSize =7
                    TabIndex =3
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtTotal"
                    ControlSource ="=[TxtTotalSolde]+[TxtTotalDéb_Mont_Reversé_m]+([TxtTotalDéb_Prin_Enc_m]-[TxtTota"
                        "lDéb_Mont_Reversé_m])"
                    Format ="Standard"
                    FontName ="Small Fonts"
                End
            End
        End
    End
End
CodeBehindForm
' See "SFrmClientAnalyseSituation.cls"
