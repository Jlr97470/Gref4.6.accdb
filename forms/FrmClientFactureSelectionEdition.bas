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
    Width =3405
    ItemSuffix =23
    Left =5040
    Top =90
    Right =8790
    Bottom =2970
    RecSrcDt = Begin
        0xe5674b4d0de7e240
    End
    Caption ="Edition des factures Clients"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
        End
        Begin TextBox
            Width =1701
            LabelX =-1701
        End
        Begin ComboBox
            Width =1701
            LabelX =-1701
        End
        Begin FormHeader
            Height =15
            BackColor =12632256
            Name ="EntêteFormulaire0"
        End
        Begin Section
            Height =1133
            BackColor =12632256
            Name ="Détail0"
            Begin
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1695
                    Top =210
                    Width =1146
                    Name ="TxtHonoraire"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Top =210
                            Width =1695
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiHonoraire"
                            Caption ="Honoraire          :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1695
                    Top =720
                    Width =1146
                    TabIndex =1
                    Name ="TxtFactureNumero"
                    Format ="00000000"
                    DefaultValue ="=nz(DMax(\"Fac_Num\",\"FACTURE\",\"Fac_Num Like '\" & Format(Right(Year(Date()),"
                        "2),\"00\") & Format(Month(Date()),\"00\") & \"*'\"),Format(Right(Year(Date()),2)"
                        ",\"00\") & Format(Month(Date()),\"00\") & \"0000\")+1"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Top =720
                            Width =1695
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiFactureNumero"
                            Caption ="N° Facture         :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =566
            BackColor =12632256
            Name ="PiedFormulaire1"
            Begin
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =570
                    Top =60
                    Width =2295
                    Height =450
                    FontWeight =300
                    Name ="CmdOk"
                    Caption ="Ok"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "FrmClientFactureSelectionEdition.cls"
