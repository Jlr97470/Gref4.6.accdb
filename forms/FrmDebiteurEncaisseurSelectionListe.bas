Version =20
VersionRequired =20
Begin Form
    AllowEditing = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =7596
    ItemSuffix =78
    Left =3630
    Top =90
    Right =11580
    Bottom =5025
    RecSrcDt = Begin
        0xda896aba2ce8e240
    End
    RecordSource ="SELECT DEBITEUR.*, ADRESSE_DEBITEUR.AD_CodePostal, ADRESSE_DEBITEUR.AD_Ville, CO"
        "DE_GESTION.CG_Code_Action FROM (CODE_GESTION INNER JOIN DEBITEUR ON CODE_GESTION"
        ".CG_Code=DEBITEUR.Déb_Situation) INNER JOIN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Cod"
        "e=ADRESSE_DEBITEUR.AD_Code_Débiteur WHERE (((Len([Déb_Code_Enc]))=2) AND ((CODE_"
        "GESTION.CG_Code_Action)=3)); "
    Caption ="Edition des fiches encaisseurs"
    OnOpen ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
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
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =187
            Height =187
        End
        Begin CheckBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =187
            Height =187
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =1701
            Height =255
            LabelX =-1701
        End
        Begin ListBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin ComboBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =1701
            Height =255
            LabelX =-1701
        End
        Begin FormHeader
            Height =2026
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    Top =1785
                    Width =1140
                    Height =240
                    Name ="EtiDéb_Code"
                    Caption ="Code"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =95
                    Left =1140
                    Top =1785
                    Width =3600
                    Height =240
                    Name ="EtiDéb_Nom"
                    Caption ="Nom"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =113
                    Top =113
                    Width =471
                    Height =456
                    TabIndex =8
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
                        0x000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =95
                            Left =4740
                            Top =1785
                            Width =1155
                            Height =240
                            Name ="EtiDéb_Date_der_rel"
                            Caption ="Date"
                        End
                    End
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =120
                    Top =645
                    Width =455
                    Height =455
                    FontSize =7
                    FontWeight =400
                    TabIndex =9
                    Name ="CmdImprimer"
                    Caption ="Imprimer"
                    OnClick ="=ImpVisu(0)"
                    PictureData = Begin
                        0x2800000019000000180000000100040000000000800100000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888888888888888888888000008888000000000000 ,
                        0x000088888880ffff8880888888888888888008888880ffff8880888888888888 ,
                        0x888080888880ffff8000000000000000000008088880ff00808888888888bbbb ,
                        0x888800808880ffff808888888888bbbb888808008880ffff8088888888887777 ,
                        0x888808808880ff008088888888888888888808880880ffff8000000000000000 ,
                        0x000008880880ffff8088888888888888888080880880ffff8808888888888888 ,
                        0x888088080880ffff8880000000000000000008800880ffff888880ffffffffff ,
                        0xff0880880880fff0888880ffffffffffff0088070880fff08888880ff0000000 ,
                        0xfff000008880ffff8888880ffffffffffff008888880000088888880ffffffff ,
                        0xffff088888888888888888880f00000000ff088888888888888888880fffffff ,
                        0xfffff088888888888888888880fffffffffff088888888888888888880000000 ,
                        0x0000008888888888000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000
                    End
                    FontName ="Small Fonts"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    Left =5895
                    Top =1785
                    Width =435
                    Height =240
                    Name ="EtiDéb_Code_Enc"
                    Caption ="Enc."
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =690
                    Top =15
                    Width =6801
                    Height =1730
                    BackColor =8421504
                    Name ="RecRech"
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =2546
                    Top =317
                    Width =1131
                    Height =240
                    Name ="TxtDéb_Code"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =1237
                            Top =317
                            Width =1230
                            Height =240
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiRech_Déb_Code"
                            Caption ="Code  :"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =1298
                    Top =585
                    Width =2376
                    Height =240
                    TabIndex =1
                    Name ="TxtDéb_Nom"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =720
                            Top =585
                            Width =555
                            Height =240
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiRech_Déb_Nom"
                            Caption ="Nom :"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =223
                    Left =3720
                    Top =105
                    Width =525
                    Height =195
                    FontWeight =400
                    BackColor =8421504
                    ForeColor =65535
                    Name ="EtiClass"
                    Caption ="Class."
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =2537
                    Top =855
                    Width =1131
                    Height =240
                    TabIndex =2
                    Name ="TxtDéb_Date_der_relDebut"
                    Format ="Short Date"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =1170
                            Top =855
                            Width =1290
                            Height =240
                            FontSize =7
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiDéb_Date_der_relDebut"
                            Caption ="Entre le :"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =223
                    TextAlign =2
                    Left =6617
                    Top =225
                    Width =397
                    Height =240
                    TabIndex =4
                    Name ="TxtDéb_Code_Enc"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =5190
                            Top =225
                            Width =1365
                            Height =240
                            FontSize =7
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiRech_Déb_Code_Enc"
                            Caption ="Code Encaisseur :"
                        End
                    End
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =223
                    AccessKey =82
                    Left =4988
                    Top =1360
                    Width =2055
                    Height =315
                    TabIndex =11
                    Name ="CmdOK"
                    Caption ="&Rechercher"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =5882
                    Top =764
                    Width =1131
                    Height =240
                    TabIndex =5
                    Name ="TxtAD_CodePostal"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =4515
                            Top =764
                            Width =1290
                            Height =240
                            FontSize =7
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiAD_Code_Postal"
                            Caption ="Code Postal :"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =4643
                    Top =1034
                    Width =2376
                    Height =240
                    TabIndex =6
                    Name ="TxtAD_Ville"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =4065
                            Top =1034
                            Width =555
                            Height =240
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiAD_Ville"
                            Caption ="Ville :"
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =255
                    Left =690
                    Top =180
                    Width =6797
                    Height =1550
                    TabIndex =7
                    Name ="CadClass"
                    DefaultValue ="1"
                    Begin
                        Begin OptionButton
                            OverlapFlags =247
                            Left =7170
                            Top =270
                            OptionValue =4
                            Name ="OptDéb_Code_Enc"
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =3840
                            Top =360
                            OptionValue =1
                            Name ="OptDéb_Code"
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =3840
                            Top =645
                            OptionValue =2
                            Name ="OptDéb_Nom"
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =7171
                            Top =825
                            OptionValue =5
                            Name ="OptAD_CodePostal"
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =3840
                            Top =915
                            OptionValue =3
                            Name ="OptDéb_Date_der_rel"
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =7171
                            Top =1095
                            OptionValue =6
                            Name ="OptAD_Ville"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =2537
                    Top =1125
                    Width =1131
                    TabIndex =3
                    Name ="TxtDéb_Date_der_relFin"
                    Format ="Short Date"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =1170
                            Top =1125
                            Width =1290
                            Height =240
                            FontSize =7
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiDéb_Date_der_relFin"
                            Caption ="Et le :"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =86
                    TextFontFamily =34
                    Left =105
                    Top =1185
                    Width =472
                    Height =472
                    FontSize =8
                    TabIndex =10
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
            End
        End
        Begin Section
            Height =540
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =30
                    Top =15
                    Width =1101
                    Height =240
                    Name ="TxtDéb_Code1"
                    ControlSource ="Déb_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =4755
                    Top =15
                    Width =1131
                    Height =240
                    TabIndex =1
                    Name ="TxtDéb_Date_der_rel"
                    ControlSource ="Déb_Date_der_rel"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =1155
                    Top =15
                    Width =3576
                    Height =240
                    TabIndex =2
                    Name ="TxtDéb_Nom1"
                    ControlSource ="Déb_Nom"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =5915
                    Top =15
                    Width =397
                    Height =240
                    TabIndex =3
                    Name ="TxtDéb_Code_Enc1"
                    ControlSource ="Déb_Code_Enc"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =1155
                    Top =285
                    Width =1011
                    Height =240
                    TabIndex =4
                    Name ="TxtAD_CodePostal1"
                    ControlSource ="AD_CodePostal"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =2190
                    Top =285
                    Width =3696
                    Height =240
                    TabIndex =5
                    Name ="TxtAD_Ville1"
                    ControlSource ="AD_Ville"
                End
            End
        End
        Begin FormFooter
            Height =226
            BackColor =12632256
            Name ="PiedFormulaire"
        End
    End
End
CodeBehindForm
' See "FrmDebiteurEncaisseurSelectionListe.cls"
