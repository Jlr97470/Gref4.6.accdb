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
    Width =7935
    ItemSuffix =81
    Left =3315
    Top =795
    Right =11250
    Bottom =7050
    RecSrcDt = Begin
        0x4666cc102dede240
    End
    RecordSource ="SELECT DEBITEUR.*, ADRESSE_DEBITEUR.AD_CodePostal, ADRESSE_DEBITEUR.AD_Ville FRO"
        "M DEBITEUR INNER JOIN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Code=ADRESSE_DEBITEUR.AD_"
        "Code_Débiteur WHERE DEBITEUR.Déb_MonnaieEnCours=2; "
    Caption ="Liste des Débiteurs"
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
            Height =2266
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Top =2025
                    Width =1080
                    Height =240
                    Name ="EtiDéb_Code"
                    Caption ="Code"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    Left =1110
                    Top =2025
                    Width =3585
                    Height =240
                    Name ="EtiDéb_Nom"
                    Caption ="Nom"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =105
                    Top =30
                    Width =486
                    Height =456
                    TabIndex =12
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
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =87
                    Left =4695
                    Top =2025
                    Width =1140
                    Height =240
                    Name ="EtiDéb_Mont_Prin_m"
                    Caption ="Créance"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =5895
                    Top =2026
                    Width =420
                    Height =240
                    Name ="EtiDéb_Code_Enc"
                    Caption ="Enc."
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =6330
                    Top =2026
                    Width =420
                    Height =240
                    Name ="EtiDéb_Situation"
                    Caption ="Sit."
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =690
                    Top =15
                    Width =7101
                    Height =1925
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
                            Name ="EtDéb_Code"
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
                    Name ="TxtDéb_Mont_Prin_m"
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
                            Name ="EtiRech_Déb_Mont_Prin_m"
                            Caption ="Créance :"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =223
                    TextAlign =2
                    Left =6917
                    Top =225
                    Width =397
                    Height =240
                    TabIndex =5
                    Name ="TxtDéb_Code_Enc"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =5490
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
                    Left =5280
                    Top =1575
                    Width =2055
                    Height =315
                    TabIndex =9
                    Name ="CmdOK"
                    Caption ="&Rechercher"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =223
                    TextAlign =2
                    Left =6919
                    Top =495
                    Width =397
                    Height =240
                    TabIndex =6
                    Name ="TxtDéb_Situation"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =5492
                            Top =495
                            Width =1365
                            Height =240
                            FontSize =7
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiRech_Déb_Situation"
                            Caption ="Code Situation :"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =6182
                    Top =764
                    Width =1131
                    Height =240
                    TabIndex =7
                    Name ="TxtAD_CodePostal"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =4815
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
                    Left =4943
                    Top =1034
                    Width =2376
                    Height =240
                    TabIndex =8
                    Name ="TxtAD_Ville"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =4365
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
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =6765
                    Top =2026
                    Width =420
                    Height =240
                    Name ="EtiDéb_niv_relance"
                    Caption ="Rel."
                End
                Begin ComboBox
                    SpecialEffect =0
                    OverlapFlags =223
                    ColumnCount =2
                    ListRows =2
                    Left =2537
                    Top =1125
                    Width =1131
                    TabIndex =3
                    Name ="CmbDéb_MonnaieEnCours"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Mon_Id, Mon_Nom FROM MONNAIE; "
                    ColumnWidths ="0;1134"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =1170
                            Top =1125
                            Width =1290
                            Height =240
                            FontSize =7
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiDéb_MonnaieEnCours"
                            Caption ="Monnaie :"
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =255
                    Left =690
                    Top =180
                    Width =6968
                    Height =1745
                    TabIndex =11
                    Name ="CadClass"
                    DefaultValue ="1"
                    Begin
                        Begin OptionButton
                            OverlapFlags =247
                            Left =7470
                            Top =270
                            OptionValue =5
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
                            Left =7471
                            Top =539
                            OptionValue =6
                            Name ="OptDéb_Situation"
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
                            Left =7471
                            Top =825
                            OptionValue =7
                            Name ="OptAD_CodePostal"
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =3840
                            Top =915
                            OptionValue =3
                            Name ="OptDéb_Mont_Prin_m"
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =7471
                            Top =1095
                            OptionValue =8
                            Name ="OptAD_Ville"
                        End
                        Begin OptionButton
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =3840
                            Top =1200
                            OptionValue =4
                            Name ="OptDéb_MonnaieEnCours"
                        End
                    End
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =105
                    Top =525
                    Width =485
                    Height =455
                    FontSize =7
                    FontWeight =400
                    TabIndex =13
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
                        0x0000008888888888
                    End
                    FontName ="Small Fonts"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =86
                    TextFontFamily =34
                    Left =105
                    Top =1020
                    Width =487
                    Height =472
                    FontSize =8
                    TabIndex =14
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
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    Left =7170
                    Top =2025
                    Width =525
                    Height =240
                    Name ="EtiDéb_Npai1"
                    Caption ="Npai."
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =2537
                    Top =1425
                    Width =1131
                    Height =240
                    TabIndex =4
                    Name ="TxtDéb_Réf_Client"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =1170
                            Top =1425
                            Width =1290
                            Height =240
                            FontSize =7
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiDéb_Réf_Client"
                            Caption ="Référence :"
                        End
                    End
                End
                Begin CheckBox
                    TripleState = NotDefault
                    OverlapFlags =247
                    Left =7110
                    Top =1350
                    TabIndex =10
                    Name ="ChkDéb_Npai"
                    StatusBarText ="Debiteur Npai"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =6420
                            Top =1305
                            Width =555
                            Height =240
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiDéb_Npai"
                            Caption ="Npai :"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =73
                    TextFontFamily =34
                    Left =105
                    Top =1530
                    Width =486
                    Height =471
                    FontSize =8
                    TabIndex =15
                    Name ="CmdExporter"
                    Caption ="&i"
                    StatusBarText ="Imprimer le rapport"
                    OnClick ="=ImpVisu(5)"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaad000000000000add0330880000330da ,
                        0xa0330880000330add0330880000330daa0330000000330add0333333333330da ,
                        0xa0330000000330add030fffffff030daa030fffcfff030add030ffcccff030da ,
                        0xa030fffcccfdadadd030ffffcccacadaa000000daccccdaddadadadadacccada ,
                        0xadadadadaccccdad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
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
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =7560
                    Top =30
                    Width =320
                    Height =245
                    FontSize =6
                    FontWeight =400
                    TabIndex =10
                    Name ="CmdDetail"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                    FontName ="Small Fonts"
                End
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
                    OverlapFlags =93
                    TextAlign =1
                    Left =4755
                    Top =15
                    Width =1131
                    Height =240
                    TabIndex =2
                    Name ="TxtDéb_Mont_Prin_m1"
                    ControlSource ="Déb_Mont_Prin_m2"
                    Format ="Standard"
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
                    TabIndex =1
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
                    TabIndex =4
                    Name ="TxtDéb_Code_Enc1"
                    ControlSource ="Déb_Code_Enc"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =6350
                    Top =15
                    Width =397
                    Height =240
                    TabIndex =5
                    Name ="TxtDéb_Situation1"
                    ControlSource ="Déb_Situation"
                    AfterUpdate ="[Event Procedure]"
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
                    TabIndex =7
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
                    TabIndex =8
                    Name ="TxtAD_Ville1"
                    ControlSource ="AD_Ville"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =6785
                    Top =15
                    Width =397
                    Height =240
                    TabIndex =6
                    Name ="TxtDéb_niv_relance"
                    ControlSource ="Déb_niv_relance"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    Left =4755
                    Top =15
                    Width =1127
                    Height =242
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="TxtMon_Nom"
                    ControlSource ="=Dlookup(\"Mon_Nom\",\"MONNAIE\",\"Mon_ID=\" & [Déb_MonnaieEnCours])"
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7305
                    Top =45
                    TabIndex =9
                    Name ="ChkDéb_Npai1"
                    ControlSource ="Déb_Npai"
                    StatusBarText ="Debiteur Npai"
                    AfterUpdate ="[Event Procedure]"
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
' See "FrmFicheDebiteurListeRechercher.cls"
