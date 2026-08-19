Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =7650
    ItemSuffix =53
    Left =2775
    Top =1470
    Right =10785
    Bottom =9540
    Filter ="Déb_Code='AU00500009'"
    RecSrcDt = Begin
        0xb8829322efe5e240
    End
    RecordSource ="SELECT DEBITEUR.* FROM DEBITEUR; "
    Caption ="Détail du Principal"
    OnOpen ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Rectangle
            Width =850
            Height =850
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
            Height =567
            BackColor =12632256
            Name ="EntêteFormulaire0"
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
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
            End
        End
        Begin Section
            Height =5685
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =2
                    Left =6207
                    Top =1155
                    Width =981
                    TabIndex =5
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtMonnaie"
                    ControlSource ="=DLookUp(\"Mon_Nom\",\"MONNAIE\",\"Mon_Id=\" & [Déb_Mont_Prin_mo])"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =4380
                            Top =1155
                            Width =1755
                            Height =240
                            FontWeight =700
                            ForeColor =8421504
                            Name ="EtiMonnaie"
                            Caption ="Monnaie d'origine :"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1358
                    Top =1474
                    Width =2493
                    Height =255
                    TabIndex =6
                    Name ="TxtDéb_Réf1_Client"
                    ControlSource ="Déb_Réf1_Client"
                    StatusBarText ="Pourcentage"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    Left =5219
                    Top =1470
                    Width =1939
                    Height =255
                    TabIndex =7
                    Name ="TxtDéb_Réf2_Client"
                    ControlSource ="Déb_Réf2_Client"
                    StatusBarText ="Montant des honoraires"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =4260
                    Top =1470
                    Width =915
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiDéb_Réf2_Client"
                    Caption ="Réf.2 :"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =90
                    Top =1470
                    Width =1230
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiDéb_Réf1_Client"
                    Caption ="Réf.1 ou Lot :"
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1395
                    Top =2775
                    Width =5808
                    Height =960
                    TabIndex =13
                    Name ="TxtDéb_Observation"
                    ControlSource ="Déb_Observation"
                    StatusBarText ="Pourcentage"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1373
                    Top =195
                    Width =5790
                    Height =255
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Typ_Cré"
                    ControlSource ="Déb_Typ_Cré"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1373
                    Top =829
                    Width =1053
                    Height =255
                    TabIndex =2
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Mont_Récup_Cli"
                    ControlSource ="Déb_Mont_Récup_Cli"
                    Format ="Standard"
                End
                Begin Label
                    OverlapFlags =85
                    Left =615
                    Top =195
                    Width =720
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8421504
                    Name ="EtiDéb_Typ_Cré"
                    Caption ="Objet  :"
                End
                Begin Label
                    OverlapFlags =85
                    Left =2490
                    Top =810
                    Width =1830
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8421504
                    Name ="EtiDéb_Mont_Récup_Cli"
                    Caption =": Acompte antérieur"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =2
                    Left =3203
                    Top =1151
                    Width =1053
                    Height =255
                    TabIndex =4
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Date_Cré"
                    ControlSource ="Déb_Date_Cré"
                    Format ="dd/mm/yyyy"
                End
                Begin Label
                    OverlapFlags =85
                    Left =2505
                    Top =1155
                    Width =675
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8421504
                    Name ="EtiDéb_Date_Cré"
                    Caption ="Date  :"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1373
                    Top =1144
                    Width =1053
                    Height =255
                    TabIndex =3
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Mont_Prin_m"
                    ControlSource ="Déb_Mont_Prin_m1"
                    Format ="Standard"
                End
                Begin Label
                    OverlapFlags =85
                    Left =75
                    Top =1140
                    Width =1275
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8421504
                    Name ="EtiDéb_Mont_Prin_m"
                    Caption ="Principal dû  :"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1373
                    Top =514
                    Width =1053
                    Height =255
                    TabIndex =1
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Mont_Origine_Cli"
                    ControlSource ="Déb_Mont_Origine_Cli"
                    Format ="Standard"
                End
                Begin Label
                    OverlapFlags =85
                    Left =2490
                    Top =510
                    Width =1530
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8421504
                    Name ="EtiDéb_Mont_Origine_Cli"
                    Caption =": Principal initial"
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1358
                    Top =1808
                    Width =663
                    Height =255
                    TabIndex =8
                    Name ="TxtDéb_RIB_Banque"
                    ControlSource ="Déb_RIB_Banque"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =90
                    Top =1800
                    Width =1230
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiRIB"
                    Caption ="RIB :"
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =2070
                    Top =1815
                    Width =663
                    Height =255
                    TabIndex =9
                    Name ="TxtDéb_RIB_Guichet"
                    ControlSource ="Déb_RIB_Guichet"
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =2775
                    Top =1815
                    Width =1248
                    Height =255
                    TabIndex =10
                    Name ="TxtDéb_RIB_Compte"
                    ControlSource ="Déb_RIB_Compte"
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =4065
                    Top =1815
                    Width =348
                    Height =255
                    TabIndex =11
                    Name ="TxtDéb_RIB_Clé"
                    ControlSource ="Déb_RIB_Clé"
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1358
                    Top =2160
                    Width =3051
                    TabIndex =12
                    Name ="TxtDéb_RIB_Dom"
                    ControlSource ="Déb_RIB_Dom"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =735
                            Top =2145
                            Width =585
                            Height =240
                            FontWeight =700
                            Name ="EtiDéb_RIB_Dom"
                            Caption ="Dom :"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =1380
                    Top =2490
                    Width =2025
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiDéb_Observation"
                    Caption ="Suivi du dossier : "
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1395
                    Top =4140
                    Width =5808
                    Height =960
                    TabIndex =14
                    Name ="TxtDéb_Phoning"
                    ControlSource ="Déb_Phoning"
                    StatusBarText ="Pourcentage"
                End
                Begin Label
                    OverlapFlags =85
                    Left =1380
                    Top =3855
                    Width =2025
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiDéb_Phoning"
                    Caption ="Suivi du telephone : "
                End
            End
        End
        Begin FormFooter
            Height =675
            BackColor =12632256
            Name ="PiedFormulaire"
            Begin
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3405
                    Top =120
                    Width =1140
                    Height =465
                    FontWeight =300
                    Name ="CmdOK"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "FrmFicheDebiteurDetail.cls"
