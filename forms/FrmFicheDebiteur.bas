Version =20
VersionRequired =20
Begin Form
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    PictureTiling = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    PictureSizeMode =1
    GridX =5
    GridY =5
    Width =9297
    ItemSuffix =220
    Left =2805
    Top =1275
    Right =12390
    Bottom =7725
    AfterDelConfirm ="[Event Procedure]"
    Filter ="Déb_Code='2000100039'"
    RecSrcDt = Begin
        0xf28ea10837efe240
    End
    RecordSource ="SELECT DEBITEUR.* FROM DEBITEUR; "
    Caption ="Détail d'un Débiteur"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnKeyDown ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
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
        Begin Line
            Width =1701
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
        Begin Subform
            Width =1701
            Height =1701
        End
        Begin ToggleButton
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
        End
        Begin Tab
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =675
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =113
                    Top =113
                    Width =471
                    Height =471
                    TabIndex =4
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
                    OverlapFlags =93
                    TextAlign =1
                    Left =1290
                    Top =150
                    Width =2580
                    Height =405
                    FontSize =14
                    ForeColor =8355711
                    Name ="EtiFicheDebiteurTitreOmbre"
                    Caption ="La Fiche Débiteur"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =1260
                    Top =105
                    Width =2580
                    Height =405
                    FontSize =14
                    ForeColor =16777215
                    Name ="EtiFicheDebiteurTitre"
                    Caption ="La Fiche Débiteur"
                End
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =93
                    Left =3990
                    Top =120
                    Width =1800
                    Height =420
                    BackColor =12632256
                    Name ="RecCode_Débiteur"
                End
                Begin TextBox
                    AutoTab = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =4350
                    Top =210
                    Width =1080
                    Name ="TxtCode_Débiteur"
                    StatusBarText ="Code du débiteur"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =5910
                    Top =165
                    Width =1256
                    Height =347
                    TabIndex =2
                    Name ="CmdInverseur"
                    Caption ="Inverseur"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =7230
                    Top =165
                    Width =1826
                    Height =347
                    TabIndex =3
                    Name ="CmdBasculer"
                    Caption ="Basculer En Franc"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =87
                    TextFontFamily =34
                    Left =4050
                    Top =195
                    Width =276
                    Height =291
                    TabIndex =5
                    Name ="CmdPrecedent"
                    Caption ="&w"
                    StatusBarText ="Ajouter un nouveau type d'adresse"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadadadad1dadadaadadadad11adadaddadadad111dadada ,
                        0xadadad1111adadaddadad11111dadadaadadad1111adadaddadadad111dadada ,
                        0xadadadad11adadaddadadadad1dadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =88
                    TextFontFamily =34
                    Left =5475
                    Top =195
                    Width =276
                    Height =291
                    TabIndex =1
                    Name ="CmdSuivant"
                    Caption ="&x"
                    StatusBarText ="Ajouter un nouveau type d'adresse"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadada1adadadadaadadad11adadadaddadada111adadada ,
                        0xadadad1111adadaddadada11111adadaadadad1111adadaddadada111adadada ,
                        0xadadad11adadadaddadada1adadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
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
            CanGrow = NotDefault
            Height =5215
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin Subform
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =300
                    Top =3405
                    Width =6330
                    Height =1650
                    TabIndex =27
                    Name ="SFrmFicheDebiteurEncaissement"
                    SourceObject ="Form.SFrmFicheDebiteurEncaissement"
                    LinkChildFields ="Enc_Code_Débiteur"
                    LinkMasterFields ="Déb_Code"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =90
                    Top =60
                    Width =227
                    Height =3862
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="TxtDéb_MonnaieEnCours2"
                    ControlSource ="=DLookUp(\"Mon_Nom\",\"MONNAIE\",\"Mon_Id=\" & [Déb_MonnaieEnCours])"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =5460
                    Top =75
                    Width =227
                    Height =4312
                    FontSize =12
                    ForeColor =16777215
                    Name ="EtiMonnaieVue2"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    ColumnCount =2
                    ListWidth =3232
                    Left =1380
                    Top =1875
                    Width =561
                    TabIndex =13
                    Name ="CmbDéb_Situation"
                    ControlSource ="Déb_Situation"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CODE_GESTION.CG_Code, CODE_GESTION.CG_Libellé FROM CODE_GESTION; "
                    ColumnWidths ="398;2835"
                    StatusBarText ="Code situation"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =1
                            Left =375
                            Top =1875
                            Width =1020
                            Height =240
                            Name ="EtiDéb_Situation"
                            Caption ="Situation  :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    Left =1538
                    Top =465
                    Width =3675
                    TabIndex =7
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                    StatusBarText ="Nom du débiteur"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =1
                            Left =435
                            Top =465
                            Width =1035
                            Height =240
                            Name ="EtiDéb_Nom"
                            Caption ="Nom        :"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =1538
                    Top =135
                    Width =1035
                    TabIndex =2
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =1
                            Left =435
                            Top =135
                            Width =1035
                            Height =240
                            Name ="EtiDéb_Code"
                            Caption ="Code       :"
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =93
                    Left =5745
                    Top =270
                    Width =3375
                    Height =1245
                    Name ="RecClient"
                End
                Begin Label
                    OverlapFlags =223
                    Left =5940
                    Top =135
                    Width =570
                    Height =240
                    ForeColor =16777215
                    Name ="EtiDéb_Code_Client"
                    Caption ="Client"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    Left =4516
                    Top =135
                    Width =697
                    TabIndex =3
                    Name ="CmbDéb_Qualité"
                    ControlSource ="Déb_Qualité"
                    RowSourceType ="Value List"
                    RowSource ="ASS;ENF;ENT;M;M/M;MLE;MME;MR;STE"
                    StatusBarText ="Qualité du débiteur"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =3600
                            Top =135
                            Width =810
                            Height =240
                            Name ="EtiDéb_Qualité"
                            Caption ="Qualité :"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =6510
                    Top =135
                    Width =645
                    Height =240
                    FontWeight =700
                    TabIndex =4
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="TxtDéb_Code_Client"
                    ControlSource ="Déb_Code_Client"
                    StatusBarText ="Code client"
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =8820
                    Top =405
                    Width =245
                    Height =245
                    TabIndex =6
                    Name ="CmdFicheClient"
                    Caption ="..."
                    StatusBarText ="Fiche du client"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000000a000000040000000100040000000000200000000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888800000080080080080000008008008008000000 ,
                        0x8888888888000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =1
                    BackStyle =0
                    Left =7755
                    Top =675
                    Width =993
                    TabIndex =8
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="TxtDéb_Date_der_rég"
                    ControlSource ="=Dmax(\"Reg_Date\",\"SelReglementDetail\",\"RegD_Déb_Code='\" & [Déb_Code] & \"'"
                        "\")"
                    Format ="Short Date"
                    BeforeUpdate ="[Event procedure]"
                    AfterUpdate ="[Event procedure]"
                    OnDblClick ="[Event procedure]"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =215
                            Left =5940
                            Top =675
                            Width =1770
                            Height =240
                            Name ="EtiDéb_Date_der_règ"
                            Caption ="Dernier règlement le"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =3
                    BackStyle =0
                    Left =7605
                    Top =930
                    Width =1143
                    TabIndex =11
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="TxtDéb_Mont_reversé_m1"
                    ControlSource ="=nz(DSum(\"Enc_Prin_m1\",\"ENCAISSEMENT\",\"Enc_Code_Débiteur='\" & [Déb_Code] &"
                        " \"' AND Enc_Montant_mo=1 AND Enc_Situation Is Not Null\"),0)"
                    Format ="#,##0.00\" F\""
                    AfterUpdate ="[Event procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnKeyDown ="[Event Procedure]"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =7764
                    Top =1920
                    Width =993
                    TabIndex =20
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Mont_Prin_m"
                    ControlSource ="Déb_Mont_Prin_m2"
                    Format ="Standard"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =0
                            Left =7035
                            Top =1920
                            Width =675
                            Height =240
                            Name ="EtiDéb_Mont_Prin_m"
                            Caption ="Prin.  :"
                        End
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    Left =5940
                    Top =1920
                    Width =330
                    Height =240
                    Name ="EtiDu"
                    Caption ="Dû"
                End
                Begin TextBox
                    OverlapFlags =93
                    Left =7764
                    Top =2220
                    Width =993
                    TabIndex =21
                    Name ="TxtDéb_Frais_m"
                    ControlSource ="Déb_Frais_m2"
                    Format ="Standard"
                    StatusBarText ="Frais de la créance"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =0
                            Left =7035
                            Top =2220
                            Width =675
                            Height =240
                            Name ="EtiDéb_Frais_m"
                            Caption ="Frais  :"
                        End
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    Left =5925
                    Top =2940
                    Width =855
                    Height =240
                    Name ="EtiEncaisse"
                    Caption ="Encaissé"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =7764
                    Top =2940
                    Width =993
                    TabIndex =23
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Prin_enc_m"
                    ControlSource ="=Round(NZ(DSum(\"Enc_Prin_m2\",\"ENCAISSEMENT\",\"Enc_Code_Débiteur='\" & [Déb_C"
                        "ode] & \"'\"),0),2)"
                    Format ="Standard"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =0
                            Left =7035
                            Top =2940
                            Width =660
                            Height =240
                            Name ="EtiDéb_Prin_enc_m"
                            Caption ="Prin.  :"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =7764
                    Top =3240
                    Width =993
                    TabIndex =24
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Frais_enc_m"
                    ControlSource ="=Round(NZ(DSum(\"Enc_Frais_m2\",\"ENCAISSEMENT\",\"Enc_Code_Débiteur='\" & [Déb_"
                        "Code] & \"'\"),0),2)"
                    Format ="Standard"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =0
                            Left =7035
                            Top =3240
                            Width =675
                            Height =240
                            Name ="EtiDéb_Frais_enc_m"
                            Caption ="Frais  :"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    Left =7764
                    Top =2535
                    Width =993
                    TabIndex =22
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="TxtTotalDu"
                    ControlSource ="=Round(Déb_Mont_Prin_m2,2)+Round(Déb_Frais_m2,2)"
                    Format ="Standard"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =1
                            Left =7035
                            Top =2535
                            Width =675
                            Height =240
                            FontWeight =400
                            ForeColor =16777215
                            Name ="EtiTotalDu"
                            Caption ="Total"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    Left =7764
                    Top =3555
                    Width =993
                    TabIndex =25
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="TxtTotalEncaisse"
                    ControlSource ="=Round(NZ(DSum(\"Enc_Prin_m2\",\"ENCAISSEMENT\",\"Enc_Code_Débiteur='\" & [Déb_C"
                        "ode] & \"'\"),0),2)+Round(NZ(DSum(\"Enc_Frais_m2\",\"ENCAISSEMENT\",\"Enc_Code_D"
                        "ébiteur='\" & [Déb_Code] & \"'\"),0),2)"
                    Format ="Standard"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =1
                            Left =7035
                            Top =3570
                            Width =660
                            Height =240
                            FontWeight =400
                            ForeColor =16777215
                            Name ="EtiTotalEncaisse"
                            Caption ="Total"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =7764
                    Top =3900
                    Width =993
                    TabIndex =26
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Date_der_enc"
                    ControlSource ="=DMax(\"Enc_Date\",\"ENCAISSEMENT\",\"Enc_Code_Débiteur='\" & [Déb_Code] & \"'\""
                        ")"
                    Format ="Short Date"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =1
                            Left =7035
                            Top =3900
                            Width =660
                            Height =240
                            Name ="EtiDéb_Date_der_enc"
                            Caption ="Date"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    BackStyle =0
                    Left =7860
                    Top =4635
                    Width =993
                    TabIndex =37
                    BackColor =12632256
                    ForeColor =255
                    Name ="TxtSolde"
                    ControlSource ="=[TxtTotalDu]-[TxtTotalEncaisse]"
                    Format ="Standard"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =1553
                    Top =2535
                    Width =1038
                    TabIndex =18
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Date_der_rel"
                    ControlSource ="Déb_Date_der_rel"
                    Format ="dd/mm/yyyy"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =0
                            Left =375
                            Top =2535
                            Width =1140
                            Height =240
                            Name ="EtiDéb_Date_der_rel"
                            Caption ="Date"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =1553
                    Top =2869
                    Width =1038
                    TabIndex =19
                    Name ="TxtDéb_Date_proch_rel"
                    ControlSource ="Déb_Date_proch_rel"
                    Format ="dd/mm/yyyy"
                    StatusBarText ="Date de la prochaine relance"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =0
                            Left =375
                            Top =2865
                            Width =1140
                            Height =240
                            Name ="EtiDéb_Date_proch_rel"
                            Caption ="Prochaine"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    Left =2010
                    Top =1875
                    Width =2294
                    Height =240
                    TabIndex =14
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="TxtCG_Libellé"
                    ControlSource ="=CmbDéb_Situation.column(1)"
                End
                Begin Line
                    OverlapFlags =93
                    Left =7770
                    Top =2520
                    Width =990
                    Name ="TrtDu"
                End
                Begin Line
                    OverlapFlags =93
                    Left =7770
                    Top =3540
                    Width =990
                    Name ="TrtEncaisse"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =4905
                    Top =1875
                    Width =397
                    TabIndex =15
                    Name ="TxtDéb_Code_Enc"
                    ControlSource ="Déb_Code_Enc"
                    StatusBarText ="Numéro de l'encaisseur"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =4335
                            Top =1875
                            Width =540
                            Height =240
                            Name ="EtiDéb_Code_Enc"
                            Caption ="Enc.:"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    Left =6735
                    Top =4620
                    Width =993
                    TabIndex =38
                    BackColor =12632256
                    ForeColor =255
                    Name ="TxtSoldePrincipal"
                    ControlSource ="=[Déb_Mont_Prin_m1]-[TxtDéb_Prin_enc_m]"
                    Format ="Standard"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =6975
                            Top =4350
                            Width =480
                            Height =240
                            Name ="EtiSoldePrincipal"
                            Caption ="Prin."
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =255
                    BackStyle =0
                    Left =7860
                    Top =4620
                    Width =993
                    TabIndex =39
                    BackColor =12632256
                    ForeColor =255
                    Name ="TxtSoldeFrais"
                    ControlSource ="=[Déb_Frais_m1]-[TxtDéb_Frais_enc_m]"
                    Format ="Standard"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =8100
                            Top =4350
                            Width =480
                            Height =240
                            Name ="EtiSoldeFrais"
                            Caption ="Frais"
                        End
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =0
                    Left =5925
                    Top =4620
                    Width =690
                    Height =240
                    Name ="EtiSolde"
                    Caption ="Solde :"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    Left =5940
                    Top =405
                    Width =2804
                    Height =240
                    TabIndex =5
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="TxtCli_Nom"
                    ControlSource ="=DLookUp(\"Cli_Nom\",\"CLIENT\",\"Cli_Code='\" & [Déb_Code_Client] & \"'\")"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    ColumnCount =2
                    ListRows =7
                    ListWidth =3232
                    Left =1380
                    Top =2205
                    Width =561
                    TabIndex =16
                    Name ="CmbDéb_niv_relance"
                    ControlSource ="Déb_niv_relance"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RELANCE.Rel_N°, RELANCE.Rel_Nom FROM RELANCE INNER JOIN CODE_GESTION ON R"
                        "ELANCE.Rel_Code_Action=CODE_GESTION.CG_Code_Action WHERE (((CODE_GESTION.CG_Code"
                        ")=Formulaires!FrmFicheDebiteur!CmbDéb_Situation)); "
                    ColumnWidths ="398;2835"
                    StatusBarText ="Niveau de relance"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =1
                            Left =375
                            Top =2205
                            Width =990
                            Height =240
                            Name ="EtiDéb_niv_relance"
                            Caption ="Relance  :"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4965
                    Top =810
                    Width =260
                    Height =245
                    TabIndex =10
                    Name ="CmdTypeAdresse"
                    Caption ="..."
                    StatusBarText ="Ajouter un nouveau type d'adresse"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000000a000000040000000100040000000000200000000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888800000080080080080000008008008008000000 ,
                        0x8888888888000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    Left =1545
                    Top =795
                    Width =3210
                    Height =885
                    TabIndex =9
                    Name ="SFrmFicheDebiteurAdresse"
                    SourceObject ="Form.SFrmFicheDebiteurAdresse"
                    LinkChildFields ="AD_Code_Débiteur"
                    LinkMasterFields ="Déb_Code"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =0
                            Left =435
                            Top =810
                            Width =1035
                            Height =240
                            Name ="EtiSFrmFicheDebiteurAdresse"
                            Caption ="Adresses  :"
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =1665
                    Top =911
                    Width =3198
                    Height =904
                    BackColor =8421504
                    Name ="RecSFrmFicheDebiteurAdresse"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    Left =5445
                    Top =60
                    Width =227
                    Height =4312
                    FontSize =12
                    ForeColor =255
                    Name ="EtiMonnaieVue1"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    BackStyle =0
                    Left =7607
                    Top =1185
                    Width =1143
                    TabIndex =12
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="TxtDéb_Mont_reversé_m2"
                    ControlSource ="=nz(DSum(\"Enc_Prin_m2\",\"ENCAISSEMENT\",\"Enc_Code_Débiteur='\" & [Déb_Code] &"
                        " \"' AND Enc_Montant_mo=2 AND Enc_Situation Is Not Null\"),0)"
                    Format ="#,##0.00 €"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =5940
                    Top =945
                    Width =1425
                    Height =240
                    Name ="EtiDéb_Mont_reversé_m1"
                    Caption ="Montant réglé :"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =7380
                    Top =1200
                    Width =180
                    Height =240
                    Name ="EtiDéb_Mont_reversé_m2"
                    Caption ="+"
                End
                Begin TextBox
                    OverlapFlags =93
                    Left =1718
                    Top =3735
                    Width =3510
                    TabIndex =28
                    Name ="TxtDéb_Typ_Cré"
                    ControlSource ="Déb_Typ_Cré"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =0
                            Left =735
                            Top =3750
                            Width =960
                            Height =240
                            ForeColor =16777215
                            Name ="EtiDéb_Typ_Cré"
                            Caption ="Objet      :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =1718
                    Top =4050
                    Width =993
                    TabIndex =29
                    Name ="TxtDéb_Date_Cré"
                    ControlSource ="Déb_Date_Cré"
                    Format ="dd/mm/yyyy"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =0
                            Left =735
                            Top =4050
                            Width =975
                            Height =240
                            ForeColor =16777215
                            Name ="EtiDéb_Date_Cré"
                            Caption ="Date       :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    Left =1718
                    Top =4369
                    Width =993
                    TabIndex =30
                    Name ="TxtDéb_Mont_Prin"
                    ControlSource ="Déb_Mont_Prin_m2"
                    Format ="Standard"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =0
                            Left =735
                            Top =4365
                            Width =960
                            Height =240
                            ForeColor =16777215
                            Name ="EtiDéb_Mont_Prin"
                            Caption ="Montant  :"
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =255
                    Left =345
                    Top =3540
                    Width =5055
                    Height =1515
                    Name ="RecPrincipal"
                End
                Begin Label
                    OverlapFlags =255
                    Left =735
                    Top =3390
                    Width =825
                    Height =240
                    ForeColor =16777215
                    Name ="EtiPrincipal"
                    Caption ="Principal"
                End
                Begin CommandButton
                    OverlapFlags =255
                    TextFontFamily =34
                    Left =2850
                    Top =4365
                    Width =260
                    Height =245
                    TabIndex =33
                    Name ="CmdDebiteurDetail"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000000a000000040000000100040000000000200000000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888800000080080080080000008008008008000000 ,
                        0x8888888888000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =255
                    Left =5745
                    Top =1785
                    Width =3375
                    Height =3270
                    Name ="RecCreance"
                End
                Begin Rectangle
                    SpecialEffect =1
                    BackStyle =0
                    OverlapFlags =255
                    Left =5865
                    Top =4350
                    Width =3135
                    Height =615
                    Name ="RecSolde"
                End
                Begin Label
                    OverlapFlags =247
                    Left =5940
                    Top =1650
                    Width =795
                    Height =240
                    ForeColor =16777215
                    Name ="EtiCreance"
                    Caption ="Créance"
                End
                Begin TextBox
                    OverlapFlags =255
                    TextAlign =2
                    Left =3105
                    Top =4695
                    Width =1053
                    TabIndex =34
                    Name ="TxtDéb_Date_Représentation"
                    ControlSource ="Déb_Date_Représentation"
                    Format ="dd/mm/yyyy"
                    StatusBarText ="Pourcentage"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =1
                            Left =1395
                            Top =4725
                            Width =1710
                            Height =240
                            Name ="EtiDéb_Date_Représentation"
                            Caption ="Représentation le :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    ColumnCount =2
                    ListWidth =3572
                    Left =4275
                    Top =4695
                    Width =981
                    TabIndex =35
                    Name ="CmbDéb_Repré_Type"
                    ControlSource ="Déb_Repré_Type"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ENCAISSEMENT_TYPE.* FROM ENCAISSEMENT_TYPE; "
                    ColumnWidths ="737;2835"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    BackStyle =0
                    Left =3195
                    Top =4350
                    Width =453
                    FontSize =12
                    FontWeight =700
                    TabIndex =31
                    BackColor =12632256
                    ForeColor =6723891
                    Name ="TxtDéb_Observation"
                    ControlSource ="=IIf([Déb_Observation] Is Null,\"\",\"ü\")"
                    FontName ="Wingdings"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =75
                    Top =45
                    Width =227
                    Height =3862
                    FontSize =12
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8388608
                    Name ="TxtDéb_MonnaieEnCours1"
                    ControlSource ="=DLookUp(\"Mon_Nom\",\"MONNAIE\",\"Mon_Id=\" & [Déb_MonnaieEnCours])"
                End
                Begin CommandButton
                    Transparent = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =5850
                    Top =4320
                    Width =3165
                    Height =660
                    TabIndex =36
                    Name ="CmdSolde"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    Left =315
                    Top =3165
                    Width =2850
                    Height =240
                    ForeColor =16777215
                    Name ="EtiSFrmFicheDebiteurEncaissement"
                    Caption ="Encaissements"
                End
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =255
                    Left =405
                    Top =3585
                    Width =6345
                    Height =1560
                    BackColor =8421504
                    Name ="RecSFrmFicheDebiteurEncaissement"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    Left =2010
                    Top =2235
                    Width =3299
                    Height =240
                    TabIndex =17
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="TxtRel_Nom"
                    ControlSource ="=CmbDéb_niv_relance.column(1)"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    BackStyle =0
                    Left =3750
                    Top =4350
                    Width =453
                    FontSize =12
                    FontWeight =700
                    TabIndex =32
                    BackColor =12632256
                    ForeColor =6723891
                    Name ="TxtDéb_Phoning"
                    ControlSource ="=IIf([Déb_Phoning] Is Null,\"\",\"v\")"
                    FontName ="Wingdings"
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =5175
                    Top =2595
                    TabIndex =40
                    Name ="Déb_Npai"
                    ControlSource ="Déb_Npai"
                    StatusBarText ="Debiteur Npai"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =4095
                    Top =2565
                    Width =1005
                    Height =240
                    Name ="EtiDéb_Npai"
                    Caption ="Npai:"
                End
            End
        End
        Begin FormFooter
            Height =570
            BackColor =12632256
            Name ="PiedFormulaire"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =82
                    Left =4980
                    Top =90
                    Width =2098
                    Height =405
                    Name ="CmdDebiteurListe"
                    Caption ="&Recherche Débiteur"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =78
                    Left =7140
                    Top =90
                    Width =2098
                    Height =405
                    TabIndex =1
                    Name ="CmdDebiteurNouveau"
                    Caption ="&Nouveau Débiteur"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    AccessKey =67
                    Left =645
                    Top =90
                    Width =2098
                    Height =405
                    TabIndex =2
                    Name ="CmdEcheance"
                    Caption ="Echeances &CB"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =67
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =93
                    AccessKey =69
                    Left =2805
                    Top =90
                    Width =2098
                    Height =405
                    TabIndex =3
                    Name ="CmdEncaissement"
                    Caption ="&Encaissements"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Visible = NotDefault
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =80
                    Left =2820
                    Top =90
                    Width =2098
                    Height =405
                    TabIndex =4
                    Name ="CmdPrincipal"
                    Caption ="&Principal"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =86
                    TextFontFamily =34
                    Left =75
                    Top =60
                    Width =487
                    Height =472
                    FontSize =8
                    TabIndex =5
                    Name ="CmdVisualiser"
                    Caption ="&v"
                    OnClick ="[Event Procedure]"
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
    End
End
CodeBehindForm
' See "FrmFicheDebiteur.cls"
