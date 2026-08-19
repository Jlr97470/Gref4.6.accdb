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
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =6810
    ItemSuffix =50
    Left =5250
    Top =330
    Right =12420
    Bottom =6720
    RecSrcDt = Begin
        0x917ba68f4eeee240
    End
    RecordSource ="SELECT SelDebiteurEncaissementIncorrect.* FROM SelDebiteurEncaissementIncorrect "
        "WHERE (((IIf(IsNull(Formulaires!FrmDebiteurEncaissementIncorrect!TxtDateDebut),T"
        "rue,Formulaires!FrmDebiteurEncaissementIncorrect!TxtDateDebut<=[Déb_Date_Cré]))="
        "True) And ((IIf(IsNull(Formulaires!FrmDebiteurEncaissementIncorrect!TxtDateFin),"
        "True,Formulaires!FrmDebiteurEncaissementIncorrect!TxtDateFin>=[Déb_Date_Cré]))=T"
        "rue)); "
    Caption ="Liste des Débiteurs incorrects"
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
            Height =1905
            BackColor =12632256
            Name ="EntêteFormulaire1"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Top =1665
                    Width =1080
                    Height =240
                    Name ="EtiDéb_Code"
                    Caption ="Code"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =1140
                    Top =1665
                    Width =1260
                    Height =240
                    Name ="EtiDéb_Situation"
                    Caption ="Situation"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =120
                    Top =105
                    Width =471
                    Height =486
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
                        0x00000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =3765
                    Top =1665
                    Width =1245
                    Height =240
                    FontWeight =400
                    Name ="EtiDéb_Enc_Prin_m"
                    Caption ="Prin Encaissé"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =2460
                    Top =1665
                    Width =1245
                    Height =240
                    Name ="EtiDéb_Mont_Prin_m"
                    Caption ="Principal"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =86
                    TextFontFamily =34
                    Left =690
                    Top =105
                    Width =472
                    Height =487
                    FontSize =8
                    TabIndex =1
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
                        0x0fffffff000088880fffffff080888880fffffff008888880000000008888888 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    ListRows =2
                    Left =5580
                    Top =400
                    Width =1131
                    TabIndex =2
                    Name ="CmbMonnaie"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Monnaie.Mon_Id, Monnaie.Mon_Nom FROM MONNAIE; "
                    ColumnWidths ="0;1134"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =1
                            Left =5585
                            Top =60
                            Width =1110
                            Height =240
                            FontSize =7
                            BackColor =8421504
                            Name ="EtiMonnaie"
                            Caption ="Monnaie :"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1262
                    Top =780
                    Width =1131
                    Height =240
                    TabIndex =3
                    Name ="TxtDateDebut"
                    Format ="Short Date"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =780
                            Width =1020
                            Height =240
                            FontSize =7
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiRech_Fac_DateDebut"
                            Caption ="Entre le :"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1262
                    Top =1050
                    Width =1131
                    TabIndex =4
                    Name ="TxtDateFin"
                    Format ="Short Date"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =1050
                            Width =1020
                            Height =240
                            FontSize =7
                            FontWeight =400
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiRech_Fac_DateFin"
                            Caption ="Et le :"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =82
                    Left =4695
                    Top =1245
                    Width =2055
                    Height =315
                    TabIndex =5
                    Name ="CmdOK"
                    Caption ="&Rechercher"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =5055
                    Top =1665
                    Width =1185
                    Height =240
                    FontWeight =400
                    Name ="EtiDéb_Date_Cré"
                    Caption ="Date Cré"
                End
            End
        End
        Begin Section
            Height =285
            BackColor =12632256
            Name ="Détail0"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6420
                    Width =320
                    Height =245
                    FontSize =6
                    FontWeight =400
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
                    TabIndex =1
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =1155
                    Top =15
                    Width =1281
                    Height =240
                    TabIndex =2
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Situation"
                    ControlSource ="Déb_Situation"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =3765
                    Top =15
                    Width =1281
                    Height =240
                    TabIndex =3
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Enc_Prin_m"
                    ControlSource ="Déb_Enc_Prin_m1"
                    Format ="Standard"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =2460
                    Top =15
                    Width =1281
                    Height =240
                    TabIndex =4
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Mont_Prin_m"
                    ControlSource ="Déb_Mont_Prin_m1"
                    Format ="Standard"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5070
                    Width =1176
                    TabIndex =5
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtDéb_Date_Cré"
                    ControlSource ="Déb_Date_Cré"
                    Format ="Short Date"
                    StatusBarText ="Debiteur Date De Creance"
                End
            End
        End
        Begin FormFooter
            Height =226
            BackColor =12632256
            Name ="PiedFormulaire2"
        End
    End
End
CodeBehindForm
' See "FrmDebiteurEncaissementIncorrect.cls"
