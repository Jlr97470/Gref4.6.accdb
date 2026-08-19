Version =20
VersionRequired =20
Begin Report
    AllowEditing = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    ViewsAllowed =1
    GridX =5
    GridY =5
    Width =8988
    ItemSuffix =49
    Left =4110
    Top =3105
    RecSrcDt = Begin
        0x7002abc374ece240
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
            Height =1080
            Name ="EntêteFormulaire1"
            Begin
                Begin Label
                    BackStyle =0
                    Left =2070
                    Top =840
                    Width =1080
                    Height =240
                    Name ="EtiDéb_Code"
                    Caption ="Code"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =3180
                    Top =840
                    Width =1290
                    Height =240
                    Name ="EtiDéb_Situation"
                    Caption ="Situation"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =5790
                    Top =840
                    Width =1290
                    Height =240
                    Name ="EtiDéb_Enc_Prin_m"
                    Caption ="Prin Encaissé"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =4485
                    Top =840
                    Width =1290
                    Height =240
                    Name ="EtiDéb_Mont_Prin_m"
                    Caption ="Principal"
                End
                Begin Label
                    TextAlign =1
                    Width =8988
                    Height =480
                    FontSize =14
                    BackColor =16777215
                    Name ="Texte23"
                    Caption ="Liste des Débiteurs Incorrects : "
                    FontName ="Arial"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    TextAlign =3
                    Left =5610
                    Top =49
                    Width =3360
                    Height =285
                    FontSize =10
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                    FontName ="Arial"
                End
            End
        End
        Begin Section
            Height =270
            Name ="Détail0"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    BackStyle =0
                    Left =2055
                    Top =15
                    Width =1101
                    Height =240
                    BackColor =12632256
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    Left =3180
                    Top =15
                    Width =1281
                    Height =240
                    TabIndex =1
                    BackColor =12632256
                    Name ="TxtDéb_Situation"
                    ControlSource ="Déb_Situation"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    Left =5790
                    Top =15
                    Width =1281
                    Height =240
                    TabIndex =2
                    BackColor =12632256
                    Name ="TxtDéb_Enc_Prin_m"
                    ControlSource ="=IIf([Déb_MonnaieEnCours]=1,Format([Déb_Enc_Prin_m1],\"#,##0.00\"\" F\"\"\"),For"
                        "mat([Déb_Enc_Prin_m2],\"#,##0.00 €\"))"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    Left =4485
                    Top =15
                    Width =1281
                    Height =240
                    TabIndex =3
                    BackColor =12632256
                    Name ="TxtDéb_Mont_Prin_m"
                    ControlSource ="=IIf([Déb_MonnaieEnCours]=1,Format([Déb_Mont_Prin_m1],\"#,##0.00\"\" F\"\"\"),Fo"
                        "rmat([Déb_Mont_Prin_m2],\"#,##0.00 €\"))"
                End
            End
        End
        Begin FormFooter
            Height =226
            Name ="PiedFormulaire2"
        End
    End
End
CodeBehindForm
' See "EtaListeDebiteurEncaissementIncorrect.cls"
