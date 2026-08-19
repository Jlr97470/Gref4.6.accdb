Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    FilterOn = NotDefault
    DateGrouping =1
    GridY =10
    Width =9070
    ItemSuffix =70
    Left =1215
    Top =2535
    OnNoData ="[Event Procedure]"
    Filter ="Cli_Code='cc021' AND Déb_Date_Cré=#2/17/2006#"
    RecSrcDt = Begin
        0x185b2fe02fe9e240
    End
    RecordSource ="SELECT DISTINCTROW CLIENT.Cli_Code, CLIENT.Cli_Nom, DEBITEUR.Déb_Code, DEBITEUR."
        "Déb_Nom, DEBITEUR.Déb_Date_Cré, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb_Mont_Pri"
        "n_m2, DEBITEUR.Déb_Réf1_Client, DEBITEUR.Déb_Réf2_Client, DEBITEUR.Déb_Typ_Cré, "
        "DEBITEUR.Déb_MonnaieEnCours FROM CLIENT INNER JOIN DEBITEUR ON CLIENT.Cli_Code=D"
        "EBITEUR.Déb_Code_Client; "
    OnOpen ="[Event Procedure]"
    Begin
        Begin Label
            OldBorderStyle =1
            FontWeight =700
            FontName ="Arial"
        End
        Begin Rectangle
            BackStyle =0
            Width =850
            Height =280
        End
        Begin Line
            Width =1701
        End
        Begin BoundObjectFrame
            AutoLabel = NotDefault
            Width =1998
            Height =280
            LabelX =-1701
        End
        Begin TextBox
            AutoLabel = NotDefault
            Width =1701
            Height =280
            LabelX =-1701
            FontName ="Arial"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Cli_Code"
        End
        Begin BreakLevel
            ControlSource ="Déb_Code"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =2325
            Name ="EntêteGroupe0"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Top =56
                    Width =9033
                    Height =480
                    FontSize =14
                    Name ="EtiCompteRenduSaisie"
                    Caption ="Compte rendu de saisie : "
                End
                Begin TextBox
                    Visible = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    Left =5610
                    Top =105
                    Width =3405
                    Height =285
                    FontSize =10
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                End
                Begin Line
                    BorderWidth =2
                    Width =9033
                    Name ="Ligne26"
                End
                Begin Label
                    OldBorderStyle =0
                    Top =960
                    Width =1860
                    Height =315
                    FontSize =11
                    FontWeight =400
                    Name ="EtiCli_Code"
                    Caption ="Référence Client :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =2880
                    Top =960
                    Width =5430
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =1905
                    Top =960
                    Width =945
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtCli_Code"
                    ControlSource ="Cli_Code"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =4709
                    Top =1590
                    Width =945
                    Height =585
                    FontSize =10
                    Name ="EtiDéb_Date_Cré"
                    Caption ="Date\015\012Créance"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =5865
                    Top =1740
                    Width =1305
                    Height =285
                    FontSize =10
                    Name ="EtiDéb_Mont_Prin_m"
                    Caption ="Principal"
                End
                Begin Rectangle
                    Top =2205
                    Width =9016
                    Height =2
                    Name ="Boîte46"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =900
            Name ="Détail1"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    Top =480
                    Width =9030
                    Height =210
                    Name ="TxtDéb_Typ_Cré"
                    ControlSource ="Déb_Typ_Cré"
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =1170
                    Width =3495
                    Height =255
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Width =1155
                    Height =255
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    Left =4680
                    Top =180
                    Width =1065
                    Height =285
                    TabIndex =3
                    Name ="TxtDéb_Date_Cré"
                    ControlSource ="Déb_Date_Cré"
                    Format ="Short Date"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =5805
                    Top =180
                    Width =1425
                    Height =285
                    TabIndex =4
                    Name ="TxtDéb_Mont_Prin_m"
                    ControlSource ="Déb_Mont_Prin_m1"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =495
                    Top =255
                    Width =1695
                    Height =225
                    TabIndex =5
                    Name ="TxtDéb_Réf1_Client"
                    ControlSource ="Déb_Réf1_Client"
                End
                Begin Label
                    OldBorderStyle =0
                    Top =255
                    Width =495
                    Height =225
                    FontWeight =400
                    Name ="EtiDéb_Réf1_Client"
                    Caption ="Réf1 :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =2730
                    Top =255
                    Width =1695
                    Height =225
                    TabIndex =6
                    Name ="TxtDéb_Réf2_Client"
                    ControlSource ="Déb_Réf2_Client"
                End
                Begin Label
                    OldBorderStyle =0
                    Left =2235
                    Top =255
                    Width =495
                    Height =225
                    FontWeight =400
                    Name ="EtiDéb_Réf2_Client"
                    Caption ="Réf2 :"
                End
                Begin Rectangle
                    BorderLineStyle =3
                    Top =810
                    Width =9016
                    Height =2
                    Name ="Boîte61"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7740
                    Top =180
                    Width =1290
                    Height =285
                    TabIndex =7
                    Name ="TxtDéb_Mont_Prin_mMonnaie"
                    ControlSource ="=[TxtDéb_Mont_Prin_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =7260
                    Top =180
                    Width =435
                    Height =285
                    FontSize =10
                    FontWeight =400
                    Name ="EtiSoit"
                    Caption ="soit"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =500
            Name ="PiedGroupe1"
            Begin
                Begin Rectangle
                    Top =135
                    Width =9016
                    Height =2
                    Name ="Boîte52"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Left =4200
                    Top =150
                    Width =1545
                    Height =330
                    FontSize =12
                    Name ="Texte53"
                    Caption ="Total Client :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =5805
                    Top =180
                    Width =1425
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="TxtTotalDéb_Mont_Prin_m"
                    ControlSource ="=Round(nz(Sum([TxtDéb_Mont_Prin_m]),0),2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =7260
                    Top =180
                    Width =420
                    Height =285
                    FontSize =10
                    FontWeight =400
                    Name ="EtiTotalSoit"
                    Caption ="soit"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =7728
                    Top =180
                    Width =1290
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtTotalDéb_Mont_Prin_mMonnaie"
                    ControlSource ="=[TxtTotalDéb_Mont_Prin_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
            End
        End
    End
End
CodeBehindForm
' See "EtaCompteRenduSaisie.cls"
