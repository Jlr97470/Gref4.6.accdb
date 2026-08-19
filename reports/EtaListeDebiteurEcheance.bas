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
    ItemSuffix =75
    Left =810
    Top =1890
    OnNoData ="[Event Procedure]"
    Filter ="De_Date_éch<=Date()"
    RecSrcDt = Begin
        0x5e32183f34e8e240
    End
    RecordSource ="SELECT DISTINCTROW OPERATEUR.Ope_Code, OPERATEUR.Ope_Nom, DEBITEUR_ECHEANCE.De_D"
        "ate_éch, DEBITEUR_ECHEANCE.De_Mont_éch_m1, DEBITEUR_ECHEANCE.De_Mont_éch_m2, DEB"
        "ITEUR_ECHEANCE.De_Mont_éch_mo, DEBITEUR.Déb_Code, DEBITEUR.Déb_Nom, DEBITEUR.Déb"
        "_Mont_Prin_m1, DEBITEUR.Déb_Mont_Prin_m2, DEBITEUR.Déb_Mont_Prin_mo FROM OPERATE"
        "UR INNER JOIN (DEBITEUR INNER JOIN DEBITEUR_ECHEANCE ON DEBITEUR.Déb_Code=DEBITE"
        "UR_ECHEANCE.De_Code_Débiteur) ON OPERATEUR.Ope_Code=DEBITEUR_ECHEANCE.De_Ope_Cod"
        "e; "
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
            ControlSource ="Ope_Code"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Déb_Code"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            RepeatSection = NotDefault
            Height =2551
            Name ="EntêteGroupe"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Width =9033
                    Height =480
                    FontSize =14
                    Name ="EtiListeDebiteurEchéance"
                    Caption ="Liste des echéances CB"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    Left =5610
                    Top =49
                    Width =3405
                    Height =285
                    FontSize =10
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Top =1133
                    Width =5895
                    Height =345
                    FontSize =12
                    FontWeight =400
                    Name ="EtiEcheance"
                    Caption ="Echéances faite par GREF"
                End
                Begin Rectangle
                    BackStyle =1
                    Top =2097
                    Width =9014
                    Height =397
                    BackColor =12632256
                    Name ="Boîte73"
                End
                Begin Label
                    BackStyle =0
                    OldBorderStyle =0
                    TextAlign =1
                    Left =45
                    Top =2157
                    Width =2010
                    Height =285
                    FontSize =10
                    Name ="EtiOperateur"
                    Caption ="Operateur :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    Left =3030
                    Top =2157
                    Width =5430
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtOpe_Nom"
                    ControlSource ="Ope_Nom"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    Left =2085
                    Top =2157
                    Width =915
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtOpe_Code"
                    ControlSource ="Ope_Code"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =340
            BreakLevel =1
            Name ="EntêteGroupe1"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =6075
                    Width =510
                    Height =285
                    FontSize =10
                    Name ="EtiDe_Date_éch"
                    Caption ="Date"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =7770
                    Width =855
                    Height =285
                    FontSize =10
                    Name ="EtiDe_Mont_éch_m"
                    Caption ="Montant"
                End
                Begin Rectangle
                    Top =315
                    Width =9016
                    Height =2
                    Name ="Boîte46"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =15
                    Width =1125
                    Height =285
                    FontSize =10
                    Name ="Étiquette76"
                    Caption ="Réf."
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =1133
                    Width =525
                    Height =285
                    FontSize =10
                    Name ="Étiquette77"
                    Caption ="Nom"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =340
            Name ="Détail"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    Left =1170
                    Top =15
                    Width =4500
                    Height =240
                    FontWeight =700
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Top =15
                    Width =1155
                    Height =240
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    Left =5670
                    Top =15
                    Width =1500
                    Height =240
                    TabIndex =2
                    Name ="TxtDe_Date_éch"
                    ControlSource ="De_Date_éch"
                    Format ="Short Date"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7174
                    Top =15
                    Width =1455
                    Height =240
                    TabIndex =3
                    Name ="TxtDe_Mont_éch_m"
                    ControlSource ="=Switch([De_Mont_éch_mo]=1,Format([De_Mont_éch_m1],\"#,##0.00\"\" F\"\"\"),[De_M"
                        "ont_éch_mo]=2,Format([De_Mont_éch_m2],\"#,##0.00 €\"))"
                End
                Begin Rectangle
                    BorderLineStyle =3
                    Top =300
                    Width =9016
                    Height =2
                    Name ="Boîte61"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =0
            BreakLevel =1
            Name ="PiedGroupe"
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            Name ="PiedGroupe0"
        End
    End
End
CodeBehindForm
' See "EtaListeDebiteurEcheance.cls"
