Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DateGrouping =1
    GridY =10
    Width =9070
    ItemSuffix =79
    Left =630
    Top =1770
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xfaaf108414efe240
    End
    RecordSource ="SELECT DISTINCTROW SelListeEncaissement.Enc_Id, SelListeEncaissement.Enc_Type, S"
        "elListeEncaissement.Enc_Date, SelListeEncaissement.Enc_Montant_m1, SelListeEncai"
        "ssement.Enc_Montant_m2, SelListeEncaissement.Enc_Montant_mo, SelListeEncaissemen"
        "t.Déb_Code, SelListeEncaissement.Déb_Nom, SelListeEncaissement.ETy_Libellé, Coun"
        "t(SelListeEncaissement_1.Enc_Id) AS Numero, Int(([Numero]-1)/50) AS Groupe FROM "
        "SelListeEncaissement INNER JOIN SelListeEncaissement AS SelListeEncaissement_1 O"
        "N SelListeEncaissement.Enc_Type=SelListeEncaissement_1.Enc_Type WHERE (((SelList"
        "eEncaissement_1.Déb_Code)=SelListeEncaissement.Déb_Code) And ((SelListeEncaissem"
        "ent_1.Enc_Id)<=SelListeEncaissement.Enc_Id)) Or (((SelListeEncaissement_1.Déb_Co"
        "de)<SelListeEncaissement.Déb_Code)) GROUP BY SelListeEncaissement.Enc_Id, SelLis"
        "teEncaissement.Enc_Type, SelListeEncaissement.Enc_Date, SelListeEncaissement.Enc"
        "_Montant_m1, SelListeEncaissement.Enc_Montant_m2, SelListeEncaissement.Enc_Monta"
        "nt_mo, SelListeEncaissement.Déb_Code, SelListeEncaissement.Déb_Nom, SelListeEnca"
        "issement.ETy_Libellé; "
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
            ControlSource ="Enc_Type"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Groupe"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Enc_Montant_mo"
        End
        Begin BreakLevel
            ControlSource ="Déb_Code"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="EntêteÉtat"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="EntêteGroupe0"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =9033
                    Name ="Ligne26"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =2497
            BreakLevel =1
            Name ="EntêteGroupe1"
            Begin
                Begin Rectangle
                    BackStyle =1
                    Top =2100
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
                    Top =2160
                    Width =2010
                    Height =285
                    FontSize =10
                    Name ="EtiEnc_Type"
                    Caption ="Mode de règlement :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    Left =3030
                    Top =2160
                    Width =5430
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="TxtETy_Libellé"
                    ControlSource ="ETy_Libellé"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    Left =2085
                    Top =2160
                    Width =915
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtEnc_Type"
                    ControlSource ="=\"(\" & [Enc_Type] & \")\""
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Width =9033
                    Height =480
                    FontSize =14
                    Name ="EtiListeEncaissement"
                    Caption ="Bordereau de remises de chèques"
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
                    TabIndex =2
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Top =1133
                    Width =5895
                    Height =615
                    FontSize =12
                    FontWeight =400
                    Name ="EtiListeEncaissement1"
                    Caption ="Remise faite par GREF\015\012Compte à créditer : ..............................."
                        "............"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =340
            BreakLevel =2
            Name ="EntêteGroupe"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =6075
                    Width =510
                    Height =285
                    FontSize =10
                    Name ="Texte42"
                    Caption ="Date"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =7770
                    Width =855
                    Height =285
                    FontSize =10
                    Name ="EtiEnc_Montant"
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
                    Left =22
                    Width =1125
                    Height =285
                    FontSize =10
                    Name ="EtiDéb_Code"
                    Caption ="Réf."
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =1140
                    Width =525
                    Height =285
                    FontSize =10
                    Name ="EtiDéb_Nom"
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
                    Left =5805
                    Top =15
                    Width =1065
                    Height =240
                    TabIndex =2
                    Name ="TxtEnc_Date"
                    ControlSource ="Enc_Date"
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
                    Name ="TxtEnc_Montant_m"
                    ControlSource ="=IIf([Enc_Montant_mo]=1,Format([Enc_Montant_m1],\"#,##0.00\"\" F\"\"\"),Format(["
                        "Enc_Montant_m2],\"#,##0.00 €\"))"
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
            Height =300
            BreakLevel =2
            Name ="PiedGroupe"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Left =6330
                    Top =15
                    Width =675
                    Height =285
                    FontSize =10
                    Name ="EtiTotalEnc_Montant_m"
                    Caption ="Total :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7005
                    Top =15
                    Width =1635
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="TxtTotalEnc_Montant_m"
                    ControlSource ="=IIf([Enc_Montant_mo]=1,Format(Sum([Enc_Montant_m1]),\"#,##0.00\"\" F\"\"\"),For"
                        "mat(Sum([Enc_Montant_m2]),\"#,##0.00 €\"))"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            BreakLevel =1
            Name ="PiedGroupe2"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="PiedÉtat"
        End
    End
End
CodeBehindForm
' See "EtaListeEncaissement.cls"
