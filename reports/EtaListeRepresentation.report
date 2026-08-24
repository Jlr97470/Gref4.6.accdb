Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    FilterOn = NotDefault
    DateGrouping =1
    GridY =10
    Width =9090
    ItemSuffix =76
    Left =945
    Top =1830
    OnNoData ="[Event Procedure]"
    RecordSource ="SELECT SelListeRepresentation.Déb_Code, SelListeRepresentation.Déb_Nom, SelListe"
        "Representation.Déb_Date_Représentation, SelListeRepresentation.Déb_Repré_Type, S"
        "elListeRepresentation.Déb_Mont_Prin_m1, SelListeRepresentation.Déb_Mont_Prin_m2,"
        " SelListeRepresentation.Déb_Mont_Prin_mo, SelListeRepresentation.ETy_Libellé, Co"
        "unt(SelListeRepresentation_1.Déb_Code) AS Numero, Int(([Numero]-1)/50) AS Groupe"
        " FROM SelListeRepresentation INNER JOIN SelListeRepresentation AS SelListeRepres"
        "entation_1 ON SelListeRepresentation.Déb_Repré_Type=SelListeRepresentation_1.Déb"
        "_Repré_Type WHERE (((SelListeRepresentation_1.Déb_Code)<=SelListeRepresentation."
        "Déb_Code)) GROUP BY SelListeRepresentation.Déb_Code, SelListeRepresentation.Déb_"
        "Nom, SelListeRepresentation.Déb_Date_Représentation, SelListeRepresentation.Déb_"
        "Repré_Type, SelListeRepresentation.Déb_Mont_Prin_m1, SelListeRepresentation.Déb_"
        "Mont_Prin_m2, SelListeRepresentation.Déb_Mont_Prin_mo, SelListeRepresentation.ET"
        "y_Libellé; "
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
            ControlSource ="Déb_Repré_Type"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Groupe"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Déb_Mont_Prin_mo"
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
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =2494
            BreakLevel =1
            Name ="EntêteGroupe2"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Width =9033
                    Height =480
                    FontSize =14
                    Name ="EtiListeRepresentation"
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
                    Name ="TxtDate"
                    ControlSource ="=Now()"
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
                    Name ="EtiListeRepresentation1"
                    Caption ="Remise faite par GREF\015\012Compte à créditer : ..............................."
                        "............"
                End
                Begin Rectangle
                    BackStyle =1
                    Top =2097
                    Width =9014
                    Height =397
                    BackColor =12632256
                    Name ="RecReglement"
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
                    Name ="EtiDéb_Repré_Type"
                    Caption ="Mode de règlement :"
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
                    Name ="TxtETy_Libellé"
                    ControlSource ="ETy_Libellé"
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
                    Name ="TxtDéb_Repré_Type"
                    ControlSource ="=\"(\" & [Déb_Repré_Type] & \")\""
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =340
            BreakLevel =2
            Name ="EntêteGroupe1"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =6075
                    Width =510
                    Height =285
                    FontSize =10
                    Name ="EtiDéb_Date_Représentation"
                    Caption ="Date"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =7770
                    Width =855
                    Height =285
                    FontSize =10
                    Name ="EtiDéb_Mont_Prin_m2"
                    Caption ="Montant"
                End
                Begin Rectangle
                    Top =315
                    Width =9016
                    Height =2
                    Name ="RecListeRepresentation1"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =15
                    Width =1125
                    Height =285
                    FontSize =10
                    Name ="EtiDéb_Code"
                    Caption ="Réf."
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =1133
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
            Height =490
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
                    Width =1200
                    Height =240
                    TabIndex =2
                    Name ="TxtDéb_Date_Représentation"
                    ControlSource ="Déb_Date_Représentation"
                    Format ="Short Date"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7099
                    Top =15
                    Width =1530
                    Height =240
                    TabIndex =3
                    Name ="TxtDéb_Mont_Prin_m"
                    ControlSource ="=IIf([Déb_Mont_Prin_mo]=1,Format([Déb_Mont_Prin_m1],\"#,##0.00\"\" F\"\"\"),Form"
                        "at([Déb_Mont_Prin_m2],\"#,##0.00 €\"))"
                End
                Begin Rectangle
                    BorderLineStyle =3
                    Top =300
                    Width =9016
                    Height =2
                    Name ="RecListeRepresentation"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =566
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
                    Name ="EtiTotalDéb_Mont_Prin_m2"
                    Caption ="Total :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7095
                    Top =15
                    Width =1545
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="TxtTotalDéb_Mont_Prin_m2"
                    ControlSource ="=IIf([Déb_Mont_Prin_mo]=1,Format(Sum([Déb_Mont_Prin_m1]),\"#,##0.00\"\" F\"\"\")"
                        ",Format(Sum([Déb_Mont_Prin_m2]),\"#,##0.00 €\"))"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            BreakLevel =1
            Name ="PiedGroupe0"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="PiedÉtat"
        End
    End
End
CodeBehindForm
' See "EtaListeRepresentation.cls"
