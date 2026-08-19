Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DateGrouping =1
    GridY =10
    Width =9070
    ItemSuffix =78
    Left =810
    Top =1680
    RecSrcDt = Begin
        0x3ec97f3d10efe240
    End
    RecordSource ="SELECT CODE_GESTION.CG_Code, CODE_GESTION.CG_Code_Groupe, CODE_GESTION.CG_Libell"
        "é, Nz(Count([Déb_Code]),0) AS NbDebiteur, Nz(Sum(SelClientAnalyseSituation.Déb_M"
        "ont_Prin_m1),0) AS Déb_Mont_Prin_m1, Nz(Sum(SelClientAnalyseSituation.Déb_Mont_P"
        "rin_m2),0) AS Déb_Mont_Prin_m2, Nz(Sum(SelClientAnalyseSituation.Déb_Mont_Revers"
        "é_m1),0) AS Déb_Mont_Reversé_m1, Nz(Sum(SelClientAnalyseSituation.Déb_Mont_Rever"
        "sé_m2),0) AS Déb_Mont_Reversé_m2, Nz(Sum(SelClientAnalyseSituation.Déb_Prin_Enc_"
        "m1),0) AS Déb_Prin_Enc_m1, Nz(Sum(SelClientAnalyseSituation.Déb_Prin_Enc_m2),0) "
        "AS Déb_Prin_Enc_m2 FROM CODE_GESTION LEFT JOIN SelClientAnalyseSituation ON CODE"
        "_GESTION.CG_Code=SelClientAnalyseSituation.Déb_Situation GROUP BY CODE_GESTION.C"
        "G_Code, CG_Code_Groupe, CODE_GESTION.CG_Libellé; "
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
            KeepTogether =1
            ControlSource ="CG_Code_Groupe"
        End
        Begin BreakLevel
            ControlSource ="CG_Code"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =2385
            Name ="EntêteEtat"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Top =116
                    Width =6243
                    Height =480
                    FontSize =14
                    Name ="EtiAnalyseSituationClientSynthese"
                    Caption ="Analyse de situation Client (Synthèse)"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    Left =6291
                    Top =170
                    Width =2730
                    Height =285
                    FontSize =10
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =9018
                    Name ="TrtAnalyseSituationClientSynthese"
                End
                Begin Label
                    OldBorderStyle =0
                    Left =180
                    Top =1080
                    Width =2025
                    Height =330
                    FontSize =12
                    FontWeight =400
                    Name ="EtiCli_Nom"
                    Caption ="Référence Client :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =3165
                    Top =1080
                    Width =5430
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtCli_Nom"
                    ControlSource ="=DLookUp(\"Cli_Nom\",\"CLIENT\",\"Cli_Code = '\" & [TxtDéb_Code_Client] & \"'\")"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =2235
                    Top =1080
                    Width =900
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtDéb_Code_Client"
                    ControlSource ="=Forms!FrmClientAnalyseSituation!TxtDéb_Code_Client"
                End
                Begin Label
                    OldBorderStyle =0
                    Left =290
                    Top =1803
                    Width =690
                    Height =330
                    FontSize =12
                    Name ="EtiCG_Code"
                    Caption ="Code"
                End
                Begin Label
                    OldBorderStyle =0
                    Left =1032
                    Top =1809
                    Width =855
                    Height =330
                    FontSize =12
                    Name ="EtiCG_Libellé"
                    Caption ="Libellé"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =3
                    Left =5105
                    Top =1803
                    Width =1005
                    Height =330
                    FontSize =12
                    Name ="EtiNbDebiteur"
                    Caption ="Nombre"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =3
                    Left =6741
                    Top =1809
                    Width =870
                    Height =315
                    FontSize =12
                    Name ="EtiSolde"
                    Caption ="Solde"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =3
                    Left =8150
                    Top =1803
                    Width =390
                    Height =315
                    FontSize =12
                    Name ="EtiPourcentage"
                    Caption ="%"
                End
                Begin Rectangle
                    Left =236
                    Top =1752
                    Width =8561
                    Height =412
                    Name ="RecAnalyseSituationClientSynthese2"
                End
                Begin Rectangle
                    Left =180
                    Top =1695
                    Width =8659
                    Height =526
                    Name ="RecAnalyseSituationClientSynthese1"
                End
                Begin Label
                    OldBorderStyle =0
                    Top =600
                    Width =2475
                    Height =345
                    FontSize =12
                    Name ="EtiMon_Nom"
                    Caption ="Débiteurs traités en :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =2475
                    Top =600
                    Width =1905
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    Name ="TxtMon_Nom"
                    ControlSource ="=DLookUp(\"Mon_Nom\",\"MONNAIE\",\"Mon_Id=\" & Forms!FrmClientAnalyseSituation!C"
                        "mbMonnaie)"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =7
            Name ="EntêteGroupe"
            Begin
                Begin Rectangle
                    Left =255
                    Width =8546
                    Height =7
                    Name ="RecCG_Code_Groupe"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =331
            Name ="Détail"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    Left =285
                    Top =30
                    Width =576
                    FontSize =10
                    Name ="TxtCG_Code"
                    ControlSource ="CG_Code"
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =1032
                    Top =30
                    Width =4071
                    FontSize =10
                    TabIndex =1
                    Name ="TxtCG_Libellé"
                    ControlSource ="CG_Libellé"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =5220
                    Top =30
                    Width =846
                    FontSize =10
                    TabIndex =2
                    Name ="TxtNbDebiteur"
                    ControlSource ="NbDebiteur"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =6200
                    Top =30
                    Width =1446
                    FontSize =10
                    TabIndex =3
                    Name ="TxtSolde"
                    ControlSource ="=[Déb_Mont_Prin_m1]-[Déb_Prin_Enc_m1]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7718
                    Top =30
                    Width =906
                    FontSize =10
                    TabIndex =4
                    Name ="TxtPourcentage"
                    ControlSource ="=IIf([TxtTotal]=0,0,[TxtSolde]/[TxtTotal])"
                    Format ="Percent"
                End
                Begin Rectangle
                    Left =221
                    Width =4
                    Height =331
                    Name ="RecCG_Libellé"
                End
                Begin Rectangle
                    Left =8792
                    Width =4
                    Height =331
                    Name ="RecCG_Code"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =510
            Name ="PiedGroupe"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =5145
                    Top =45
                    Width =921
                    FontSize =10
                    FontWeight =700
                    Name ="TxtNbDebiteurGroupe"
                    ControlSource ="=Sum([NbDebiteur])"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =6140
                    Top =45
                    Width =1506
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtSoldeGroupe"
                    ControlSource ="=Sum([Déb_Mont_Prin_m1]-[Déb_Prin_Enc_m1])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7718
                    Top =45
                    Width =906
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtPourcentageGroupe"
                    ControlSource ="=IIf([TxtTotal]=0,0,[TxtSoldeGroupe]/[TxtTotal])"
                    Format ="Percent"
                End
                Begin Rectangle
                    Left =225
                    Width =8561
                    Height =367
                    Name ="RecCG_Code_Groupe1"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =2820
                    Top =45
                    Width =1731
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="TxtSoldeGroupeMonnaie"
                    ControlSource ="=[TxtSoldeGroupe]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =1109
            Name ="PiedEtat"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    TextAlign =3
                    Left =1695
                    Width =1266
                    Height =225
                    FontSize =7
                    TabIndex =6
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtTotalSolde"
                    ControlSource ="=Sum([Déb_Mont_Prin_m1]-[Déb_Prin_Enc_m1])"
                    Format ="Standard"
                    FontName ="Small Fonts"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    TextAlign =3
                    Left =3015
                    Width =1266
                    Height =225
                    FontSize =7
                    TabIndex =7
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtTotalDéb_Mont_Reversé_m"
                    ControlSource ="=Sum([Déb_Mont_Reversé_m1])+Round(Sum([Déb_Mont_Reversé_m2])*DonnerTauxConvertio"
                        "n(2),2)"
                    Format ="Standard"
                    FontName ="Small Fonts"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    TextAlign =3
                    Left =4320
                    Width =1266
                    Height =225
                    FontSize =7
                    TabIndex =8
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtTotalDéb_Prin_Enc_m"
                    ControlSource ="=Sum([Déb_Prin_Enc_m1])"
                    Format ="Standard"
                    FontName ="Small Fonts"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    TextAlign =3
                    Left =5610
                    Width =1266
                    Height =225
                    FontSize =7
                    TabIndex =9
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtTotal"
                    ControlSource ="=[TxtTotalSolde]+[TxtTotalDéb_Mont_Reversé_m]+([TxtTotalDéb_Prin_Enc_m]-[TxtTota"
                        "lDéb_Mont_Reversé_m])"
                    Format ="Standard"
                    FontName ="Small Fonts"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =6035
                    Top =311
                    Width =1611
                    FontSize =10
                    FontWeight =700
                    Name ="TxtTotalReversé"
                    ControlSource ="=[TxtTotalDéb_Mont_Reversé_m]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin Rectangle
                    Left =225
                    Top =255
                    Width =8561
                    Height =412
                    Name ="RecReversé"
                End
                Begin Label
                    OldBorderStyle =0
                    Left =283
                    Top =283
                    Width =3405
                    Height =330
                    FontSize =12
                    Name ="EtiTotalReversé"
                    Caption ="Déjà  reversé au client"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =6050
                    Top =753
                    Width =1596
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtTotalResteAVerser"
                    ControlSource ="=[TxtTotalDéb_Prin_Enc_m]-[TxtTotalDéb_Mont_Reversé_m]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin Rectangle
                    Left =225
                    Top =697
                    Width =8561
                    Height =412
                    Name ="RecResteAVerser"
                End
                Begin Label
                    OldBorderStyle =0
                    Left =283
                    Top =737
                    Width =3405
                    Height =330
                    FontSize =12
                    Name ="EtiTotalResteAVerser"
                    Caption ="Reste à verser au client "
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7755
                    Top =315
                    Width =861
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtTotalReverséPourcentage"
                    ControlSource ="=IIf([TxtTotal]=0,0,[TxtTotalReversé]/[TxtTotal])"
                    Format ="Percent"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7770
                    Top =750
                    Width =861
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="TxtTotalResteAVerserPourcentage"
                    ControlSource ="=IIf([TxtTotal]=0,0,[TxtTotalResteAVerser]/[TxtTotal])"
                    Format ="Percent"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =3945
                    Top =330
                    Width =1611
                    FontWeight =700
                    TabIndex =4
                    Name ="TxtTotalReverséMonnaie"
                    ControlSource ="=[TxtTotalDéb_Mont_Reversé_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =3960
                    Top =772
                    Width =1596
                    FontWeight =700
                    TabIndex =5
                    Name ="TxtTotalResteAVerserMonnaie"
                    ControlSource ="=([TxtTotalDéb_Prin_Enc_m]-[TxtTotalDéb_Mont_Reversé_m])*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
            End
        End
    End
End
CodeBehindForm
' See "EtaAnalyseSituationClientSynthese.cls"
