Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DateGrouping =1
    GridY =10
    Width =9070
    ItemSuffix =90
    Left =2535
    Top =1980
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x8006aecdadece240
    End
    RecordSource ="SELECT DISTINCTROW FACTURE.Fac_Date, CLIENT_GROUPE.CG_Nom, FACTURE.Fac_Num, FACT"
        "URE.Fac_Honoraire, FACTURE.Fac_TVA, FACTURE.Fac_Montant_Direct_m1, FACTURE.Fac_M"
        "ontant_Direct_m2, FACTURE.Fac_Montant_Encaisser_m1, FACTURE.Fac_Montant_Encaisse"
        "r_m2, FACTURE.Fac_Montant_mo FROM (CLIENT_GROUPE INNER JOIN CLIENT ON CLIENT_GRO"
        "UPE.CG_Id=CLIENT.Cli_Groupe) INNER JOIN FACTURE ON CLIENT.Cli_Code=FACTURE.Fac_C"
        "odeCli WHERE (((FACTURE.Fac_Editer)=True)); "
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
            ControlSource ="Fac_Date"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="CG_Nom"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Fac_Num"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="EntêteGroupe1"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            RepeatSection = NotDefault
            Height =3165
            BreakLevel =1
            Name ="EntêteGroupe"
            Begin
                Begin Label
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    Left =2535
                    Top =900
                    Width =4005
                    Height =375
                    FontSize =14
                    Name ="EtiRecapitulatifFactures"
                    Caption ="RECAPITULATIF FACTURES"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    Left =2910
                    Top =1800
                    Width =3255
                    Height =375
                    FontSize =14
                    Name ="EtiTVA"
                    Caption ="TVA FR 53 378 409 171"
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1140
                    Top =1365
                    Width =6795
                    Height =345
                    FontSize =14
                    FontWeight =700
                    Name ="TxtCG_Nom"
                    ControlSource ="CG_Nom"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Top =2880
                    Width =1425
                    Height =285
                    FontSize =10
                    Name ="EtiFac_Num"
                    Caption ="N°FACTURE"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =7365
                    Top =2880
                    Width =1455
                    Height =285
                    FontSize =10
                    Name ="EtiFac_Montant_mTTC"
                    Caption ="HONO TTC"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =5100
                    Top =2880
                    Width =1455
                    Height =285
                    FontSize =10
                    Name ="EtiFac_Montant_mTVA"
                    Caption ="TVA"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =2505
                    Top =2880
                    Width =1470
                    Height =285
                    FontSize =10
                    Name ="EtiFac_Montant_mHT"
                    Caption ="HONO HT"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            BreakLevel =2
            Name ="EntêteGroupe0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            Name ="Détail"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =4
                    TextAlign =3
                    Left =7365
                    Top =15
                    Width =1455
                    Height =240
                    ColumnOrder =2
                    Name ="TxtFac_Montant_mTTC"
                    ControlSource ="=Round(Round(nz([Fac_Montant_Direct_m1],0),2)*[Fac_Honoraire]/100*(1+[Fac_TVA]),"
                        "2)+Round(Round(nz([Fac_Montant_Encaisser_m1],0),2)*[Fac_Honoraire]/100*(1+[Fac_T"
                        "VA]),2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    Top =15
                    Width =1425
                    Height =240
                    ColumnOrder =1
                    TabIndex =1
                    Name ="TxtFac_Num"
                    ControlSource ="Fac_Num"
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =4
                    TextAlign =3
                    Left =5100
                    Top =15
                    Width =1455
                    Height =240
                    ColumnOrder =0
                    TabIndex =2
                    Name ="TxtFac_Montant_mTVA"
                    ControlSource ="=[TxtFac_Montant_mTTC]*[Fac_TVA]/(1+[Fac_TVA])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =4
                    TextAlign =3
                    Left =2505
                    Top =15
                    Width =1470
                    Height =240
                    ColumnOrder =3
                    TabIndex =3
                    Name ="TxtFac_Montant_mHT"
                    ControlSource ="=[TxtFac_Montant_mTTC]-[TxtFac_Montant_mTVA]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =750
            BreakLevel =1
            Name ="PiedGroupe1"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Width =1702
                    Height =285
                    FontSize =10
                    Name ="EtiCG_NomTotal"
                    Caption ="TOTAL GENERAL"
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =12
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2505
                    Width =1470
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="TxtCG_NomFac_Montant_mHT"
                    ControlSource ="=[TxtCG_NomFac_Montant_mTTC]-[TxtCG_NomFac_Montant_mTVA]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =12
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5100
                    Width =1470
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtCG_NomFac_Montant_mTVA"
                    ControlSource ="=[TxtCG_NomFac_Montant_mTTC]*[Fac_TVA]/(1+[Fac_TVA])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =12
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7365
                    Width =1470
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtCG_NomFac_Montant_mTTC"
                    ControlSource ="=Sum(Round(Round(nz([Fac_Montant_Direct_m1],0),2)*[Fac_Honoraire]/100*(1+[Fac_TV"
                        "A]),2))+Sum(Round(Round(nz([Fac_Montant_Encaisser_m1],0),2)*[Fac_Honoraire]/100*"
                        "(1+[Fac_TVA]),2))"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin Line
                    Width =9060
                    Name ="TrtGC_Nom1"
                End
            End
        End
    End
End
CodeBehindForm
' See "EtaSyntheseFactureSousGroupe.cls"
