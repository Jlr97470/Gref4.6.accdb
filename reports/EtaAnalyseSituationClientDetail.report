Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    FilterOn = NotDefault
    DateGrouping =1
    GridY =10
    Width =9033
    ItemSuffix =67
    Left =2370
    Top =1725
    OnNoData ="[Event Procedure]"
    Filter ="[Déb_Code_Client]='88961' AND [Déb_MonnaieEnCours]=2"
    RecSrcDt = Begin
        0xbd00fbc074ece240
    End
    RecordSource ="SELECT DISTINCTROW DEBITEUR.Déb_Code_Client, DEBITEUR.Déb_Code, DEBITEUR.Déb_Nom"
        ", DEBITEUR.Déb_Réf1_Client, DEBITEUR.Déb_Réf2_Client, DEBITEUR.Déb_Date_Cré, DEB"
        "ITEUR.Déb_Typ_Cré, Round(DEBITEUR.Déb_Mont_Prin_m1,2) AS Déb_Mont_Prin_m1, Round"
        "(DEBITEUR.Déb_Mont_Prin_m2,2) AS Déb_Mont_Prin_m2, Round(Nz(DSum(\"Enc_Prin_m1\""
        ",\"ENCAISSEMENT\",\"Enc_Code_Débiteur='\" & [Déb_Code] & \"'\"),0),2) AS Déb_Enc"
        "_Prin_m1, Round(Nz(DSum(\"Enc_Prin_m2\",\"ENCAISSEMENT\",\"Enc_Code_Débiteur='\""
        " & [Déb_Code] & \"'\"),0),2) AS Déb_Enc_Prin_m2, Round(Nz(DSum(\"Enc_Prin_m1\",\""
        "ENCAISSEMENT\",\"Enc_Code_Débiteur='\" & [Déb_Code] & \"' AND Enc_Montant_mo=1 A"
        "ND Enc_Situation Is Not Null\"),0),2) AS Déb_Mont_Reversé_m1, Round(Nz(DSum(\"En"
        "c_Prin_m2\",\"ENCAISSEMENT\",\"Enc_Code_Débiteur='\" & [Déb_Code] & \"' AND Enc_"
        "Montant_mo=2 AND Enc_Situation Is Not Null\"),0),2) AS Déb_Mont_Reversé_m2, DEBI"
        "TEUR.Déb_MonnaieEnCours, CODE_GESTION.CG_Code, CODE_GESTION.CG_Libellé, CODE_GES"
        "TION.CG_Code_Groupe, CODE_GESTION.CG_Code_Action FROM CODE_GESTION INNER JOIN DE"
        "BITEUR ON CODE_GESTION.CG_Code=DEBITEUR.Déb_Situation; "
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
            ControlSource ="Déb_Code_Client"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Déb_Code_Client"
        End
        Begin BreakLevel
            ControlSource ="Déb_Nom"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1395
            Name ="EntêteGroupe"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Top =56
                    Width =9033
                    Height =480
                    FontSize =14
                    Name ="EtiAnalyseSituationClientDetail"
                    Caption ="Analyse de situation Client (Détail)"
                End
                Begin TextBox
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
                    Name ="TrtAnalyseSituationClientDetail"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =2880
                    Top =1080
                    Width =5430
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtCli_Nom"
                    ControlSource ="=DLookUp(\"Cli_Nom\",\"CLIENT\",\"Cli_Code='\" & [Déb_Code_Client] & \"'\")"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =1905
                    Top =1080
                    Width =945
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtDéb_Code_Client"
                    ControlSource ="Déb_Code_Client"
                    Begin
                        Begin Label
                            OldBorderStyle =0
                            Top =1080
                            Width =1875
                            Height =300
                            FontSize =11
                            FontWeight =400
                            Name ="EtiCli_Code"
                            Caption ="Référence Client :"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =2475
                    Top =540
                    Width =1905
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    Name ="TxtMon_Nom"
                    ControlSource ="=Dlookup(\"Mon_Nom\",\"MONNAIE\",\"Mon_Id=\" & [Déb_MonnaieEnCours])"
                    Begin
                        Begin Label
                            OldBorderStyle =0
                            Left =45
                            Top =540
                            Width =2475
                            Height =345
                            FontSize =12
                            Name ="EtiMon_Nom"
                            Caption ="Débiteurs traités en :"
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            RepeatSection = NotDefault
            Height =617
            BreakLevel =1
            Name ="EntêteGroupe1"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =4514
                    Width =945
                    Height =585
                    FontSize =10
                    Name ="EtiDéb_Date_Cré"
                    Caption ="Date\015\012Créance"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =5640
                    Top =135
                    Width =945
                    Height =285
                    FontSize =10
                    Name ="EtiPrincipal"
                    Caption ="Principal"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =6750
                    Top =135
                    Width =1125
                    Height =285
                    FontSize =10
                    Name ="EtiRèglement"
                    Caption ="Règlement"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =8335
                    Top =135
                    Width =645
                    Height =285
                    FontSize =10
                    Name ="EtiSolde"
                    Caption ="Solde"
                End
                Begin Rectangle
                    Top =615
                    Width =9016
                    Height =2
                    Name ="TrtAnalyseSituationClientDetail2"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =1077
            Name ="Détail"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    Top =480
                    Width =3975
                    Height =210
                    Name ="TxtDéb_Typ_Cré"
                    ControlSource ="Déb_Typ_Cré"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Top =690
                    Width =285
                    Height =210
                    TabIndex =1
                    Name ="TxtCG_Code"
                    ControlSource ="CG_Code"
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =1170
                    Width =3495
                    Height =255
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Width =1155
                    Height =255
                    FontWeight =700
                    TabIndex =3
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =285
                    Top =690
                    Width =3660
                    Height =210
                    TabIndex =4
                    Name ="TxtCG_Libellé"
                    ControlSource ="CG_Libellé"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    Left =4440
                    Top =283
                    Width =1065
                    Height =285
                    TabIndex =5
                    Name ="TxtDéb_Date_Cré"
                    ControlSource ="Déb_Date_Cré"
                    Format ="Short Date"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    Left =5483
                    Top =283
                    Width =1200
                    Height =285
                    TabIndex =6
                    Name ="TxtPrincipal"
                    ControlSource ="=IIf([CG_Code_Action] In (6,7),0,[Déb_Mont_Prin_m1]-[Déb_Mont_reversé_m1]-[Déb_M"
                        "ont_reversé_m2]*DonnerTauxConvertion(2))"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7832
                    Top =283
                    Width =1185
                    Height =285
                    TabIndex =7
                    Name ="TxtSolde"
                    ControlSource ="=IIf([CG_Code_Action] In (6,7),0,[Déb_Mont_Prin_m1]-[Déb_Mont_Reversé_m1]-[Déb_M"
                        "ont_Reversé_m2]*DonnerTauxConvertion(2))"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    Left =6690
                    Top =285
                    Width =1140
                    Height =285
                    TabIndex =8
                    Name ="TxtRèglement"
                    ControlSource ="=IIf([CG_Code_Action]=7,0,[Déb_Mont_Reversé_m1]+[Déb_Mont_Reversé_m2]*DonnerTaux"
                        "Convertion(2))"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =495
                    Top =255
                    Width =1695
                    Height =225
                    TabIndex =9
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
                    TabIndex =10
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
                    Top =1005
                    Width =9016
                    Height =2
                    Name ="TrtAnalyseSituationClientDetail1"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =365
            BreakLevel =1
            Name ="PiedGroupe"
            Begin
                Begin Rectangle
                    Width =9016
                    Height =2
                    Name ="TrtAnalyseSituationClientDetail3"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7153
                    Top =50
                    Width =1845
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="TxtTotal"
                    ControlSource ="=Sum(IIf([CG_Code_Action] In (1,2,3,4,5,9,10),[Déb_Mont_Prin_m1]-[Déb_Enc_Prin_m"
                        "1],0))"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    Begin
                        Begin Label
                            OldBorderStyle =0
                            TextAlign =1
                            Top =50
                            Width =2325
                            Height =315
                            FontSize =12
                            Name ="EtiTotal"
                            Caption ="Total Client"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =4395
                    Top =45
                    Width =1845
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtTotalMonnaie"
                    ControlSource ="=[TxtTotal]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            Name ="PiedGroupe1"
        End
    End
End
CodeBehindForm
' See "EtaAnalyseSituationClientDetail.cls"
