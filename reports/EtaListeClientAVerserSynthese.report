Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    FilterOn = NotDefault
    DateGrouping =1
    GridX =5
    GridY =5
    Width =9066
    ItemSuffix =71
    Left =2670
    Top =2355
    OnNoData ="[Event Procedure]"
    Filter ="Enc_Date>=#1/1/2005# AND Enc_Date<=#12/31/2005#"
    RecSrcDt = Begin
        0x1fdbbcb5b5ece240
    End
    RecordSource ="SELECT DISTINCTROW CLIENT.Cli_Code, CLIENT.Cli_Nom, ENCAISSEMENT.Enc_Date, Sum(I"
        "If([Enc_Montant_mo]=1,[Enc_Prin_m1],0)) AS ResteAverser_m1, Sum(IIf([Enc_Montant"
        "_mo]=2,[Enc_Prin_m2],0)) AS ResteAverser_m2 FROM ENCAISSEMENT_TYPE INNER JOIN (C"
        "LIENT INNER JOIN (DEBITEUR INNER JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code=ENCAISSE"
        "MENT.Enc_Code_Débiteur) ON CLIENT.Cli_Code=DEBITEUR.Déb_Code_Client) ON ENCAISSE"
        "MENT_TYPE.ETy_Code=ENCAISSEMENT.Enc_Type WHERE (((ENCAISSEMENT.Enc_Situation) Is"
        " Null) AND ((ENCAISSEMENT_TYPE.ETy_Mode)<>2) AND ((ENCAISSEMENT.Enc_Prin_m1)<>0)"
        ") GROUP BY CLIENT.Cli_Code, CLIENT.Cli_Nom, ENCAISSEMENT.Enc_Date; "
    OnOpen ="[Event Procedure]"
    Begin
        Begin Label
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Rectangle
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="Cli_Code"
        End
        Begin FormHeader
            Height =1927
            Name ="EntêteEtat"
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Top =491
                    Width =5943
                    Height =480
                    FontSize =14
                    FontWeight =700
                    Name ="Texte23"
                    Caption ="Liste des Clients"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =6066
                    Top =545
                    Width =2955
                    Height =285
                    FontSize =10
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                End
                Begin Line
                    BorderWidth =2
                    Top =435
                    Width =9018
                    Name ="TrtListeClientAVerserSynthese"
                End
                Begin Rectangle
                    BackStyle =0
                    Top =1814
                    Width =9030
                    Height =0
                    Name ="RecListeClientAVerserSynthese"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =5101
                    Top =1215
                    Width =1155
                    Height =570
                    FontSize =10
                    Name ="EtiCli_Date_der_reg"
                    Caption ="Dernier\015\012versement"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =6236
                    Top =1215
                    Width =1395
                    Height =570
                    FontSize =10
                    Name ="EtiResteAVerser_m1"
                    Caption ="Reste à verser\015\012en Francs"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =7586
                    Top =1215
                    Width =1425
                    Height =570
                    FontSize =10
                    Name ="EtiResteAVerser_m2"
                    Caption ="Reste à verser\015\012en Euros"
                End
            End
        End
        Begin Section
            Height =0
            Name ="Détail"
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =285
            Name ="PiedGroupe1"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Width =1026
                    Height =285
                    FontSize =10
                    Name ="TxtCli_Code"
                    ControlSource ="Cli_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =1020
                    Width =4056
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =5100
                    Width =1161
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="TxtCli_Date_der_reg"
                    ControlSource ="=DLast(\"Reg_Date\",\"REGLEMENT\",\"Reg_Cli_Code='\" & [Cli_Code] & \"'\")"
                    Format ="Short Date"
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =6270
                    Width =1356
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="TxtResteAVerser_m1"
                    ControlSource ="=Sum([ResteAverser_m1])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Locked = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =7815
                    Width =1191
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="TxtResteAVerser_m2"
                    ControlSource ="=Sum([ResteAverser_m2])"
                    Format ="#,##0.00 €"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =7635
                    Width =195
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Eti+"
                    Caption ="+"
                    EventProcPrefix ="Eti_"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =1020
            Name ="PiedEtat1"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =2205
                    Top =675
                    Width =1896
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="TxtTotal_m1"
                    ControlSource ="=[TxtTotalResteAVerser_m1]+[TxtTotalResteAVerser_m2]*DonnerTauxConvertion(2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =15
                            Top =675
                            Width =2145
                            Height =330
                            FontSize =12
                            FontWeight =700
                            Name ="EtiTotal_m1"
                            Caption ="Total Général :"
                        End
                    End
                End
                Begin Line
                    Left =90
                    Top =570
                    Width =8970
                    Name ="RecListeClientAVerserSynthese1"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =5790
                    Top =120
                    Width =1836
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtTotalResteAVerser_m1"
                    ControlSource ="=Sum([ResteAVerser_m1])"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =7620
                    Top =120
                    Width =1446
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtTotalResteAVerser_m2"
                    ControlSource ="=Sum([ResteAVerser_m2])"
                    Format ="#,##0.00 €"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =4095
                    Top =675
                    Width =570
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="EtiTotal_m2"
                    Caption ="soit "
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =4680
                    Top =675
                    Width =2256
                    Height =330
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    Name ="TxtTotal_m2"
                    ControlSource ="=[TxtTotalResteAVerser_m2]+[TxtTotalResteAVerser_m1]*DonnerTauxConvertion(1)"
                    Format ="#,##0.00 €"
                End
            End
        End
    End
End
CodeBehindForm
' See "EtaListeClientAVerserSynthese.cls"
