Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    FilterOn = NotDefault
    DateGrouping =1
    GridY =10
    Width =9066
    ItemSuffix =186
    Left =90
    Top =2025
    Tag ="2"
    Filter ="(Mon_Id=2 AND Fra_Grille = 1)"
    RecSrcDt = Begin
        0x7680eb63d1e9e240
    End
    RecordSource ="SELECT DISTINCTROW FRAIS_GRILLE.Fra_Grille_Libellé, FRAIS.*, Monnaie.Mon_Id, Mon"
        "naie.Mon_Nom FROM Monnaie, FRAIS_GRILLE INNER JOIN FRAIS ON FRAIS_GRILLE.Fra_Gri"
        "lle=FRAIS.Fra_Grille; "
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
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
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
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            KeepTogether =1
            ControlSource ="Fra_Grille"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1530
            Name ="EntêteEtat3"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Top =112
                    Width =9063
                    Height =480
                    FontSize =14
                    Name ="EtiFraisRelance"
                    Caption ="Frais de Recouvrement  ( Tranche \\ Niveau de Relance )"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    Top =648
                    Width =2880
                    Height =225
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Short Date"
                End
                Begin Line
                    BorderWidth =2
                    Top =56
                    Width =9063
                    Name ="Ligne26"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =49
                    Left =2925
                    Top =645
                    Width =846
                    Height =225
                    TabIndex =1
                    Name ="Relance1"
                    ControlSource ="=DLookUp(\"Tx_Frais\",\"TAUX\",\"[Tx_N°] = 1\")"
                    FontName ="Courier New"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =49
                    Left =3825
                    Top =645
                    Width =786
                    Height =225
                    TabIndex =2
                    Name ="Relance2"
                    ControlSource ="=DLookUp(\"Tx_Frais\",\"TAUX\",\"[Tx_N°] = 2\")"
                    FontName ="Courier New"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =49
                    Left =4725
                    Top =645
                    Width =786
                    Height =225
                    TabIndex =3
                    Name ="Relance3"
                    ControlSource ="=DLookUp(\"Tx_Frais\",\"TAUX\",\"[Tx_N°] = 3\")"
                    FontName ="Courier New"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =49
                    Left =5625
                    Top =645
                    Width =786
                    Height =225
                    TabIndex =4
                    Name ="Relance4"
                    ControlSource ="=DLookUp(\"Tx_Frais\",\"TAUX\",\"[Tx_N°] = 4\")"
                    FontName ="Courier New"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =49
                    Left =6480
                    Top =645
                    Width =786
                    Height =225
                    TabIndex =5
                    Name ="Relance5"
                    ControlSource ="=DLookUp(\"Tx_Frais\",\"TAUX\",\"[Tx_N°] = 5\")"
                    FontName ="Courier New"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =49
                    Left =7380
                    Top =645
                    Width =786
                    Height =225
                    TabIndex =6
                    Name ="Relance6"
                    ControlSource ="=DLookUp(\"Tx_Frais\",\"TAUX\",\"[Tx_N°] = 6\")"
                    FontName ="Courier New"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =49
                    Left =8280
                    Top =645
                    Width =786
                    Height =225
                    TabIndex =7
                    Name ="Relance7"
                    ControlSource ="=DLookUp(\"Tx_Frais\",\"TAUX\",\"[Tx_N°] = 7\")"
                    FontName ="Courier New"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1455
            Name ="EntêteGroupe6"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    Top =465
                    Width =9060
                    FontSize =10
                    FontWeight =700
                    Name ="Fra_Grille_Libellé"
                    ControlSource ="Fra_Grille_Libellé"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =870
                    Top =1170
                    Width =1170
                    Height =225
                    Name ="Texte17"
                    Caption ="1"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =2040
                    Top =1170
                    Width =1170
                    Height =225
                    Name ="Texte172"
                    Caption ="2"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =3210
                    Top =1170
                    Width =1170
                    Height =225
                    Name ="Texte173"
                    Caption ="3"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =4380
                    Top =1170
                    Width =1170
                    Height =225
                    Name ="Texte174"
                    Caption ="4"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =5550
                    Top =1170
                    Width =1170
                    Height =225
                    Name ="Texte175"
                    Caption ="5"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =6720
                    Top =1170
                    Width =1170
                    Height =225
                    Name ="Texte176"
                    Caption ="6"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =7890
                    Top =1170
                    Width =1170
                    Height =225
                    Name ="Texte177"
                    Caption ="7"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =870
                    Top =870
                    Width =8190
                    Height =300
                    FontSize =9
                    Name ="Texte178"
                    Caption ="Niveau de Relance"
                End
                Begin Line
                    BorderWidth =2
                    Top =450
                    Width =9063
                    Name ="Ligne182"
                End
                Begin Line
                    BorderWidth =2
                    Top =765
                    Width =9063
                    Name ="Ligne183"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    Top =870
                    Width =855
                    FontWeight =700
                    TabIndex =1
                    Name ="Texte185"
                    ControlSource ="Mon_Nom"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =4650
            Name ="Détail1"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Width =291
                    Height =225
                    Name ="Fra_Tranche1"
                    ControlSource ="Fra_Tranche1"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =240
                    Width =291
                    Height =225
                    TabIndex =1
                    Name ="Fra_Tranche2"
                    ControlSource ="Fra_Tranche2"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =480
                    Width =291
                    Height =225
                    TabIndex =2
                    Name ="Fra_Tranche3"
                    ControlSource ="Fra_Tranche3"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =720
                    Width =291
                    Height =225
                    TabIndex =3
                    Name ="Fra_Tranche4"
                    ControlSource ="Fra_Tranche4"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =960
                    Width =291
                    Height =225
                    TabIndex =4
                    Name ="Fra_Tranche5"
                    ControlSource ="Fra_Tranche5"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =1200
                    Width =291
                    Height =225
                    TabIndex =5
                    Name ="Fra_Tranche6"
                    ControlSource ="Fra_Tranche6"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =1440
                    Width =291
                    Height =225
                    TabIndex =6
                    Name ="Fra_Tranche7"
                    ControlSource ="Fra_Tranche7"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =1680
                    Width =291
                    Height =225
                    TabIndex =7
                    Name ="Fra_Tranche8"
                    ControlSource ="Fra_Tranche8"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =1920
                    Width =291
                    Height =225
                    TabIndex =8
                    Name ="Fra_Tranche9"
                    ControlSource ="Fra_Tranche9"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =2145
                    Width =291
                    Height =225
                    TabIndex =9
                    Name ="Fra_Tranche10"
                    ControlSource ="Fra_Tranche10"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =2385
                    Width =291
                    Height =225
                    TabIndex =10
                    Name ="Fra_Tranche11"
                    ControlSource ="Fra_Tranche11"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =2625
                    Width =291
                    Height =225
                    TabIndex =11
                    Name ="Fra_Tranche12"
                    ControlSource ="Fra_Tranche12"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =2865
                    Width =291
                    Height =225
                    TabIndex =12
                    Name ="Fra_Tranche13"
                    ControlSource ="Fra_Tranche13"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =3105
                    Width =291
                    Height =225
                    TabIndex =13
                    Name ="Fra_Tranche14"
                    ControlSource ="Fra_Tranche14"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =3345
                    Width =291
                    Height =225
                    TabIndex =14
                    Name ="Fra_Tranche15"
                    ControlSource ="Fra_Tranche15"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =3585
                    Width =291
                    Height =225
                    TabIndex =15
                    Name ="Fra_Tranche16"
                    ControlSource ="Fra_Tranche16"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =3825
                    Width =291
                    Height =225
                    TabIndex =16
                    Name ="Fra_Tranche17"
                    ControlSource ="Fra_Tranche17"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Top =4065
                    Width =291
                    Height =225
                    TabIndex =17
                    Name ="Fra_Tranche18"
                    ControlSource ="Fra_Tranche18"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin Rectangle
                    BackStyle =1
                    Left =105
                    Width =325
                    Height =4300
                    Name ="Boîte181"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte41"
                    Caption ="1"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =240
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte43"
                    Caption ="2"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =480
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte45"
                    Caption ="3"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =720
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte47"
                    Caption ="4"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =960
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte49"
                    Caption ="5"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =1200
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte51"
                    Caption ="6"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =1440
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte53"
                    Caption ="7"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =1680
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte55"
                    Caption ="8"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =1921
                    Width =360
                    Height =225
                    BackColor =12632256
                    Name ="Texte57"
                    Caption ="9"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =2145
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte61"
                    Caption ="10"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =2385
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte63"
                    Caption ="11"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =2625
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte65"
                    Caption ="12"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =2865
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte67"
                    Caption ="13"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =3105
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte69"
                    Caption ="14"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =3345
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte71"
                    Caption ="15"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =3585
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte73"
                    Caption ="16"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =3825
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte75"
                    Caption ="17"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    Left =435
                    Top =4065
                    Width =360
                    Height =240
                    BackColor =12632256
                    Name ="Texte77"
                    Caption ="18"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Width =1176
                    Height =240
                    TabIndex =18
                    Name ="11"
                    ControlSource ="=[Fra_Tranche1]+[Fra_Tranche1]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl11"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =240
                    Width =1176
                    Height =240
                    TabIndex =19
                    Name ="21"
                    ControlSource ="=[Fra_Tranche2]+[Fra_Tranche2]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl21"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =480
                    Width =1176
                    Height =240
                    TabIndex =20
                    Name ="31"
                    ControlSource ="=[Fra_Tranche3]+[Fra_Tranche3]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl31"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =720
                    Width =1176
                    Height =240
                    TabIndex =21
                    Name ="41"
                    ControlSource ="=[Fra_Tranche4]+[Fra_Tranche4]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl41"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =960
                    Width =1176
                    Height =240
                    TabIndex =22
                    Name ="51"
                    ControlSource ="=[Fra_Tranche5]+[Fra_Tranche5]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl51"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =1200
                    Width =1176
                    Height =240
                    TabIndex =23
                    Name ="61"
                    ControlSource ="=[Fra_Tranche6]+[Fra_Tranche6]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl61"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =1440
                    Width =1176
                    Height =240
                    TabIndex =24
                    Name ="71"
                    ControlSource ="=[Fra_Tranche7]+[Fra_Tranche7]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl71"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =1680
                    Width =1176
                    Height =240
                    TabIndex =25
                    Name ="81"
                    ControlSource ="=[Fra_Tranche8]+[Fra_Tranche8]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl81"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =1920
                    Width =1176
                    Height =240
                    TabIndex =26
                    Name ="91"
                    ControlSource ="=[Fra_Tranche9]+[Fra_Tranche9]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl91"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =2145
                    Width =1176
                    Height =240
                    TabIndex =27
                    Name ="101"
                    ControlSource ="=[Fra_Tranche10]+[Fra_Tranche10]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl101"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =2385
                    Width =1176
                    Height =240
                    TabIndex =28
                    Name ="111"
                    ControlSource ="=[Fra_Tranche11]+[Fra_Tranche11]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl111"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =2625
                    Width =1176
                    Height =240
                    TabIndex =29
                    Name ="121"
                    ControlSource ="=[Fra_Tranche12]+[Fra_Tranche12]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl121"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =2865
                    Width =1176
                    Height =240
                    TabIndex =30
                    Name ="131"
                    ControlSource ="=[Fra_Tranche13]+[Fra_Tranche13]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl131"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =3105
                    Width =1176
                    Height =240
                    TabIndex =31
                    Name ="141"
                    ControlSource ="=[Fra_Tranche14]+[Fra_Tranche14]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl141"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =3345
                    Width =1176
                    Height =240
                    TabIndex =32
                    Name ="151"
                    ControlSource ="=[Fra_Tranche15]+[Fra_Tranche15]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl151"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =3585
                    Width =1176
                    Height =240
                    TabIndex =33
                    Name ="161"
                    ControlSource ="=[Fra_Tranche16]+[Fra_Tranche16]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl161"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =3825
                    Width =1176
                    Height =240
                    TabIndex =34
                    Name ="171"
                    ControlSource ="=[Fra_Tranche17]+[Fra_Tranche17]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl171"
                End
                Begin TextBox
                    TextAlign =3
                    Left =870
                    Top =4065
                    Width =1176
                    Height =240
                    TabIndex =35
                    Name ="181"
                    ControlSource ="=[Fra_Tranche18]+[Fra_Tranche18]*[Relance1]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl181"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Width =1176
                    Height =240
                    TabIndex =36
                    Name ="12"
                    ControlSource ="=[11]+[11]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl12"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =240
                    Width =1176
                    Height =240
                    TabIndex =37
                    Name ="22"
                    ControlSource ="=[21]+[21]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl22"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =480
                    Width =1176
                    Height =240
                    TabIndex =38
                    Name ="32"
                    ControlSource ="=[31]+[31]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl32"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =720
                    Width =1176
                    Height =240
                    TabIndex =39
                    Name ="42"
                    ControlSource ="=[41]+[41]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl42"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =960
                    Width =1176
                    Height =240
                    TabIndex =40
                    Name ="52"
                    ControlSource ="=[51]+[51]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl52"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =1200
                    Width =1176
                    Height =240
                    TabIndex =41
                    Name ="62"
                    ControlSource ="=[61]+[61]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl62"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =1440
                    Width =1176
                    Height =240
                    TabIndex =42
                    Name ="72"
                    ControlSource ="=[71]+[71]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl72"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =1680
                    Width =1176
                    Height =240
                    TabIndex =43
                    Name ="82"
                    ControlSource ="=[81]+[81]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl82"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =1920
                    Width =1176
                    Height =240
                    TabIndex =44
                    Name ="92"
                    ControlSource ="=[91]+[91]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl92"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =2145
                    Width =1176
                    Height =240
                    TabIndex =45
                    Name ="102"
                    ControlSource ="=[101]+[101]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl102"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =2385
                    Width =1176
                    Height =240
                    TabIndex =46
                    Name ="112"
                    ControlSource ="=[111]+[111]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl112"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =2625
                    Width =1176
                    Height =240
                    TabIndex =47
                    Name ="122"
                    ControlSource ="=[121]+[121]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl122"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =2865
                    Width =1176
                    Height =240
                    TabIndex =48
                    Name ="132"
                    ControlSource ="=[131]+[131]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl132"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =3105
                    Width =1176
                    Height =240
                    TabIndex =49
                    Name ="142"
                    ControlSource ="=[141]+[141]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl142"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =3345
                    Width =1176
                    Height =240
                    TabIndex =50
                    Name ="152"
                    ControlSource ="=[151]+[151]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl152"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =3585
                    Width =1176
                    Height =240
                    TabIndex =51
                    Name ="162"
                    ControlSource ="=[161]+[161]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl162"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =3825
                    Width =1176
                    Height =240
                    TabIndex =52
                    Name ="172"
                    ControlSource ="=[171]+[171]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl172"
                End
                Begin TextBox
                    TextAlign =3
                    Left =2040
                    Top =4065
                    Width =1176
                    Height =240
                    TabIndex =53
                    Name ="182"
                    ControlSource ="=[181]+[181]*[Relance2]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl182"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Width =1176
                    Height =240
                    TabIndex =54
                    Name ="13"
                    ControlSource ="=[12]+[12]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl13"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =240
                    Width =1176
                    Height =240
                    TabIndex =55
                    Name ="23"
                    ControlSource ="=[22]+[22]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl23"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =480
                    Width =1176
                    Height =240
                    TabIndex =56
                    Name ="33"
                    ControlSource ="=[32]+[32]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl33"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =720
                    Width =1176
                    Height =240
                    TabIndex =57
                    Name ="43"
                    ControlSource ="=[42]+[42]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl43"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =960
                    Width =1176
                    Height =240
                    TabIndex =58
                    Name ="53"
                    ControlSource ="=[52]+[52]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl53"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =1200
                    Width =1176
                    Height =240
                    TabIndex =59
                    Name ="63"
                    ControlSource ="=[62]+[62]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl63"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =1440
                    Width =1176
                    Height =240
                    TabIndex =60
                    Name ="73"
                    ControlSource ="=[72]+[72]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl73"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =1680
                    Width =1176
                    Height =240
                    TabIndex =61
                    Name ="83"
                    ControlSource ="=[82]+[82]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl83"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =1920
                    Width =1176
                    Height =240
                    TabIndex =62
                    Name ="93"
                    ControlSource ="=[92]+[92]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl93"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =2145
                    Width =1176
                    Height =240
                    TabIndex =63
                    Name ="103"
                    ControlSource ="=[102]+[102]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl103"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =2385
                    Width =1176
                    Height =240
                    TabIndex =64
                    Name ="113"
                    ControlSource ="=[112]+[112]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl113"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =2625
                    Width =1176
                    Height =240
                    TabIndex =65
                    Name ="123"
                    ControlSource ="=[122]+[122]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl123"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =2865
                    Width =1176
                    Height =240
                    TabIndex =66
                    Name ="133"
                    ControlSource ="=[132]+[132]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl133"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =3105
                    Width =1176
                    Height =240
                    TabIndex =67
                    Name ="143"
                    ControlSource ="=[142]+[142]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl143"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =3345
                    Width =1176
                    Height =240
                    TabIndex =68
                    Name ="153"
                    ControlSource ="=[152]+[152]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl153"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =3585
                    Width =1176
                    Height =240
                    TabIndex =69
                    Name ="163"
                    ControlSource ="=[162]+[162]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl163"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =3825
                    Width =1176
                    Height =240
                    TabIndex =70
                    Name ="173"
                    ControlSource ="=[172]+[172]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl173"
                End
                Begin TextBox
                    TextAlign =3
                    Left =3210
                    Top =4065
                    Width =1176
                    Height =240
                    TabIndex =71
                    Name ="183"
                    ControlSource ="=[182]+[182]*[Relance3]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl183"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Width =1176
                    Height =240
                    TabIndex =72
                    Name ="14"
                    ControlSource ="=[13]+[13]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl14"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =240
                    Width =1176
                    Height =240
                    TabIndex =73
                    Name ="24"
                    ControlSource ="=[23]+[23]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl24"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =480
                    Width =1176
                    Height =240
                    TabIndex =74
                    Name ="34"
                    ControlSource ="=[33]+[33]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl34"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =720
                    Width =1176
                    Height =240
                    TabIndex =75
                    Name ="44"
                    ControlSource ="=[43]+[43]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl44"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =960
                    Width =1176
                    Height =240
                    TabIndex =76
                    Name ="54"
                    ControlSource ="=[53]+[53]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl54"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =1200
                    Width =1176
                    Height =240
                    TabIndex =77
                    Name ="64"
                    ControlSource ="=[63]+[63]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl64"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =1440
                    Width =1176
                    Height =240
                    TabIndex =78
                    Name ="74"
                    ControlSource ="=[73]+[73]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl74"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =1680
                    Width =1176
                    Height =240
                    TabIndex =79
                    Name ="84"
                    ControlSource ="=[83]+[83]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl84"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =1920
                    Width =1176
                    Height =240
                    TabIndex =80
                    Name ="94"
                    ControlSource ="=[93]+[93]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl94"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =2145
                    Width =1176
                    Height =240
                    TabIndex =81
                    Name ="104"
                    ControlSource ="=[103]+[103]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl104"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =2385
                    Width =1176
                    Height =240
                    TabIndex =82
                    Name ="114"
                    ControlSource ="=[113]+[113]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl114"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =2625
                    Width =1176
                    Height =240
                    TabIndex =83
                    Name ="124"
                    ControlSource ="=[123]+[123]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl124"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =2865
                    Width =1176
                    Height =240
                    TabIndex =84
                    Name ="134"
                    ControlSource ="=[133]+[133]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl134"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =3105
                    Width =1176
                    Height =240
                    TabIndex =85
                    Name ="144"
                    ControlSource ="=[143]+[143]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl144"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =3345
                    Width =1176
                    Height =240
                    TabIndex =86
                    Name ="154"
                    ControlSource ="=[153]+[153]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl154"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =3585
                    Width =1176
                    Height =240
                    TabIndex =87
                    Name ="164"
                    ControlSource ="=[163]+[163]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl164"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =3825
                    Width =1176
                    Height =240
                    TabIndex =88
                    Name ="174"
                    ControlSource ="=[173]+[173]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl174"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4380
                    Top =4065
                    Width =1176
                    Height =240
                    TabIndex =89
                    Name ="184"
                    ControlSource ="=[183]+[183]*[Relance4]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl184"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Width =1176
                    Height =240
                    TabIndex =90
                    Name ="15"
                    ControlSource ="=[14]+[14]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl15"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =240
                    Width =1176
                    Height =240
                    TabIndex =91
                    Name ="25"
                    ControlSource ="=[24]+[24]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl25"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =480
                    Width =1176
                    Height =240
                    TabIndex =92
                    Name ="35"
                    ControlSource ="=[34]+[34]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl35"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =720
                    Width =1176
                    Height =240
                    TabIndex =93
                    Name ="45"
                    ControlSource ="=[44]+[44]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl45"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =960
                    Width =1176
                    Height =240
                    TabIndex =94
                    Name ="55"
                    ControlSource ="=[54]+[54]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl55"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =1200
                    Width =1176
                    Height =240
                    TabIndex =95
                    Name ="65"
                    ControlSource ="=[64]+[64]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl65"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =1440
                    Width =1176
                    Height =240
                    TabIndex =96
                    Name ="75"
                    ControlSource ="=[74]+[74]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl75"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =1680
                    Width =1176
                    Height =240
                    TabIndex =97
                    Name ="85"
                    ControlSource ="=[84]+[84]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl85"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =1920
                    Width =1176
                    Height =240
                    TabIndex =98
                    Name ="95"
                    ControlSource ="=[94]+[94]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl95"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =2145
                    Width =1176
                    Height =240
                    TabIndex =99
                    Name ="105"
                    ControlSource ="=[104]+[104]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl105"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =2385
                    Width =1176
                    Height =240
                    TabIndex =100
                    Name ="115"
                    ControlSource ="=[114]+[114]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl115"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =2625
                    Width =1176
                    Height =240
                    TabIndex =101
                    Name ="125"
                    ControlSource ="=[124]+[124]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl125"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =2865
                    Width =1176
                    Height =240
                    TabIndex =102
                    Name ="135"
                    ControlSource ="=[134]+[134]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl135"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =3105
                    Width =1176
                    Height =240
                    TabIndex =103
                    Name ="145"
                    ControlSource ="=[144]+[144]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl145"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =3345
                    Width =1176
                    Height =240
                    TabIndex =104
                    Name ="155"
                    ControlSource ="=[154]+[154]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl155"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =3585
                    Width =1176
                    Height =240
                    TabIndex =105
                    Name ="165"
                    ControlSource ="=[164]+[164]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl165"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =3825
                    Width =1176
                    Height =240
                    TabIndex =106
                    Name ="175"
                    ControlSource ="=[174]+[174]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl175"
                End
                Begin TextBox
                    TextAlign =3
                    Left =5550
                    Top =4065
                    Width =1176
                    Height =240
                    TabIndex =107
                    Name ="185"
                    ControlSource ="=[184]+[184]*[Relance5]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl185"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Width =1176
                    Height =240
                    TabIndex =108
                    Name ="16"
                    ControlSource ="=[15]+[15]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl16"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =240
                    Width =1176
                    Height =240
                    TabIndex =109
                    Name ="26"
                    ControlSource ="=[25]+[25]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl26"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =480
                    Width =1176
                    Height =240
                    TabIndex =110
                    Name ="36"
                    ControlSource ="=[35]+[35]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl36"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =720
                    Width =1176
                    Height =240
                    TabIndex =111
                    Name ="46"
                    ControlSource ="=[45]+[45]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl46"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =960
                    Width =1176
                    Height =240
                    TabIndex =112
                    Name ="56"
                    ControlSource ="=[55]+[55]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl56"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =1200
                    Width =1176
                    Height =240
                    TabIndex =113
                    Name ="66"
                    ControlSource ="=[65]+[65]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl66"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =1440
                    Width =1176
                    Height =240
                    TabIndex =114
                    Name ="76"
                    ControlSource ="=[75]+[75]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl76"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =1680
                    Width =1176
                    Height =240
                    TabIndex =115
                    Name ="86"
                    ControlSource ="=[85]+[85]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl86"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =1920
                    Width =1176
                    Height =240
                    TabIndex =116
                    Name ="96"
                    ControlSource ="=[95]+[95]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl96"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =2145
                    Width =1176
                    Height =240
                    TabIndex =117
                    Name ="106"
                    ControlSource ="=[105]+[105]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl106"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =2385
                    Width =1176
                    Height =240
                    TabIndex =118
                    Name ="116"
                    ControlSource ="=[115]+[115]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl116"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =2625
                    Width =1176
                    Height =240
                    TabIndex =119
                    Name ="126"
                    ControlSource ="=[125]+[125]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl126"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =2865
                    Width =1176
                    Height =240
                    TabIndex =120
                    Name ="136"
                    ControlSource ="=[135]+[135]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl136"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =3105
                    Width =1176
                    Height =240
                    TabIndex =121
                    Name ="146"
                    ControlSource ="=[145]+[145]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl146"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =3345
                    Width =1176
                    Height =240
                    TabIndex =122
                    Name ="156"
                    ControlSource ="=[155]+[155]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl156"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =3585
                    Width =1176
                    Height =240
                    TabIndex =123
                    Name ="166"
                    ControlSource ="=[165]+[165]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl166"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =3825
                    Width =1176
                    Height =240
                    TabIndex =124
                    Name ="176"
                    ControlSource ="=[175]+[175]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl176"
                End
                Begin TextBox
                    TextAlign =3
                    Left =6720
                    Top =4065
                    Width =1176
                    Height =240
                    TabIndex =125
                    Name ="186"
                    ControlSource ="=[185]+[185]*[Relance6]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl186"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Width =1176
                    Height =240
                    TabIndex =126
                    Name ="17"
                    ControlSource ="=[16]+[16]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl17"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =240
                    Width =1176
                    Height =240
                    TabIndex =127
                    Name ="27"
                    ControlSource ="=[26]+[26]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl27"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =480
                    Width =1176
                    Height =240
                    TabIndex =128
                    Name ="37"
                    ControlSource ="=[36]+[36]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl37"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =720
                    Width =1176
                    Height =240
                    TabIndex =129
                    Name ="47"
                    ControlSource ="=[46]+[46]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl47"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =960
                    Width =1176
                    Height =240
                    TabIndex =130
                    Name ="57"
                    ControlSource ="=[56]+[56]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl57"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =1200
                    Width =1176
                    Height =240
                    TabIndex =131
                    Name ="67"
                    ControlSource ="=[66]+[66]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl67"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =1440
                    Width =1176
                    Height =240
                    TabIndex =132
                    Name ="77"
                    ControlSource ="=[76]+[76]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl77"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =1680
                    Width =1176
                    Height =240
                    TabIndex =133
                    Name ="87"
                    ControlSource ="=[86]+[86]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl87"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =1920
                    Width =1176
                    Height =240
                    TabIndex =134
                    Name ="97"
                    ControlSource ="=[96]+[96]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl97"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =2145
                    Width =1176
                    Height =240
                    TabIndex =135
                    Name ="107"
                    ControlSource ="=[106]+[106]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl107"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =2385
                    Width =1176
                    Height =240
                    TabIndex =136
                    Name ="117"
                    ControlSource ="=[116]+[116]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl117"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =2625
                    Width =1176
                    Height =240
                    TabIndex =137
                    Name ="127"
                    ControlSource ="=[126]+[126]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl127"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =2865
                    Width =1176
                    Height =240
                    TabIndex =138
                    Name ="137"
                    ControlSource ="=[136]+[136]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl137"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =3105
                    Width =1176
                    Height =240
                    TabIndex =139
                    Name ="147"
                    ControlSource ="=[146]+[146]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl147"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =3345
                    Width =1176
                    Height =240
                    TabIndex =140
                    Name ="157"
                    ControlSource ="=[156]+[156]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl157"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =3585
                    Width =1176
                    Height =240
                    TabIndex =141
                    Name ="167"
                    ControlSource ="=[166]+[166]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl167"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =3825
                    Width =1176
                    Height =240
                    TabIndex =142
                    Name ="177"
                    ControlSource ="=[176]+[176]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl177"
                End
                Begin TextBox
                    TextAlign =3
                    Left =7890
                    Top =4065
                    Width =1176
                    Height =240
                    TabIndex =143
                    Name ="187"
                    ControlSource ="=[186]+[186]*[Relance7]/100"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    EventProcPrefix ="Ctl187"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =187
                    Top =1305
                    Width =180
                    Height =1590
                    FontSize =9
                    Name ="Texte179"
                    Caption ="T\015\012r\015\012a\015\012n\015\012c\015\012h\015\012e"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="PiedEtat4"
        End
    End
End
CodeBehindForm
' See "EtaFraisRelance.cls"
