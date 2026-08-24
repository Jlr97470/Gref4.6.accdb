Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8985
    DatasheetFontHeight =10
    ItemSuffix =17
    Left =1065
    Top =1665
    DatasheetGridlinesColor =12632256
    Filter ="(Déb_Code='8827203038')"
    RecSrcDt = Begin
        0x62cf270131eee240
    End
    RecordSource ="SELECT DEBITEUR.Déb_Qualité, DEBITEUR.Déb_Code, DEBITEUR.Déb_Nom, DEBITEUR.Déb_M"
        "onnaieEnCours, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb_Mont_Prin_m2, DEBITEUR.Dé"
        "b_Date_Cré, DEBITEUR.Déb_Typ_Cré, DEBITEUR.Déb_Nom_Banque, DEBITEUR.Déb_Num_Chèq"
        "ue, DEBITEUR.Déb_Date_Chèque, CLIENT.Cli_Nom, ADRESSE_DEBITEUR.AD_Adr1, ADRESSE_"
        "DEBITEUR.AD_Adr2, ADRESSE_DEBITEUR.AD_Adr3, ADRESSE_DEBITEUR.AD_Adr4, ADRESSE_DE"
        "BITEUR.AD_CodePostal, ADRESSE_DEBITEUR.AD_Ville, ADRESSE_DEBITEUR.AD_Pays FROM ("
        "CLIENT INNER JOIN DEBITEUR ON CLIENT.Cli_Code=DEBITEUR.Déb_Code_Client) INNER JO"
        "IN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Code=ADRESSE_DEBITEUR.AD_Code_Débiteur WHERE"
        " (((ADRESSE_DEBITEUR.AD_TypeAdresse_ID)=DLookUp(\"Adresse_Courrier\",\"PARAMETRE"
        "_GREF\"))); "
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =11340
            Name ="Détail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4260
                    Top =1905
                    Width =3981
                    Height =330
                    ColumnWidth =540
                    FontSize =12
                    Name ="TxtDéb_Qualité"
                    ControlSource ="=[Déb_Qualité] & \" \" & [Déb_Nom]"
                    StatusBarText ="Debiteur Qualite"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4260
                    Top =2235
                    Width =3981
                    Height =330
                    FontSize =12
                    TabIndex =1
                    Name ="TxtAD_Adr1"
                    ControlSource ="AD_Adr1"
                    StatusBarText ="Adresse Debiteur Ligne 1"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4260
                    Top =2565
                    Width =3981
                    Height =330
                    FontSize =12
                    TabIndex =2
                    Name ="TxtAD_Adr2"
                    ControlSource ="AD_Adr2"
                    StatusBarText ="Adresse Debiteur Ligne 1"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4260
                    Top =2895
                    Width =3981
                    Height =330
                    FontSize =12
                    TabIndex =3
                    Name ="TxtAD_Adr3"
                    ControlSource ="AD_Adr3"
                    StatusBarText ="Adresse Debiteur Ligne 1"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4260
                    Top =3225
                    Width =3981
                    Height =330
                    FontSize =12
                    TabIndex =4
                    Name ="TxtAD_Adr4"
                    ControlSource ="AD_Adr4"
                    StatusBarText ="Adresse Debiteur Ligne 1"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4260
                    Top =3555
                    Width =3981
                    Height =330
                    FontSize =12
                    TabIndex =5
                    Name ="TxtAD_CodePostal"
                    ControlSource ="=[AD_CodePostal] & \" \" & [AD_Ville] & \" \" & [AD_Pays]"
                    StatusBarText ="Adresse Debiteur Code Postal"
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =285
                    Top =6240
                    Width =8511
                    Height =1005
                    FontSize =12
                    TabIndex =6
                    Name ="TxtAttestationPaiement1"
                    ControlSource ="=\"Nous, Cabinet GREF, élisant domicile 4 rue Plumet 75015 PARIS, certifions avo"
                        "ir reçu en date du \" & DMax(\"Enc_Date\",\"ENCAISSEMENT\",\"Enc_Code_Débiteur='"
                        "\" & [Déb_Code] & \"'\") & \" de \" & [Déb_Qualité] & \" \" & [Déb_Nom] & \" la "
                        "somme de : \""
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3975
                    Top =7365
                    Height =330
                    FontSize =12
                    TabIndex =7
                    Name ="TxtDéb_Mont_Prin_m"
                    ControlSource ="=IIf([Déb_MonnaieEnCours]=1,Format([Déb_Mont_Prin_m1],\"#,##0.00\"\" F\"\"\"),Fo"
                        "rmat([Déb_Mont_Prin_m2],\"#,##0.00 €\"))"
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =285
                    Top =8505
                    Width =8511
                    Height =1065
                    FontSize =12
                    TabIndex =8
                    Name ="TxtAttestationPaiement2"
                    ControlSource ="=\"en règlement du chèque \" & [Déb_Nom_Banque] & \" N° \" & [Déb_Num_Chèque] & "
                        "\" émis le \" & [Déb_Date_Chèque] & \" au profit de la Société \" & [Cli_Nom]"
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextFontFamily =34
                    Left =285
                    Top =9870
                    Width =4815
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="EtiAttestationPaiement1"
                    Caption ="Pour servir et valoir ce que de droit."
                End
                Begin Label
                    TextFontFamily =34
                    Left =6240
                    Top =10770
                    Width =2265
                    Height =330
                    FontSize =12
                    Name ="EtiAttestationPaiement2"
                    Caption ="La Direction"
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4260
                    Top =4770
                    Width =3966
                    Height =330
                    FontSize =12
                    TabIndex =9
                    Name ="TxtDate"
                    ControlSource ="=\"Paris, Le \" & Date()"
                End
            End
        End
    End
End
