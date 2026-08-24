Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    DateGrouping =1
    GridY =10
    Width =9063
    ItemSuffix =36
    Left =90
    Top =1800
    RecSrcDt = Begin
        0xe5625be8d4e9e240
    End
    RecordSource ="SELECT DISTINCTROW TAUX.* FROM TAUX; "
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
            KeepTogether =1
            ControlSource ="Tx_Grille"
        End
        Begin BreakLevel
            ControlSource ="Tx_N°"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =928
            Name ="EntêteEtat3"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Top =112
                    Width =9063
                    Height =480
                    FontSize =14
                    Name ="EtiTauxRelance"
                    Caption ="Taux de Relance"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    Top =648
                    Width =2880
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Medium Date"
                End
                Begin Line
                    BorderWidth =2
                    Top =56
                    Width =9063
                    Name ="Ligne26"
                End
            End
        End
        Begin PageHeader
            Height =195
            Name ="EntêtePage0"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1020
            Name ="EntêteGroupe0"
            Begin
                Begin Label
                    OldBorderStyle =0
                    Left =4860
                    Top =720
                    Width =1530
                    Height =225
                    Name ="EtiCA_Libellé"
                    Caption ="Action"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    Left =765
                    Top =345
                    Width =3375
                    Height =250
                    FontWeight =700
                    Name ="TxtTxG_Grille_Libellé"
                    ControlSource ="=dlookup(\"TxG_Grille_Libellé\",\"TAUX_GRILLE\",\"TxG_Grille=\" & [Tx_Grille])"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    Left =75
                    Top =345
                    Width =630
                    Height =250
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtTx_Grille"
                    ControlSource ="Tx_Grille"
                End
                Begin Label
                    OldBorderStyle =0
                    Left =1485
                    Top =705
                    Width =1455
                    Height =225
                    Name ="EtiRelNom"
                    Caption ="Nom"
                End
                Begin Label
                    OldBorderStyle =0
                    Left =6585
                    Top =705
                    Width =240
                    Height =225
                    Name ="EtiTx_N°"
                    Caption ="N°"
                End
                Begin Label
                    OldBorderStyle =0
                    Left =7185
                    Top =705
                    Width =810
                    Height =225
                    Name ="EtiTx_Délai"
                    Caption ="Délai (j)"
                End
                Begin Label
                    OldBorderStyle =0
                    Left =8100
                    Top =705
                    Width =885
                    Height =225
                    Name ="EtiTx_Frais"
                    Caption ="Frais (%)"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =645
                    Width =9063
                    Height =15
                    Name ="Ligne27"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =1530
                    Top =926
                    Width =7518
                    Name ="Ligne28"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =270
            Name ="Détail1"
            Begin
                Begin TextBox
                    Left =1485
                    Width =3375
                    Height =250
                    Name ="TxtRel_Nom"
                    ControlSource ="=DLookUp(\"Rel_Nom\",\"RELANCE\",\"Rel_Code_Action=\" & [Tx_Code_Action] & \" AN"
                        "D Rel_N°=\" & [Tx_N°])"
                End
                Begin TextBox
                    TextAlign =2
                    Left =6390
                    Width =630
                    Height =250
                    TabIndex =1
                    Name ="TxtTx_N°"
                    ControlSource ="Tx_N°"
                End
                Begin TextBox
                    TextAlign =2
                    Left =7020
                    Width =900
                    Height =250
                    TabIndex =2
                    Name ="TxtTx_Délai"
                    ControlSource ="Tx_Délai"
                    StatusBarText ="Type ou Taille du champ"
                End
                Begin TextBox
                    TextAlign =2
                    Left =7920
                    Width =1068
                    Height =250
                    TabIndex =3
                    Name ="TxtTx_Frais"
                    ControlSource ="Tx_Frais"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4860
                    Width =1536
                    Height =255
                    ColumnWidth =1605
                    TabIndex =4
                    Name ="TxtCA_Libellé"
                    ControlSource ="=Dlookup(\"CA_Libellé\",\"CODE_ACTION\",\"CA_Code=\" & [Tx_Code_Action])"
                    StatusBarText ="Grille Taux Code Action"
                End
            End
        End
        Begin PageFooter
            Height =300
            Name ="PiedPage2"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =505
            Name ="PiedEtat4"
        End
    End
End
