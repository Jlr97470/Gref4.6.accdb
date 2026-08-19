Operation =1
Option =0
Where ="(((CODE_GESTION.CG_Code_Action) In (1,2)) AND ((ADRESSE_DEBITEUR.AD_TypeAdresse_"
    "ID)=DLookUp('Adresse_Courrier','PARAMETRE_GREF')))"
Begin InputTables
    Name ="CODE_GESTION"
    Name ="CLIENT"
    Name ="DEBITEUR"
    Name ="ADRESSE_DEBITEUR"
End
Begin OutputColumns
    Alias ="CG_New_Code_Action"
    Expression ="CInt(Switch([CG_Code_Action]=1 And ([Déb_niv_relance]+IIf(IsNull([Déb_Date_der_r"
        "el]),0,1))<=[Cli_Nb_Relance],1,[CG_Code_Action]=1 And ([Déb_niv_relance]+IIf(IsN"
        "ull([Déb_Date_der_rel]),0,1))>[Cli_Nb_Relance] And IsNull([AD_Tel]) And IsNull(["
        "AD_Tel2]),3,[CG_Code_Action]=1 And ([Déb_niv_relance]+IIf(IsNull([Déb_Date_der_r"
        "el]),0,1))>[Cli_Nb_Relance] And (Not IsNull([AD_Tel]) Or Not IsNull([AD_Tel2])),"
        "2,[CG_Code_Action]=2,3))"
    Alias ="Déb_New_niv_relance"
    Expression ="CInt(Switch([CG_Code_Action]=1 And ([Déb_niv_relance]+IIf(IsNull([Déb_Date_der_r"
        "el]),0,1))<=[Cli_Nb_Relance],([Déb_niv_relance]+IIf(IsNull([Déb_Date_der_rel]),0"
        ",1)),[CG_Code_Action]=1 And ([Déb_niv_relance]+IIf(IsNull([Déb_Date_der_rel]),0,"
        "1))>[Cli_Nb_Relance] And IsNull([AD_Tel]) And IsNull([AD_Tel2]),7,[CG_Code_Actio"
        "n]=1 And ([Déb_niv_relance]+IIf(IsNull([Déb_Date_der_rel]),0,1))>[Cli_Nb_Relance"
        "] And (Not IsNull([AD_Tel]) Or Not IsNull([AD_Tel2])),8,[CG_Code_Action]=2,7))"
    Expression ="DEBITEUR.Déb_Date_proch_rel"
    Alias ="NbDebiteur"
    Expression ="Count(DEBITEUR.Déb_Code)"
End
Begin Joins
    LeftTable ="CLIENT"
    RightTable ="DEBITEUR"
    Expression ="CLIENT.Cli_Code=DEBITEUR.Déb_Code_Client"
    Flag =1
    LeftTable ="CODE_GESTION"
    RightTable ="DEBITEUR"
    Expression ="CODE_GESTION.CG_Code=DEBITEUR.Déb_Situation"
    Flag =1
    LeftTable ="DEBITEUR"
    RightTable ="ADRESSE_DEBITEUR"
    Expression ="DEBITEUR.Déb_Code=ADRESSE_DEBITEUR.AD_Code_Débiteur"
    Flag =1
End
Begin Groups
    Expression ="CInt(Switch([CG_Code_Action]=1 And ([Déb_niv_relance]+IIf(IsNull([Déb_Date_der_r"
        "el]),0,1))<=[Cli_Nb_Relance],1,[CG_Code_Action]=1 And ([Déb_niv_relance]+IIf(IsN"
        "ull([Déb_Date_der_rel]),0,1))>[Cli_Nb_Relance] And IsNull([AD_Tel]) And IsNull(["
        "AD_Tel2]),3,[CG_Code_Action]=1 And ([Déb_niv_relance]+IIf(IsNull([Déb_Date_der_r"
        "el]),0,1))>[Cli_Nb_Relance] And (Not IsNull([AD_Tel]) Or Not IsNull([AD_Tel2])),"
        "2,[CG_Code_Action]=2,3))"
    GroupLevel =0
    Expression ="CInt(Switch([CG_Code_Action]=1 And ([Déb_niv_relance]+IIf(IsNull([Déb_Date_der_r"
        "el]),0,1))<=[Cli_Nb_Relance],([Déb_niv_relance]+IIf(IsNull([Déb_Date_der_rel]),0"
        ",1)),[CG_Code_Action]=1 And ([Déb_niv_relance]+IIf(IsNull([Déb_Date_der_rel]),0,"
        "1))>[Cli_Nb_Relance] And IsNull([AD_Tel]) And IsNull([AD_Tel2]),7,[CG_Code_Actio"
        "n]=1 And ([Déb_niv_relance]+IIf(IsNull([Déb_Date_der_rel]),0,1))>[Cli_Nb_Relance"
        "] And (Not IsNull([AD_Tel]) Or Not IsNull([AD_Tel2])),8,[CG_Code_Action]=2,7))"
    GroupLevel =0
    Expression ="DEBITEUR.Déb_Date_proch_rel"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbMemo "Filter" ="((SelRelanceNbDebiteur.Déb_New_niv_relance=8))"
dbText "Description" ="Relances Nb Debiteurs"
Begin
    Begin
        dbText "Name" ="DEBITEUR.Déb_Date_proch_rel"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Déb_New_Niv_Relance"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="CG_New_Code_Action"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="NbDebiteur"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =7
    Top =30
    Right =1003
    Bottom =350
    Left =-1
    Top =-1
    Right =985
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="CODE_GESTION"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =289
        Bottom =120
        Top =15
        Name ="CLIENT"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =120
        Top =0
        Name ="DEBITEUR"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =560
        Bottom =120
        Top =11
        Name ="ADRESSE_DEBITEUR"
        Name =""
    End
End
