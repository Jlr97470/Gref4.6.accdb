dbMemo "SQL" ="SELECT ENCAISSEMENT.Enc_Id, ENCAISSEMENT.Enc_Type, ENCAISSEMENT.Enc_Date, ENCAIS"
    "SEMENT.Enc_Montant_m1, ENCAISSEMENT.Enc_Montant_m2, ENCAISSEMENT.Enc_Montant_mo,"
    " DEBITEUR.Déb_Code, DEBITEUR.Déb_Nom, ENCAISSEMENT_TYPE.ETy_Libellé\015\012FROM "
    "DEBITEUR INNER JOIN (ENCAISSEMENT INNER JOIN ENCAISSEMENT_TYPE ON ENCAISSEMENT.E"
    "nc_Type=ENCAISSEMENT_TYPE.ETy_Code) ON DEBITEUR.Déb_Code=ENCAISSEMENT.Enc_Code_D"
    "ébiteur\015\012WHERE Enc_Date>=#2/1/2006# AND Enc_Date<=#2/28/2006# AND Enc_Type"
    " In('1','2','3','4','5','6','7','8','9','XXXXX');\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
