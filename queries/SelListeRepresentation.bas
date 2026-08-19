dbMemo "SQL" ="SELECT DEBITEUR.Déb_Code, DEBITEUR.Déb_Nom, DEBITEUR.Déb_Date_Représentation, DE"
    "BITEUR.Déb_Repré_Type, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb_Mont_Prin_m2, DEB"
    "ITEUR.Déb_Mont_Prin_mo, ENCAISSEMENT_TYPE.ETy_Libellé\015\012FROM DEBITEUR INNER"
    " JOIN ENCAISSEMENT_TYPE ON DEBITEUR.Déb_Repré_Type=ENCAISSEMENT_TYPE.ETY_Code\015"
    "\012WHERE Déb_Date_Représentation>=#2/1/2006# AND Déb_Date_Représentation<=#2/28"
    "/2006# AND Déb_Repré_Type In('1','2','3','4','5','6','7','8','9','XXXXX');\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
