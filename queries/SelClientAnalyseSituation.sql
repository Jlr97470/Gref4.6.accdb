SELECT
  DEBITEUR.Déb_Situation,
  DEBITEUR.Déb_Code,
  Round(
    [DEBITEUR].[Déb_Mont_Prin_m1], 2
  ) AS Déb_Mont_Prin_m1,
  Round(
    [DEBITEUR].[Déb_Mont_Prin_m2], 2
  ) AS Déb_Mont_Prin_m2,
  Round(
    Sum(
      IIf(
        [CG_Code_Action] <> 7
        And [Enc_Situation] Is Not Null
        And [Enc_Montant_mo] = 1,
        [Enc_Prin_m1], 0
      )
    ),
    2
  ) AS Déb_Mont_Reversé_m1,
  Round(
    Sum(
      IIf(
        [CG_CODE_Action] <> 7
        And [Enc_Situation] Is Not Null
        And [Enc_Montant_mo] = 2,
        [Enc_Prin_m2], 0
      )
    ),
    2
  ) AS Déb_Mont_Reversé_m2,
  Round(
    Sum(
      IIf(
        [CG_Code_Action] <> 7, [Enc_Prin_m1],
        0
      )
    ),
    2
  ) AS Déb_Prin_Enc_m1,
  Round(
    Sum(
      IIf(
        [CG_Code_Action] <> 7, [Enc_Prin_m2],
        0
      )
    ),
    2
  ) AS Déb_Prin_Enc_m2
FROM
  CODE_GESTION
  INNER JOIN (
    DEBITEUR
    LEFT JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur
  ) ON CODE_GESTION.CG_Code = DEBITEUR.Déb_Situation
WHERE
  [Déb_Code_Client] = '88961'
  AND [Déb_MonnaieEnCours] = 2
GROUP BY
  DEBITEUR.Déb_Situation,
  DEBITEUR.Déb_Code,
  Round(
    [DEBITEUR].[Déb_Mont_Prin_m1], 2
  ),
  Round(
    [DEBITEUR].[Déb_Mont_Prin_m2], 2
  );
