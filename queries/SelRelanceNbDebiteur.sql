SELECT
  CInt(
    Switch(
      [CG_Code_Action] = 1
      And (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ) <= [Cli_Nb_Relance],
      1,
      [CG_Code_Action] = 1
      And (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ) > [Cli_Nb_Relance]
      And IsNull([AD_Tel])
      And IsNull([AD_Tel2]),
      3,
      [CG_Code_Action] = 1
      And (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ) > [Cli_Nb_Relance]
      And (
        Not IsNull([AD_Tel])
        Or Not IsNull([AD_Tel2])
      ),
      2,
      [CG_Code_Action] = 2,
      3
    )
  ) AS CG_New_Code_Action,
  CInt(
    Switch(
      [CG_Code_Action] = 1
      And (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ) <= [Cli_Nb_Relance],
      (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ),
      [CG_Code_Action] = 1
      And (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ) > [Cli_Nb_Relance]
      And IsNull([AD_Tel])
      And IsNull([AD_Tel2]),
      7,
      [CG_Code_Action] = 1
      And (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ) > [Cli_Nb_Relance]
      And (
        Not IsNull([AD_Tel])
        Or Not IsNull([AD_Tel2])
      ),
      8,
      [CG_Code_Action] = 2,
      7
    )
  ) AS Déb_New_niv_relance,
  DEBITEUR.Déb_Date_proch_rel,
  Count(DEBITEUR.Déb_Code) AS NbDebiteur
FROM
  (
    CODE_GESTION
    INNER JOIN (
      CLIENT
      INNER JOIN DEBITEUR ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client
    ) ON CODE_GESTION.CG_Code = DEBITEUR.Déb_Situation
  )
  INNER JOIN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Code = ADRESSE_DEBITEUR.AD_Code_Débiteur
WHERE
  (
    (
      (CODE_GESTION.CG_Code_Action) In (1, 2)
    )
    AND (
      (
        ADRESSE_DEBITEUR.AD_TypeAdresse_ID
      ) = DLookUp(
        'Adresse_Courrier', 'PARAMETRE_GREF'
      )
    )
  )
GROUP BY
  CInt(
    Switch(
      [CG_Code_Action] = 1
      And (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ) <= [Cli_Nb_Relance],
      1,
      [CG_Code_Action] = 1
      And (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ) > [Cli_Nb_Relance]
      And IsNull([AD_Tel])
      And IsNull([AD_Tel2]),
      3,
      [CG_Code_Action] = 1
      And (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ) > [Cli_Nb_Relance]
      And (
        Not IsNull([AD_Tel])
        Or Not IsNull([AD_Tel2])
      ),
      2,
      [CG_Code_Action] = 2,
      3
    )
  ),
  CInt(
    Switch(
      [CG_Code_Action] = 1
      And (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ) <= [Cli_Nb_Relance],
      (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ),
      [CG_Code_Action] = 1
      And (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ) > [Cli_Nb_Relance]
      And IsNull([AD_Tel])
      And IsNull([AD_Tel2]),
      7,
      [CG_Code_Action] = 1
      And (
        [Déb_niv_relance] + IIf(
          IsNull([Déb_Date_der_rel]),
          0,
          1
        )
      ) > [Cli_Nb_Relance]
      And (
        Not IsNull([AD_Tel])
        Or Not IsNull([AD_Tel2])
      ),
      8,
      [CG_Code_Action] = 2,
      7
    )
  ),
  DEBITEUR.Déb_Date_proch_rel;
