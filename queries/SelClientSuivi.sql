SELECT
  CLIENT.Cli_Code,
  CLIENT.Cli_Nom,
  CLIENT.Cli_Date_contrat,
  CLIENT_TYPE.CT_Nom,
  CLIENT_GROUPE.CG_Nom AS SousGroupe,
  CLIENT_GROUPE_1.CG_Nom AS Groupe,
  CLIENT.Cli_Mont_Hon_m2,
  CLIENT.Cli_Pourcentage,
  CLIENT.Cli_Date_der_cré,
  DLast(
    "Reg_Date",
    "REGLEMENT",
    "Reg_Cli_Code='" & [Cli_Code] & "' AND Reg_Date>=" & ConvDate([Cli_Date_contrat])
  ) AS Cli_Dern_Reg_Date,
  CCur(
    Nz(
      DLast(
        "Reg_Montant_Encaisser_m2",
        "REGLEMENT",
        "Reg_Cli_Code='" & [Cli_Code] & "' AND Reg_Date>=" & ConvDate([Cli_Date_contrat])
      ),
      0
    )
  ) AS Cli_Dern_Reg_Montant,
  CCur(
    Nz(
      Sum(
        IIf(
          [CG_Code_Action] = 7
          And [Déb_Date_Cré] >= [Cli_Date_contrat],
          Round([Déb_Mont_Prin_m2], 2),
          0
        )
      ),
      0
    )
  ) AS Cli_Reg_Direct,
  Nz(
    Sum(
      IIf(
        [CG_Code_Action] = 10,
        Round([Déb_Mont_Prin_m2], 2),
        0
      )
    ),
    0
  ) AS Cli_Representation,
  CCur(
    Nz(
      Sum(
        IIf(
          [CG_Code_Action] <> 7
          And [Déb_Date_Cré] >= [Cli_Date_contrat]
          And [Enc_Situation] Is Not Null,
          [Enc_Prin_m2], 0
        )
      ),
      0
    )
  ) AS Cli_Reverser,
  CCur(
    Nz(
      Sum(
        IIf(
          [CG_Code_Action] <> 7
          And [Déb_Date_Cré] >= [Cli_Date_contrat]
          And [Enc_Situation] Is Null,
          [Enc_Prin_m2], 0
        )
      ),
      0
    )
  ) AS Cli_ResteAVerser,
  CCur(
    Nz(
      Sum(
        IIf(
          [CG_Code_Action] <> 7
          And [Enc_Situation] Is Null,
          [Enc_Prin_m2], 0
        )
      ),
      0
    )
  ) AS Cli_ResteAVerserGlobal
FROM
  (
    CLIENT_GROUPE
    RIGHT JOIN (
      CODE_GESTION
      RIGHT JOIN (
        (
          (
            CLIENT_TYPE
            RIGHT JOIN CLIENT ON CLIENT_TYPE.CT_Id = CLIENT.Cli_Type
          )
          LEFT JOIN (
            DEBITEUR
            LEFT JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur
          ) ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client
        )
        INNER JOIN TMP_REPRESENTANT_CLIENT ON CLIENT.Cli_Code = TMP_REPRESENTANT_CLIENT.TRCCli_Code
      ) ON CODE_GESTION.CG_Code = DEBITEUR.Déb_Situation
    ) ON CLIENT_GROUPE.CG_Id = CLIENT.Cli_Groupe
  )
  LEFT JOIN CLIENT_GROUPE AS CLIENT_GROUPE_1 ON CLIENT_GROUPE.CG_Parent_Id = CLIENT_GROUPE_1.CG_Id
WHERE
  (
    (
      (
        TMP_REPRESENTANT_CLIENT.TRCRep_Code
      ) = 'AP'
    )
    And (
      (
        TMP_REPRESENTANT_CLIENT.TRC_Flag
      ) = True
    )
  )
GROUP BY
  CLIENT.Cli_Code,
  CLIENT.Cli_Nom,
  CLIENT.Cli_Date_contrat,
  CLIENT_TYPE.CT_Nom,
  CLIENT_GROUPE.CG_Nom,
  CLIENT_GROUPE_1.CG_Nom,
  CLIENT.Cli_Mont_Hon_m2,
  CLIENT.Cli_Pourcentage,
  CLIENT.Cli_Date_der_cré,
  DLast(
    "Reg_Date",
    "REGLEMENT",
    "Reg_Cli_Code='" & [Cli_Code] & "' AND Reg_Date>=" & ConvDate([Cli_Date_contrat])
  ),
  Nz(
    DLast(
      "Reg_Montant_Encaisser_m2",
      "REGLEMENT",
      "Reg_Cli_Code='" & [Cli_Code] & "' AND Reg_Date>=" & ConvDate([Cli_Date_contrat])
    ),
    0
  );
