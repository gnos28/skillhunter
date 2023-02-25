export const TAB_NAME_CONTRATS = "CONTRATS";
export const TAB_NAME_VERSEMENT = "VERSEMENTS VARIABLE";
export const TAB_NAME_CLIENTS = "CLIENTS";
export const TAB_NAME_VARIABLE = "VARIABLE / COLLABORATEURS";
export const TAB_NAME_COLLAB = "COLLABORATEURS";
export const TAB_NAME_PARAMETRES = "PARAMETRES";
export const TAB_IMPORT_DATA = "IMPORT_DATAS";

export const TAB_CONTRATS_COL_ID = "ID CONTRAT";
export const TAB_CONTRATS_COL_COLLAB = "RÉALISÉ PAR";
export const TAB_CONTRATS_COL_DATE_DEBUT = "DATE DEBUT";
export const TAB_CONTRATS_COL_NB_SEMAINE_GARANTIE = "NB SEMAINES GARANTIE";
export const TAB_CONTRATS_COL_RUPTURE = "(RUPTURE GARANTIE)";
export const TAB_CONTRATS_COL_CLIENT = "CLIENT";
export const TAB_CONTRATS_COL_TYPE = "TYPE CONTRAT";
export const TAB_CONTRATS_COL_CANDIDAT = "CANDIDAT";
export const TAB_CONTRATS_COL_DESCRIPTION = "DESCRIPTION";
export const TAB_CONTRATS_COL_SALAIRE = "SALAIRE CANDIDAT";
export const TAB_CONTRATS_COL_PERCENT = "% CONTRAT";
export const TAB_CONTRATS_COL_DATE_FIN_GARANTIE = "DATE FIN GARANTIE";

export const TAB_CONTRATS_COL_IMPORT_ID = "IMPORT_ID CONTRAT";
export const TAB_CONTRATS_COL_IMPORT_COLLAB = "IMPORT_RÉALISÉ PAR";
export const TAB_CONTRATS_COL_IMPORT_DATE_DEBUT = "IMPORT_DATE DEBUT";
export const TAB_CONTRATS_COL_IMPORT_NB_SEMAINE_GARANTIE =
  "IMPORT_NB SEMAINES GARANTIE";
export const TAB_CONTRATS_COL_IMPORT_RUPTURE = "IMPORT_(RUPTURE GARANTIE)";
export const TAB_CONTRATS_COL_IMPORT_CLIENT = "IMPORT_CLIENT";
export const TAB_CONTRATS_COL_IMPORT_TYPE = "IMPORT_TYPE CONTRAT";
export const TAB_CONTRATS_COL_IMPORT_CANDIDAT = "IMPORT_CANDIDAT";
export const TAB_CONTRATS_COL_IMPORT_DESCRIPTION = "IMPORT_DESCRIPTION";
export const TAB_CONTRATS_COL_IMPORT_SALAIRE = "IMPORT_SALAIRE CANDIDAT";
export const TAB_CONTRATS_COL_IMPORT_PERCENT = "IMPORT_% CONTRAT";

export const COL2KEEP_CONTRATS = [
  TAB_CONTRATS_COL_ID,
  TAB_CONTRATS_COL_COLLAB,
  TAB_CONTRATS_COL_DATE_DEBUT,
  TAB_CONTRATS_COL_NB_SEMAINE_GARANTIE,
  TAB_CONTRATS_COL_RUPTURE,
  "DATE PAIEMENT CLIENT",
  TAB_CONTRATS_COL_CLIENT,
  TAB_CONTRATS_COL_TYPE,
  TAB_CONTRATS_COL_CANDIDAT,
  TAB_CONTRATS_COL_DESCRIPTION,
  TAB_CONTRATS_COL_SALAIRE,
  TAB_CONTRATS_COL_PERCENT,
];

export const COL2KEEP_CONTRATS_IMPORT = [
  TAB_CONTRATS_COL_IMPORT_ID,
  TAB_CONTRATS_COL_IMPORT_COLLAB,
  TAB_CONTRATS_COL_IMPORT_DATE_DEBUT,
  TAB_CONTRATS_COL_IMPORT_NB_SEMAINE_GARANTIE,
  TAB_CONTRATS_COL_IMPORT_RUPTURE,
  "IMPORT_DATE PAIEMENT CLIENT",
  TAB_CONTRATS_COL_IMPORT_CLIENT,
  TAB_CONTRATS_COL_IMPORT_TYPE,
  TAB_CONTRATS_COL_IMPORT_CANDIDAT,
  TAB_CONTRATS_COL_IMPORT_DESCRIPTION,
  TAB_CONTRATS_COL_IMPORT_SALAIRE,
  TAB_CONTRATS_COL_IMPORT_PERCENT,
];

export const TAB_VERSEMENT_COL_COLLAB = "NOM PRENOM COLLABORATEUR";

export const COL2KEEP_VERSEMENT = [
  TAB_VERSEMENT_COL_COLLAB,
  "DATE",
  "MONTANT VERSE",
  "AJOUT CONTRAT",
  "ALL CONTRATS",
];

export const TAB_VARIABLE_COL_COLLAB = "NOM PRENOM";

export const COL2KEEP_VARIABLE = [
  TAB_VARIABLE_COL_COLLAB,
  "DEBUT",
  "FIN",
  "FREQ VARIABLE",
  "T1 MINI",
  "T1 %",
  "T2 MINI",
  "T2 %",
  "T3 MINI",
  "T3 %",
  "T4 MINI",
  "T4 %",
];

export const TAB_COLLAB_COL_COLLAB = "NOM PRENOM";
export const TAB_COLLAB_COL_EMAIL = "EMAIL";
export const TAB_COLLAB_COL_SHEET_ID = "SHEET ID";

export const COL2KEEP_COLLAB = [
  TAB_COLLAB_COL_COLLAB,
  "CONTRAT",
  TAB_COLLAB_COL_EMAIL,
];

export const COL2KEEP_CLIENTS = ["NOM CLIENT", "NB WEEKS GARANTIE"];

export const TAB_PARAMETRES_COL_EMAIL = "NEW CONTRAT ALERT EMAIL LIST";
export const COL2KEEP_PARAMETRES = [TAB_PARAMETRES_COL_EMAIL];
