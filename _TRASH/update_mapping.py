
import json
import os

content = {
  "ENTETE": {
    "cells": {
      "A2": { "attribute": "denomination_sociale", "label": "DÉNOMINATION SOCIALE", "bold": True, "value_template": "DÉNOMINATION SOCIALE : {}" }
    }
  },
  "Fiche R1": {
    "title": "FICHE R1",
    "cells": {
      "A5": {
        "attribute": "denomination_sociale",
        "value_template": "Description sociale de l'entreprise : {}",
        "label": "Description"
      },
      "A6": {
        "attribute": "adresse_complete",
        "value_template": "Siège social : {}",
        "label": "Adresse"
      },
      "A7": {
        "attribute": "num_identification_fiscale",
        "value_template": "N° d'identification fiscal : {}",
        "label": "NIU"
      },
      "E9": {
        "attribute": "periode_exercice_str",
        "label": "Periode Exercice"
      },
      "B11": {
        "attribute": "date_arrete_comptes",
        "format": "date",
        "label": "Date Arrete"
      },
      "B13": {
        "attribute": "exercice_precedent_fin",
        "format": "date",
        "label": "Cloture Prec."
      },
      "A16": {
        "attribute": "registre_commerce",
        "label": "RCCM",
        "optional": True
      },
      "C18": { "attribute": "cnps", "label": "No Caisse", "optional": True },
      "F18": { "attribute": "code_importateur", "label": "Code Imp.", "optional": True },
      "I18": { "attribute": "activite_principale_code", "label": "Code Act.", "optional": True }
    }
  },
  "Fiche R2": {
    "title": "FICHE R2",
    "cells": {
      "B9": { "attribute": "forme_juridique", "label": "Forme Juridique" },
      "B11": { "attribute": "registre_fiscal", "label": "Regisme Fiscal" },
      "B13": { "attribute": "pays_siege", "label": "Pays", "value_template": "CAMEROUN" },
      "B20": { "attribute": "premiere_annee", "label": "1ere annee", "optional": True }
    },
    "tables": {
      "activites": {
        "data_start_row": "A29",
        "fields": {
          "B": "designation",
          "F": "code_nomenclature",
          "M": "chiffre_affaire_ht",
          "O": "pourcentage_ca"
        }
      }
    }
  },
  "Fiche R3": {
    "title": "FICHE R3",
    "tables": {
      "dirigeants": {
        "data_start_row": "A11",
        "fields": {
          "A": "nom",
          "B": "prenoms",
          "C": "qualite",
          "D": "num_identification_fiscale",
          "E": "adresse"
        }
      },
      "conseil_administration": {
        "data_start_row": "A25", 
        "fields": {
          "B": "nom",
          "C": "prenoms",
          "D": "qualite",
          "E": "adresse"
        }
      },
      "actionnaires": {
        "data_start_row": "A40",
        "fields": {
           "A": "nom",
           "C": "nationalite",
           "D": "nombre",
           "E": "montant_total"
        }
      }
    }
  },
  "NOTE 13 ": {
    "title": "NOTE 13",
    "tables": {
         "actionnaires": {
             "data_start_row": "A8",
             "fields": {
                 "A": "nom",
                 "B": "nationalite",
                 "D": "nombre",
                 "E": "montant_total"
             }
         }
    }
  }
}

with open('dsf_prefill_mapping.json', 'w', encoding='utf-8') as f:
    json.dump(content, f, indent=2, ensure_ascii=False)
    print("dsf_prefill_mapping.json updated successfully.")
