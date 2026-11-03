/*
 * tag.js — Logique de la balise de priorité dans l'objet du courriel.
 *
 * Format :
 *   [P1]            niveau seul
 *   [P1-M-12345]    niveau, type (M = Multi, R = Résidentiel), numéro
 *
 * Utilisé par le panneau (taskpane.js). Le fichier launchevent.js contient
 * une copie de la même expression régulière, parce que l'Outlook classique
 * exécute les événements dans un environnement JavaScript seul, sans page
 * HTML pour charger d'autres fichiers. Garder les deux synchronisées.
 */
(function (global) {
  "use strict";

  var NIVEAUX = ["URGENT", "P1", "P2", "P3", "CONFIDENTIEL"];
  var TYPES = { M: "Multi", R: "Résidentiel" };

  // Une seule balise par objet ; insensible à la casse pour tolérer
  // une saisie manuelle comme [p1] sur le téléphone.
  var REGEX_BALISE = /\[(URGENT|P1|P2|P3|CONFIDENTIEL)(?:-(M|R)-(\d+))?\]/i;

  /**
   * Lit la balise présente dans l'objet.
   * @returns {{niveau:string, type:string|null, numero:string|null,
   *            texte:string, index:number} | null}
   */
  function lire(objet) {
    var m = REGEX_BALISE.exec(objet || "");
    if (!m) return null;
    return {
      niveau: m[1].toUpperCase(),
      type: m[2] ? m[2].toUpperCase() : null,
      numero: m[3] || null,
      texte: m[0],
      index: m.index
    };
  }

  /** Construit la balise, ex. [P1] ou [P1-M-12345]. */
  function construire(niveau, type, numero) {
    if (NIVEAUX.indexOf(niveau) === -1) {
      throw new Error("Niveau inconnu : " + niveau);
    }
    if (numero) {
      if (!/^\d+$/.test(numero)) throw new Error("Le numéro doit contenir des chiffres seulement.");
      if (!TYPES[type]) throw new Error("Choisir Multi ou Résidentiel.");
      return "[" + niveau + "-" + type + "-" + numero + "]";
    }
    return "[" + niveau + "]";
  }

  /**
   * Applique la balise : remplace une balise existante à sa position
   * (même après RE:, TR:, FW:), sinon l'ajoute au début de l'objet.
   */
  function appliquer(objet, balise) {
    objet = objet || "";
    var existante = lire(objet);
    if (existante) {
      return objet.slice(0, existante.index) + balise + objet.slice(existante.index + existante.texte.length);
    }
    return objet.length ? balise + " " + objet : balise;
  }

  /** Retire la balise et l'espace qui la suit, s'il y a lieu. */
  function retirer(objet) {
    objet = objet || "";
    var existante = lire(objet);
    if (!existante) return objet;
    var avant = objet.slice(0, existante.index);
    var apres = objet.slice(existante.index + existante.texte.length);
    if (apres.charAt(0) === " ") apres = apres.slice(1);
    return avant + apres;
  }

  global.PrioriteTag = {
    NIVEAUX: NIVEAUX,
    TYPES: TYPES,
    REGEX_BALISE: REGEX_BALISE,
    lire: lire,
    construire: construire,
    appliquer: appliquer,
    retirer: retirer
  };
})(typeof window !== "undefined" ? window : this);
