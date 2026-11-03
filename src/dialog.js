/*
 * dialog.js — Fenêtre « Numéro de projet » du complément Priorité Bonneville.
 *
 * Ouverte par le panneau (taskpane.js) avec displayDialogAsync.
 * Paramètres reçus dans l'adresse : niveau, type (M ou R), numero.
 * Messages renvoyés au panneau (Office.context.ui.messageParent) :
 *   {"action":"appliquer","type":"M","numero":"11111"}
 *   {"action":"passer"}      niveau seul, sans projet
 * Fermer la fenêtre par le X annule tout (géré par le panneau).
 */

/* global Office */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var params = new URLSearchParams(window.location.search);
  var niveau = (params.get("niveau") || "").toUpperCase();
  var type = (params.get("type") || "").toUpperCase() || null;
  var COULEURS = { URGENT: "urgent", P1: "p1", P2: "p2", P3: "p3", CONFIDENTIEL: "confidentiel" };
  var LIBELLES = { URGENT: "Urgent", P1: "P1", P2: "P2", P3: "P3", CONFIDENTIEL: "Confidentiel" };

  function envoyer(message) {
    Office.context.ui.messageParent(JSON.stringify(message));
  }

  function balise(numero) {
    if (!numero) return "[" + niveau + "]";
    if (!type) return "";
    return niveau === "CONFIDENTIEL"
      ? "[CONFIDENTIEL_" + type + "_" + numero + "]"
      : "[" + type + "_" + numero + "_" + niveau + "]";
  }

  function majTypes() {
    document.querySelectorAll(".type").forEach(function (b) {
      b.setAttribute("aria-checked", String(b.dataset.type === type));
    });
    majApercu();
  }

  function majApercu() {
    var b = balise($("numeroProjet").value.trim());
    $("apercuDialogue").textContent = b ? "Balise : " + b : "";
  }

  function appliquer() {
    var numero = $("numeroProjet").value.trim();
    if (!numero) {
      envoyer({ action: "passer" });
      return;
    }
    if (!/^\d+$/.test(numero)) {
      $("erreurProjet").textContent = "Le numéro de projet contient des chiffres seulement.";
      return;
    }
    if (!type) {
      $("erreurProjet").textContent = "Choisissez Multi ou Résidentiel.";
      return;
    }
    envoyer({ action: "appliquer", type: type, numero: numero });
  }

  function initialiser() {
    var pastille = $("pastilleNiveau");
    pastille.textContent = LIBELLES[niveau] || niveau;
    pastille.classList.add("niveau-" + (COULEURS[niveau] || "p1"));
    $("numeroProjet").value = (params.get("numero") || "").replace(/\D/g, "");

    document.querySelectorAll(".type").forEach(function (b) {
      b.addEventListener("click", function () {
        type = type === b.dataset.type ? null : b.dataset.type;
        $("erreurProjet").textContent = "";
        majTypes();
      });
    });

    $("numeroProjet").addEventListener("input", function (e) {
      var propre = e.target.value.replace(/\D/g, "");
      if (propre !== e.target.value) e.target.value = propre;
      $("erreurProjet").textContent = "";
      majApercu();
    });

    $("btnAppliquer").addEventListener("click", appliquer);
    $("btnPasser").addEventListener("click", function () { envoyer({ action: "passer" }); });

    document.addEventListener("keydown", function (e) {
      if (e.key === "Enter") { e.preventDefault(); appliquer(); }
      if (e.key === "Escape") { e.preventDefault(); envoyer({ action: "passer" }); }
    });

    majTypes();
    $("numeroProjet").focus();
  }

  Office.onReady(initialiser);
})();
