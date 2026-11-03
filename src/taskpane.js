/*
 * taskpane.js — Panneau latéral du complément « Priorité Bonneville ».
 *
 * Paramètres de l'utilisateur (roamingSettings, suivent l'utilisateur
 * d'un poste à l'autre) :
 *   banniereActive  (booléen, défaut true)   bannière + rappel à l'envoi
 *   demanderProjet  (booléen, défaut false)  afficher la section Projet
 *   typeDefaut      ("", "M" ou "R", défaut "") type présélectionné
 */

/* global Office, PrioriteTag */
(function () {
  "use strict";

  var ID_NOTIFICATION = "prioriteBonneville";
  var ID_BOUTON = "msgComposeOpenPaneButton";
  var DEFAUTS = { banniereActive: true, demanderProjet: false, typeDefaut: "" };

  var etat = { niveau: null, type: null, projetVisible: false };
  var $ = function (id) { return document.getElementById(id); };

  // ---------- Paramètres ----------

  function lireParametre(nom) {
    var v = Office.context.roamingSettings.get(nom);
    return v === undefined || v === null ? DEFAUTS[nom] : v;
  }

  function enregistrerParametre(nom, valeur) {
    Office.context.roamingSettings.set(nom, valeur);
    Office.context.roamingSettings.saveAsync(function (r) {
      $("statutParametres").textContent = r.status === Office.AsyncResultStatus.Succeeded
        ? "Paramètres enregistrés"
        : "Les paramètres n'ont pas pu être enregistrés. Réessayez.";
    });
  }

  // ---------- Affichage ----------

  function majNiveaux() {
    document.querySelectorAll(".niveau").forEach(function (b) {
      b.setAttribute("aria-checked", String(b.dataset.niveau === etat.niveau));
    });
  }

  function majTypes() {
    document.querySelectorAll(".type").forEach(function (b) {
      b.setAttribute("aria-checked", String(b.dataset.type === etat.type));
    });
  }

  function majProjet() {
    $("sectionProjet").hidden = !etat.projetVisible;
    $("lienAjouterProjet").hidden = etat.projetVisible;
  }

  function majApercu(objet) {
    var balise = PrioriteTag.lire(objet);
    var p = $("apercu");
    p.textContent = "";
    if (balise) {
      p.appendChild(document.createTextNode("Balise actuelle : "));
      var code = document.createElement("code");
      code.textContent = balise.texte.toUpperCase();
      p.appendChild(code);
    } else {
      p.textContent = "Aucune priorité choisie";
    }
  }

  function effacerMessages() {
    $("erreurNiveau").textContent = "";
    $("erreurProjet").textContent = "";
    $("statut").textContent = "";
  }

  // ---------- Bannière ----------

  function majBanniere(balise) {
    if (lireParametre("banniereActive") === false) return;
    var item = Office.context.mailbox.item;
    var notif = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
      message: balise ? "Priorité : " + balise : "Aucune priorité choisie pour ce courriel",
      icon: "Icon.16",
      actions: [{
        actionText: balise ? "Modifier" : "Choisir la priorité",
        actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
        commandId: ID_BOUTON,
        contextData: "{}"
      }]
    };
    item.notificationMessages.replaceAsync(ID_NOTIFICATION, notif, function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        item.notificationMessages.replaceAsync(ID_NOTIFICATION, {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: notif.message,
          icon: "Icon.16",
          persistent: true
        });
      }
    });
  }

  // ---------- Objet du courriel ----------

  function lireObjet(suite) {
    Office.context.mailbox.item.subject.getAsync(function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        $("statut").textContent = "L'objet n'a pas pu être lu. Réessayez.";
        return;
      }
      suite(r.value || "");
    });
  }

  function ecrireObjet(objet, suite) {
    Office.context.mailbox.item.subject.setAsync(objet, function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        $("statut").textContent = "L'objet n'a pas pu être modifié. Réessayez.";
        return;
      }
      suite();
    });
  }

  // ---------- Actions ----------

  function appliquer() {
    effacerMessages();
    if (!etat.niveau) {
      $("erreurNiveau").textContent = "Choisissez un niveau de priorité.";
      return;
    }
    var numero = etat.projetVisible ? $("numeroProjet").value.trim() : "";
    if (numero && !/^\d+$/.test(numero)) {
      $("erreurProjet").textContent = "Le numéro de projet contient des chiffres seulement.";
      return;
    }
    if (numero && !etat.type) {
      $("erreurProjet").textContent = "Choisissez Multi ou Résidentiel.";
      return;
    }

    var balise = PrioriteTag.construire(etat.niveau, numero ? etat.type : null, numero || null);
    lireObjet(function (objet) {
      var nouvel = PrioriteTag.appliquer(objet, balise);
      ecrireObjet(nouvel, function () {
        majApercu(nouvel);
        majBanniere(balise);
        $("statut").textContent = "Balise appliquée";
      });
    });
  }

  function retirer() {
    effacerMessages();
    lireObjet(function (objet) {
      if (!PrioriteTag.lire(objet)) {
        $("statut").textContent = "Ce courriel n'a aucune balise de priorité.";
        return;
      }
      var nouvel = PrioriteTag.retirer(objet);
      ecrireObjet(nouvel, function () {
        etat.niveau = null;
        $("numeroProjet").value = "";
        etat.type = lireParametre("typeDefaut") || null;
        etat.projetVisible = lireParametre("demanderProjet") === true;
        majNiveaux(); majTypes(); majProjet(); majApercu(nouvel);
        majBanniere(null);
        $("statut").textContent = "Balise retirée";
      });
    });
  }

  // ---------- Vues ----------

  function ouvrirParametres() {
    $("paramBanniere").checked = lireParametre("banniereActive") !== false;
    $("paramProjet").checked = lireParametre("demanderProjet") === true;
    var typeDefaut = lireParametre("typeDefaut") || "";
    document.querySelectorAll('input[name="typeDefaut"]').forEach(function (r) {
      r.checked = r.value === typeDefaut;
    });
    $("statutParametres").textContent = "";
    $("vuePrincipale").hidden = true;
    $("vueParametres").hidden = false;
  }

  function fermerParametres() {
    $("vueParametres").hidden = true;
    $("vuePrincipale").hidden = false;
    // Appliquer les nouveaux réglages à la vue, sans masquer un projet déjà saisi.
    if (!$("numeroProjet").value) {
      etat.projetVisible = lireParametre("demanderProjet") === true;
      etat.type = lireParametre("typeDefaut") || etat.type;
      majTypes(); majProjet();
    }
  }

  // ---------- Initialisation ----------

  function brancher() {
    document.querySelectorAll(".niveau").forEach(function (b) {
      b.addEventListener("click", function () {
        etat.niveau = b.dataset.niveau;
        $("erreurNiveau").textContent = "";
        $("statut").textContent = "";
        majNiveaux();
      });
    });

    document.querySelectorAll(".type").forEach(function (b) {
      b.addEventListener("click", function () {
        etat.type = etat.type === b.dataset.type ? null : b.dataset.type;
        $("erreurProjet").textContent = "";
        majTypes();
      });
    });

    $("numeroProjet").addEventListener("input", function (e) {
      var propre = e.target.value.replace(/\D/g, "");
      if (propre !== e.target.value) e.target.value = propre;
      $("erreurProjet").textContent = "";
    });
    $("numeroProjet").addEventListener("keydown", function (e) {
      if (e.key === "Enter") appliquer();
    });

    $("lienAjouterProjet").addEventListener("click", function () {
      etat.projetVisible = true;
      if (!etat.type) etat.type = lireParametre("typeDefaut") || null;
      majTypes(); majProjet();
      $("numeroProjet").focus();
    });

    $("btnAppliquer").addEventListener("click", appliquer);
    $("btnRetirer").addEventListener("click", retirer);
    $("btnParametres").addEventListener("click", ouvrirParametres);
    $("btnRetour").addEventListener("click", fermerParametres);

    $("paramBanniere").addEventListener("change", function (e) {
      enregistrerParametre("banniereActive", e.target.checked);
      if (!e.target.checked) {
        Office.context.mailbox.item.notificationMessages.removeAsync(ID_NOTIFICATION);
      }
    });
    $("paramProjet").addEventListener("change", function (e) {
      enregistrerParametre("demanderProjet", e.target.checked);
    });
    document.querySelectorAll('input[name="typeDefaut"]').forEach(function (r) {
      r.addEventListener("change", function () {
        if (r.checked) enregistrerParametre("typeDefaut", r.value);
      });
    });
  }

  function initialiser() {
    brancher();
    etat.type = lireParametre("typeDefaut") || null;
    etat.projetVisible = lireParametre("demanderProjet") === true;

    lireObjet(function (objet) {
      var balise = PrioriteTag.lire(objet);
      if (balise) {
        etat.niveau = balise.niveau;
        if (balise.numero) {
          // Un projet existant est toujours affiché, peu importe le paramètre.
          etat.projetVisible = true;
          etat.type = balise.type;
          $("numeroProjet").value = balise.numero;
        }
      }
      majNiveaux(); majTypes(); majProjet(); majApercu(objet);
    });
  }

  Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) initialiser();
  });
})();
