/*
 * taskpane.js — Panneau latéral du complément « Priorité Bonneville ».
 *
 * Fonctionnement :
 *  - Un clic sur un niveau applique la balise.
 *  - Si « Demander un numéro de projet » est activé (ou si la balise
 *    actuelle contient déjà un projet), la fenêtre « Numéro de projet »
 *    (dialog.html) s'ouvre d'abord : Appliquer = niveau + projet,
 *    Passer = niveau seul, fermeture par le X = rien n'est modifié.
 *  - Le lien « Ajouter un projet » ouvre la même fenêtre pour le niveau
 *    choisi.
 *
 * Paramètres de l'utilisateur (roamingSettings) :
 *   banniereActive  (booléen, défaut true)   bannière + rappel à l'envoi
 *   demanderProjet  (booléen, défaut false)  ouvrir la fenêtre du projet
 *   typeDefaut      ("", "M" ou "R", défaut "") type présélectionné
 */

/* global Office, PrioriteTag */
(function () {
  "use strict";

  var ID_NOTIFICATION = "prioriteBonneville";
  var ID_BOUTON = "msgComposeOpenPaneButton";
  var DEFAUTS = { banniereActive: true, demanderProjet: false, typeDefaut: "" };
  var VERSION_FICHIERS = "4";

  var etat = { niveau: null, type: null, numero: null };
  var dialogue = null;
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

  /** Écrit la balise dans l'objet et met l'affichage à jour. */
  function appliquerBalise(niveau, type, numero) {
    var balise;
    try {
      balise = PrioriteTag.construire(niveau, numero ? type : null, numero || null);
    } catch (e) {
      $("statut").textContent = e.message;
      return;
    }
    lireObjet(function (objet) {
      var nouvel = PrioriteTag.appliquer(objet, balise);
      ecrireObjet(nouvel, function () {
        etat.niveau = niveau;
        etat.type = numero ? type : etat.type;
        etat.numero = numero || null;
        majNiveaux();
        majApercu(nouvel);
        majBanniere(balise);
        $("statut").textContent = "Balise appliquée : " + balise;
      });
    });
  }

  // ---------- Fenêtre « Numéro de projet » ----------

  function ouvrirDialogue(niveau) {
    if (dialogue) { try { dialogue.close(); } catch (e) { /* déjà fermée */ } dialogue = null; }

    var type = etat.type || lireParametre("typeDefaut") || "";
    var adresse = new URL("dialog.html", window.location.href);
    adresse.search = new URLSearchParams({
      v: VERSION_FICHIERS,
      niveau: niveau,
      type: type,
      numero: etat.numero || ""
    }).toString();

    Office.context.ui.displayDialogAsync(adresse.toString(),
      { height: 42, width: 28, displayInIframe: true },
      function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          // Repli : appliquer le niveau seul et le signaler.
          $("statut").textContent = "La fenêtre du projet n'a pas pu s'ouvrir ; niveau appliqué sans projet.";
          appliquerBalise(niveau, null, null);
          return;
        }
        dialogue = r.value;

        dialogue.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          var msg = {};
          try { msg = JSON.parse(arg.message); } catch (e) { /* message invalide */ }
          dialogue.close();
          dialogue = null;
          if (msg.action === "appliquer") {
            appliquerBalise(niveau, msg.type, msg.numero);
          } else if (msg.action === "passer") {
            appliquerBalise(niveau, null, null);
          }
        });

        // Fermeture par le X (ou erreur) : rien n'est modifié.
        dialogue.addEventHandler(Office.EventType.DialogEventReceived, function () {
          dialogue = null;
          majNiveaux();
        });
      });
  }

  function choisirNiveau(niveau) {
    effacerMessages();
    var demander = lireParametre("demanderProjet") === true || !!etat.numero;
    if (demander) {
      ouvrirDialogue(niveau);
    } else {
      appliquerBalise(niveau, null, null);
    }
  }

  function ajouterProjet() {
    effacerMessages();
    if (!etat.niveau) {
      $("erreurNiveau").textContent = "Choisissez d'abord un niveau de priorité.";
      return;
    }
    ouvrirDialogue(etat.niveau);
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
        etat.numero = null;
        etat.type = null;
        majNiveaux();
        majApercu(nouvel);
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
  }

  // ---------- Initialisation ----------

  function brancher() {
    document.querySelectorAll(".niveau").forEach(function (b) {
      b.addEventListener("click", function () { choisirNiveau(b.dataset.niveau); });
    });

    $("lienAjouterProjet").addEventListener("click", ajouterProjet);
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
    lireObjet(function (objet) {
      var balise = PrioriteTag.lire(objet);
      if (balise) {
        etat.niveau = balise.niveau;
        etat.type = balise.type;
        etat.numero = balise.numero;
      }
      majNiveaux();
      majApercu(objet);
    });
  }

  Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) initialiser();
  });
})();
