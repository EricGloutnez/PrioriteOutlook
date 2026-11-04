/*
 * launchevent.js — Activation par événement du complément « Priorité Bonneville ».
 *
 *  - OnNewMessageCompose : affiche une bannière en haut du courriel
 *    (« Aucune priorité choisie… » ou « Priorité : [M_12345_P1] »)
 *    avec une action qui ouvre le panneau latéral.
 *  - OnMessageSend (Smart Alerts, mode PromptUser) : si l'objet n'a
 *    aucune balise, propose de choisir une priorité ou d'envoyer
 *    quand même. L'envoi n'est jamais bloqué.
 *
 * Règle de sécurité : chaque gestionnaire appelle event.completed()
 * dans TOUS les cas, y compris en cas d'erreur (try/catch) et après un
 * délai maximal. Sans cela, Outlook affiche « … prend plus de temps que
 * prévu » au moment de l'envoi.
 *
 * Les deux événements respectent le paramètre « Bannière et rappel à
 * l'envoi » de l'utilisateur (roamingSettings, clé banniereActive).
 *
 * IMPORTANT : dans l'Outlook classique, ce fichier est exécuté seul
 * (sans page HTML). Il ne doit dépendre d'aucun autre fichier : la
 * détection de balise est donc copiée ici (voir tag.js).
 */

/* global Office */

var PRIORITE_ID_NOTIFICATION = "prioriteBonneville";
var PRIORITE_ID_BOUTON = "msgComposeOpenPaneButton";
var PRIORITE_REGEX = /\[(?:(CONFIDENTIEL)(?:_([MR])_(\d+))?|(?:([MR])_(\d+)_)?(URGENT|P1|P2|P3))\]/i;
var PRIORITE_DELAI_MAX_MS = 4000;

/*
 * Journal de diagnostic : les 30 dernières entrées sont gardées dans le
 * stockage local du navigateur (partagé avec le panneau, même origine),
 * et affichées dans les paramètres du panneau. Sans effet si le stockage
 * n'existe pas (Outlook classique, environnement JavaScript seul).
 */
function prioriteJournal(message) {
  try {
    var cle = "prioriteJournal";
    var entrees = JSON.parse(localStorage.getItem(cle) || "[]");
    entrees.push(new Date().toTimeString().slice(0, 8) + "  " + message);
    localStorage.setItem(cle, JSON.stringify(entrees.slice(-30)));
  } catch (e) { /* stockage indisponible */ }
}

prioriteJournal("launchevent.js chargé (Office " + (typeof Office) +
  ", actions " + (typeof Office !== "undefined" && Office.actions ? "présent" : "absent") + ")");

if (typeof window !== "undefined" && window.addEventListener) {
  window.addEventListener("error", function (e) {
    prioriteJournal("ERREUR JS : " + (e && e.message ? e.message : "inconnue"));
  });
}

/**
 * Enveloppe event.completed pour qu'il ne soit appelé qu'une seule fois,
 * et automatiquement après PRIORITE_DELAI_MAX_MS si rien ne l'a fait.
 */
function prioriteTerminaison(event, optionsParDefaut) {
  var fait = false;
  function terminer(options) {
    if (fait) return;
    fait = true;
    clearTimeout(minuterie);
    prioriteJournal("event.completed(" + (options ? JSON.stringify({ allowEvent: options.allowEvent }) : "") + ")");
    try {
      if (options) { event.completed(options); } else { event.completed(); }
    } catch (e) {
      prioriteJournal("ERREUR completed : " + e.message);
      // Dernier recours : terminer sans options.
      try { event.completed(); } catch (e2) { /* rien d'autre à faire */ }
    }
  }
  var minuterie = setTimeout(function () {
    prioriteJournal("délai maximal atteint");
    terminer(optionsParDefaut);
  }, PRIORITE_DELAI_MAX_MS);
  return terminer;
}

/** Lit un paramètre de l'utilisateur, avec une valeur par défaut. */
function prioriteParametre(nom, defaut) {
  try {
    var valeur = Office.context.roamingSettings.get(nom);
    return valeur === undefined || valeur === null ? defaut : valeur;
  } catch (e) {
    return defaut;
  }
}

function prioriteBanniereActive() {
  return prioriteParametre("banniereActive", true) !== false;
}

/** Affiche la bannière ; se rabat sur un message simple si l'action n'est pas prise en charge. */
function prioriteAfficherBanniere(texteBalise, terminer) {
  var item = Office.context.mailbox.item;
  var message = texteBalise
    ? "Priorité : " + texteBalise
    : "Aucune priorité choisie pour ce courriel";
  var action = texteBalise ? "Modifier" : "Choisir la priorité";

  var avecAction = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
    message: message,
    icon: "Icon.16",
    actions: [{
      actionText: action,
      actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
      commandId: PRIORITE_ID_BOUTON,
      contextData: "{}"
    }]
  };

  try {
    item.notificationMessages.replaceAsync(PRIORITE_ID_NOTIFICATION, avecAction, function (r) {
      prioriteJournal("bannière : " + r.status + (r.error ? " (" + r.error.message + ")" : ""));
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        terminer();
        return;
      }
      // Repli : message d'information sans bouton (le ruban reste disponible).
      try {
        item.notificationMessages.replaceAsync(PRIORITE_ID_NOTIFICATION, {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: message + " (bouton Priorité du ruban)",
          icon: "Icon.16",
          persistent: true
        }, function () { terminer(); });
      } catch (e) {
        terminer();
      }
    });
  } catch (e) {
    terminer();
  }
}

/** Événement : création d'un nouveau message, d'une réponse ou d'un transfert. */
function onNewMessageComposeHandler(event) {
  prioriteJournal("OnNewMessageCompose reçu");
  var terminer = prioriteTerminaison(event, null);
  try {
    if (!prioriteBanniereActive()) {
      terminer();
      return;
    }
    Office.context.mailbox.item.subject.getAsync(function (r) {
      try {
        var objet = r.status === Office.AsyncResultStatus.Succeeded ? r.value : "";
        var m = PRIORITE_REGEX.exec(objet || "");
        prioriteAfficherBanniere(m ? m[0].toUpperCase() : null, terminer);
      } catch (e) {
        terminer();
      }
    });
  } catch (e) {
    terminer();
  }
}

/** Événement : envoi du message (Smart Alerts). */
function onMessageSendHandler(event) {
  prioriteJournal("OnMessageSend reçu");
  // En cas de doute (erreur, délai dépassé), l'envoi est toujours permis.
  var terminer = prioriteTerminaison(event, { allowEvent: true });
  try {
    if (!prioriteBanniereActive()) {
      terminer({ allowEvent: true });
      return;
    }
    Office.context.mailbox.item.subject.getAsync(function (r) {
      try {
        if (r.status !== Office.AsyncResultStatus.Succeeded || PRIORITE_REGEX.test(r.value || "")) {
          terminer({ allowEvent: true });
          return;
        }
        var options = {
          allowEvent: false,
          errorMessage: "Aucune priorité choisie pour ce courriel. Vous pouvez en choisir une ou envoyer quand même."
        };
        // Bouton « Choisir la priorité » qui ouvre le panneau (Mailbox 1.14).
        // Le mode PromptUser est déjà fixé dans le manifeste : pas de sendModeOverride.
        if (Office.context.requirements.isSetSupported("Mailbox", "1.14")) {
          options.cancelLabel = "Choisir la priorité";
          options.commandId = PRIORITE_ID_BOUTON;
        }
        terminer(options);
      } catch (e) {
        terminer({ allowEvent: true });
      }
    });
  } catch (e) {
    terminer({ allowEvent: true });
  }
}

// Association des gestionnaires aux noms déclarés dans le manifeste.
// Dans Outlook sur le Web et le nouvel Outlook, Office.actions peut ne
// pas encore exister au chargement de ce fichier : on réessaie alors
// dans Office.onReady. Sans association, Outlook attend une réponse qui
// ne vient jamais (« … prend plus de temps que prévu »).
var prioriteAssocie = false;
function prioriteAssocier() {
  if (prioriteAssocie) return;
  if (typeof Office === "undefined" || !Office.actions || !Office.actions.associate) return;
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  prioriteAssocie = true;
  prioriteJournal("gestionnaires associés");
}
prioriteAssocier();

// Poignée de main avec Outlook : la page doit signaler qu'elle est prête
// avec Office.onReady. Sans cet appel, Outlook sur le Web et le nouvel
// Outlook chargent la page mais n'appellent jamais les gestionnaires
// (constaté dans le journal : « gestionnaires associés », puis rien).
// Dans l'Outlook classique (JavaScript seul), Office.onReady peut ne pas
// exister : l'association immédiate suffit.
if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady(function (info) {
    prioriteJournal("Office.onReady appelé (" + (info && info.host ? info.host : "hôte inconnu") + ")");
    prioriteAssocier();
  });
}
