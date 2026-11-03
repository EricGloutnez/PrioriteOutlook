/*
 * launchevent.js — Activation par événement du complément « Priorité Bonneville ».
 *
 *  - OnNewMessageCompose : affiche une bannière en haut du courriel
 *    (« Aucune priorité choisie… » ou « Priorité : [P1-M-12345] »)
 *    avec une action qui ouvre le panneau latéral.
 *  - OnMessageSend (Smart Alerts, mode PromptUser) : si l'objet n'a
 *    aucune balise, propose de choisir une priorité ou d'envoyer
 *    quand même. L'envoi n'est jamais bloqué.
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
var PRIORITE_REGEX = /\[(URGENT|P1|P2|P3|CONFIDENTIEL)(?:-(M|R)-(\d+))?\]/i;

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

  item.notificationMessages.replaceAsync(PRIORITE_ID_NOTIFICATION, avecAction, function (r) {
    if (r.status === Office.AsyncResultStatus.Succeeded) {
      terminer();
      return;
    }
    // Repli : message d'information sans bouton (le ruban reste disponible).
    var simple = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message + " (bouton Priorité du ruban)",
      icon: "Icon.16",
      persistent: true
    };
    item.notificationMessages.replaceAsync(PRIORITE_ID_NOTIFICATION, simple, function () {
      terminer();
    });
  });
}

/** Événement : création d'un nouveau message, d'une réponse ou d'un transfert. */
function onNewMessageComposeHandler(event) {
  if (!prioriteBanniereActive()) {
    event.completed();
    return;
  }
  Office.context.mailbox.item.subject.getAsync(function (r) {
    var objet = r.status === Office.AsyncResultStatus.Succeeded ? r.value : "";
    var m = PRIORITE_REGEX.exec(objet || "");
    prioriteAfficherBanniere(m ? m[0].toUpperCase() : null, function () {
      event.completed();
    });
  });
}

/** Événement : envoi du message (Smart Alerts). */
function onMessageSendHandler(event) {
  if (!prioriteBanniereActive()) {
    event.completed({ allowEvent: true });
    return;
  }
  Office.context.mailbox.item.subject.getAsync(function (r) {
    // En cas d'erreur de lecture, ne jamais retenir l'envoi.
    if (r.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({ allowEvent: true });
      return;
    }
    if (PRIORITE_REGEX.test(r.value || "")) {
      event.completed({ allowEvent: true });
      return;
    }

    var options = {
      allowEvent: false,
      errorMessage: "Aucune priorité choisie pour ce courriel. Vous pouvez en choisir une ou envoyer quand même."
    };
    // Bouton « Choisir la priorité » qui ouvre le panneau (Mailbox 1.14).
    if (Office.context.requirements.isSetSupported("Mailbox", "1.14")) {
      options.cancelLabel = "Choisir la priorité";
      options.commandId = PRIORITE_ID_BOUTON;
      options.sendModeOverride = Office.MailboxEnums.SendModeOverride.PromptUser;
    }
    event.completed(options);
  });
}

// Association des gestionnaires aux noms déclarés dans le manifeste.
if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
