# Guide de déploiement pour les TI — Priorité Bonneville

Ce guide s'adresse à un administrateur Microsoft 365 qui ne connaît pas le projet. Les détails d'utilisation sont dans `README.md`.

## 1. Résumé

« Priorité Bonneville » est un complément Web pour Outlook. Pendant la rédaction, l'expéditeur choisit un niveau (Urgent, P1, P2, P3, Confidentiel) et, au besoin, un numéro de projet. Le complément écrit une balise dans l'objet, précédée d'une pastille de couleur, par exemple `🟠 [M_12345_P1] Objet` ou `🟣 [CONFIDENTIEL_M_12345] Objet`. Chez les destinataires de Bonneville, des règles de réception côté serveur appliquent une catégorie réservée : « Priorité · Urgent », « Priorité · P1 », « Priorité · P2 », « Priorité · P3 » ou « CONFIDENTIEL ». Les catégories personnelles des employés ne sont jamais modifiées.

Ce qu'il fait : il lit et modifie l'objet du courriel en cours de rédaction, affiche une bannière à la création et un rappel à l'envoi si aucune priorité n'est choisie (sans jamais bloquer l'envoi), et enregistre trois préférences par utilisateur dans sa boîte aux lettres.

Ce qu'il ne fait pas : il ne lit ni le corps, ni les pièces jointes, ni les destinataires ; il n'envoie aucune donnée à l'extérieur ; il ne chiffre rien. « Confidentiel » est un indicateur visuel, pas une protection (voir la section 8).

Clients pris en charge : Outlook classique pour Windows (Microsoft 365, Windows 10 version 1903 ou plus pour les événements) et nouvel Outlook pour Windows. Mobile : aucune commande en rédaction ; les utilisateurs tapent la balise à la main, et les règles de réception s'appliquent quand même.

## 2. Rôles requis

| Étape | Rôle ou autorisation |
|---|---|
| Déploiement du complément | Administrateur général, ou administrateur Exchange avec accès aux applications intégrées |
| Catégories (Creer-Categories.ps1) | Application Entra ID avec l'autorisation d'application Microsoft Graph `MailboxSettings.ReadWrite` et `User.Read.All`, consentement administrateur |
| Règles de réception (Creer-ReglesReception.ps1) | Administrateur Exchange (Exchange Online PowerShell) |
| Mise en forme conditionnelle (optionnelle) | Aucun : exécutée sur le poste de l'utilisateur, dans sa session |

## 3. Hébergement

Les fichiers du complément sont statiques (HTML, CSS, JavaScript, images). Ils sont actuellement publiés sur GitHub Pages :
`https://ericgloutnez.github.io/PrioriteOutlook`

Pour les héberger ailleurs (Azure Static Web Apps, stockage Azure avec site statique, serveur Web interne accessible en HTTPS depuis Internet) :

1. Copier à la racine du site les dossiers `src` et `assets`.
2. Modifier `baseUrl` dans `config.json` (sans barre oblique finale).
3. Régénérer le manifeste :

```powershell
.\scripts\Generer-Manifeste.ps1 -WhatIf
.\scripts\Generer-Manifeste.ps1
```

4. Vérifier que `<baseUrl>/src/taskpane.html` s'ouvre dans un navigateur.

Le certificat HTTPS doit être valide et public : Outlook refuse les certificats non approuvés.

## 4. Déploiement, dans l'ordre

1. **Hébergement** : confirmer l'adresse (section 3) et la disponibilité de `src/taskpane.html`.
2. **Manifeste** : utiliser `manifest.xml` généré pour cette adresse.
3. **Déploiement centralisé** : Centre d'administration Microsoft 365 > Paramètres > Applications intégrées > Charger des applications personnalisées > Complément Office > téléverser `manifest.xml`.
   - Affecter d'abord à un groupe pilote, puis à toute l'organisation.
   - Mode de déploiement : **Facultatif** (installé par défaut, l'utilisateur peut le retirer). Les fonctions par événement exigent un déploiement par l'administrateur ; vérifier qu'elles restent actives en mode Facultatif auprès du groupe pilote.
4. **Catégories** : exécuter `Creer-Categories.ps1` (section 5).
5. **Règles de réception** : exécuter `Creer-ReglesReception.ps1` (section 5), après les catégories.
6. **Propagation** : le complément peut prendre jusqu'à 24 heures à apparaître dans tous les clients. Redémarrer Outlook accélère souvent l'affichage.
7. **Nouveaux employés** : relancer les scripts des étapes 4 et 5 pour leurs boîtes (ils sont idempotents), ou les planifier.

## 5. Commandes PowerShell

Toujours lancer avec `-WhatIf` d'abord, lire le résultat, puis lancer sans `-WhatIf`. Chaque script écrit un journal dans `scripts\journaux\`.

```powershell
# Modules
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser

# Catégories : connexion par application (certificat), puis exécution
Connect-MgGraph -ClientId "<id-application>" -TenantId "<id-locataire>" -CertificateThumbprint "<empreinte>"
.\scripts\Creer-Categories.ps1 -Boites pilote1@bonneville.ca, pilote2@bonneville.ca -WhatIf
.\scripts\Creer-Categories.ps1 -Boites pilote1@bonneville.ca, pilote2@bonneville.ca
.\scripts\Creer-Categories.ps1 -ToutesLesBoites -WhatIf
.\scripts\Creer-Categories.ps1 -ToutesLesBoites

# Règles de réception
Connect-ExchangeOnline
.\scripts\Creer-ReglesReception.ps1 -Boites pilote1@bonneville.ca -WhatIf
.\scripts\Creer-ReglesReception.ps1 -Boites pilote1@bonneville.ca
.\scripts\Creer-ReglesReception.ps1 -ToutesLesBoites -WhatIf
.\scripts\Creer-ReglesReception.ps1 -ToutesLesBoites
```

Avertissement Exchange : créer ou modifier des règles de réception par PowerShell peut supprimer les règles « client seulement » que l'utilisateur a désactivées dans Outlook. Prévenir le groupe pilote et vérifier sur une boîte de test avant le déploiement général.

## 6. Vérifications après déploiement

- [ ] Le bouton « Priorité » apparaît dans le ruban d'un nouveau courriel (Outlook classique et nouvel Outlook).
- [ ] La bannière apparaît à la création d'un courriel, d'une réponse et d'un transfert.
- [ ] Le panneau applique `[P1]`, `[M_12345_P1]` et `[CONFIDENTIEL_M_12345]` au début de l'objet (après RE: ou TR:), et remplace une balise existante sans la dupliquer.
- [ ] Le rappel apparaît à l'envoi d'un courriel sans balise, et « envoyer quand même » fonctionne.
- [ ] Chez un destinataire interne, chaque niveau reçoit la bonne catégorie et la bonne couleur, y compris avec un numéro de projet.
- [ ] Un courriel sans balise, ou contenant « P1 » sans crochets, ne reçoit aucune catégorie.
- [ ] Une balise tapée à la main sur un téléphone est reconnue chez le destinataire.
- [ ] Un employé peut retirer le complément et le réinstaller.

## 7. Retour en arrière

```powershell
# Retirer les règles de réception
Connect-ExchangeOnline
.\scripts\Supprimer-ReglesReception.ps1 -ToutesLesBoites -WhatIf
.\scripts\Supprimer-ReglesReception.ps1 -ToutesLesBoites
```

Retirer le complément : Centre d'administration > Paramètres > Applications intégrées > Priorité Bonneville > Supprimer l'application.

Les catégories restent dans les boîtes (elles ne gênent pas). Les balises déjà écrites dans les objets restent du texte ordinaire.

## 8. Sécurité

**Contenu du dépôt** : fichiers statiques seulement. Aucun secret, jeton, mot de passe ni adresse interne. Aucun serveur applicatif, aucune base de données.

**Autorisation du manifeste** : `ReadWriteItem`, le niveau minimal qui permet de modifier l'objet du courriel en cours de rédaction. Le complément ne demande ni `ReadWriteMailbox` ni accès à d'autres éléments.

**Données lues et écrites** : l'objet du courriel en cours de rédaction, et trois préférences de l'utilisateur (`banniereActive`, `demanderProjet`, `typeDefaut`) dans les paramètres itinérants du complément, stockés dans sa boîte aux lettres. Aucune donnée n'est envoyée à un service externe.

**Ressources externes** : Office.js est chargé depuis `appsforoffice.microsoft.com` (Microsoft). Aucune autre bibliothèque tierce.

**« Confidentiel »** : simple indicateur visuel ; il ne restreint ni le transfert ni la lecture. Si Bonneville utilise les étiquettes de confidentialité Microsoft Purview, les employés devraient les utiliser pour les courriels réellement confidentiels.

## 9. Personne-ressource

Eric — Directeur LFS, Bonneville
Courriel : ______________________
Poste : ______________________
