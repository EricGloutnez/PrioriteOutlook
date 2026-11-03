# Priorité Bonneville — complément Outlook

Complément Web pour Outlook qui ajoute un niveau de priorité, et au besoin un numéro de projet, dans l'objet des courriels. Chez les destinataires de Bonneville, une règle de réception côté serveur applique une catégorie de couleur selon la balise.

| Niveau | Balise sans projet | Balise avec projet | Couleur |
|---|---|---|---|
| Urgent | `[URGENT] Objet` | `[M_12345_URGENT] Objet` | Rouge |
| P1 | `[P1] Objet` | `[R_12345_P1] Objet` | Orange |
| P2 | `[P2] Objet` | `[M_12345_P2] Objet` | Jaune |
| P3 | `[P3] Objet` | `[R_12345_P3] Objet` | Vert |
| Confidentiel | `[CONFIDENTIEL] Objet` | `[CONFIDENTIEL_M_12345] Objet` | Mauve |

M = Multi, R = Résidentiel. Le numéro de projet contient des chiffres seulement. La balise est placée au début de l'objet (après RE: ou TR: dans une réponse ou un transfert), suivie d'un espace ; l'employé écrit ensuite le reste de l'objet. Pour un courriel confidentiel, CONFIDENTIEL passe en avant de tout.

Le complément fonctionne dans l'Outlook classique et le nouvel Outlook pour Windows, en rédaction. Sur Android et iPhone, la balise se tape à la main dans l'objet ; les règles de réception la reconnaissent de la même façon.

## Fonctionnement

À la création d'un courriel (nouveau, réponse, transfert), une bannière apparaît en haut du message avec le bouton « Choisir la priorité ». Il ouvre le panneau latéral, aussi accessible par le bouton « Priorité » du ruban (sous « Applications » dans Outlook sur le Web et le nouvel Outlook). Un clic sur un niveau écrit la balise dans l'objet ; si une balise existe déjà, elle est remplacée à sa place. « Retirer » l'enlève.

Quand « Demander un numéro de projet » est activé, une fenêtre « Numéro de projet » s'ouvre après le choix du niveau : Multi ou Résidentiel (le type par défaut est présélectionné) et le numéro. « Appliquer » écrit la balise avec le projet ; « Passer » ou Échap écrit le niveau seul ; fermer la fenêtre par le X ne modifie rien. Le lien « Ajouter un projet » ouvre la même fenêtre au besoin.

Si un courriel part sans balise, Outlook affiche un rappel au moment de l'envoi : on peut choisir une priorité ou envoyer quand même. L'envoi n'est jamais bloqué.

## Contenu du dépôt

```
manifest.template.xml   Gabarit du manifeste (jeton {{BASE_URL}})
manifest.xml            Manifeste généré pour GitHub Pages
config.json             Adresse d'hébergement (production et développement)
src/                    Panneau, fenêtre du projet, gestionnaires d'événements, logique de balise
assets/                 Icônes
scripts/                Scripts PowerShell (manifeste, catégories, règles, mise en forme)
tests/tag.test.js       Tests de la logique de balise (node tests/tag.test.js)
GUIDE-TI.md             Procédure de déploiement pour les TI
```

Aucune étape de compilation : les fichiers de `src/` sont publiés tels quels. Office.js est chargé depuis le réseau de Microsoft.

## Essai personnel

Les étapes ci-dessous ne demandent aucun droit d'administrateur.

### 1. Publier sur GitHub Pages

1. Créer le dépôt public `PrioriteOutlook` sous le compte `EricGloutnez` et y pousser le contenu de ce dossier.
2. Dans le dépôt : Settings > Pages > Source « Deploy from a branch », branche `main`, dossier `/ (root)`.
3. Attendre la publication, puis vérifier que cette page s'ouvre dans un navigateur :
   `https://ericgloutnez.github.io/PrioriteOutlook/src/taskpane.html`

Avec GitHub CLI :

```powershell
gh repo create PrioriteOutlook --public --source . --push
gh api -X POST repos/EricGloutnez/PrioriteOutlook/pages -f "source[branch]=main" -f "source[path]=/"
```

### 2. Générer le manifeste

`manifest.xml` est déjà généré pour l'adresse ci-dessus. Pour une autre adresse, modifier `config.json`, puis :

```powershell
.\scripts\Generer-Manifeste.ps1
```

### 3. Charger le complément dans Outlook (chargement manuel)

1. Ouvrir `https://aka.ms/olksideload` dans un navigateur, connecté au compte Bonneville. La fenêtre des compléments d'Outlook sur le Web s'ouvre.
2. Mes compléments > Compléments personnalisés > Ajouter un complément personnalisé > À partir d'un fichier.
3. Choisir `manifest.xml`, puis confirmer.
4. Le complément apparaît aussi dans l'Outlook classique et le nouvel Outlook du même compte après quelques minutes (redémarrer Outlook au besoin).

### 4. Simuler le résultat chez le destinataire

Les couleurs viennent des catégories et des règles de réception. Pour l'essai, deux possibilités.

Dans Outlook : créer les catégories URGENT, P1, P2, P3 et CONFIDENTIEL avec leur couleur, puis une règle par niveau, « si l'objet contient `[P1]` ou `_P1]`, classer dans la catégorie P1 » (pour Confidentiel : `[CONFIDENTIEL]` ou `[CONFIDENTIEL_`).

Par PowerShell, sur votre propre boîte, si votre compte le permet :

```powershell
.\scripts\Creer-Categories.ps1 -Boites votre.adresse@bonneville.ca -WhatIf
.\scripts\Creer-Categories.ps1 -Boites votre.adresse@bonneville.ca
.\scripts\Creer-ReglesReception.ps1 -Boites votre.adresse@bonneville.ca -WhatIf
.\scripts\Creer-ReglesReception.ps1 -Boites votre.adresse@bonneville.ca
```

Ensuite, s'envoyer des courriels avec chaque niveau et vérifier la couleur.

## Paramètres de l'utilisateur

L'engrenage du panneau ouvre les paramètres. Ils sont enregistrés dans la boîte aux lettres et suivent l'utilisateur d'un poste à l'autre.

| Paramètre | Défaut | Effet |
|---|---|---|
| Bannière et rappel à l'envoi | Activé | Désactivé : ni bannière ni rappel ; le bouton du ruban reste disponible. |
| Demander un numéro de projet | Désactivé | Activé : la fenêtre « Numéro de projet » s'ouvre après chaque choix de niveau. Désactivé : le niveau s'applique en un clic ; le lien « Ajouter un projet » ouvre la fenêtre au besoin. |
| Type de projet par défaut | Aucun | Multi ou Résidentiel est présélectionné dans la fenêtre, et reste modifiable. |

Si la balise du courriel contient déjà un projet (par exemple dans une réponse), la fenêtre s'ouvre toujours, préremplie avec le type et le numéro existants.

## Retirer le complément (employé)

Une fois déployé par les TI en mode « Facultatif », chaque employé peut le retirer de son Outlook : dans Outlook, ouvrir Applications (ou Obtenir des compléments) > Mes compléments, puis retirer « Priorité Bonneville ». Le même écran permet de le réinstaller. Le retrait vaut pour le compte, donc pour tous les postes de l'employé.

Un employé qui retire le complément continue de recevoir les couleurs, qui viennent des règles du serveur.

## Mise en forme conditionnelle (optionnelle, Outlook classique)

Met en gras et en couleur les lignes des courriels balisés dans la boîte de réception. Outlook classique seulement ; à exécuter sur chaque poste, par l'employé ou par les TI.

```powershell
.\scripts\MiseEnFormeConditionnelle.ps1 -Activer -WhatIf
.\scripts\MiseEnFormeConditionnelle.ps1 -Activer
.\scripts\MiseEnFormeConditionnelle.ps1 -Desactiver
```

Outlook n'offre pas l'orange ni un jaune lisible dans cette fonction : P1 est en bordeaux et P2 en olive.

## À valider à l'essai

Ces points n'ont pas pu être testés dans Outlook pendant le développement :

1. **Bannière avec bouton** (message de type InsightMessage avec action) dans l'Outlook classique et le nouvel Outlook. Si le bouton n'apparaît pas, le code affiche un message simple et le ruban reste disponible.
2. **Événements en chargement manuel** : la bannière et le rappel à l'envoi peuvent ne s'activer qu'une fois le complément déployé par un administrateur.
3. **Rappel à l'envoi** : bouton « Choisir la priorité » qui ouvre le panneau (exige Mailbox 1.14 ; sinon, rappel avec les boutons standards d'Outlook).
4. **Lecture des paramètres dans les événements** (roamingSettings) dans l'Outlook classique.
5. **Règles de réception** : reconnaissance de `[P1]` et `_P1]` par SubjectContainsWords, avec les crochets et le trait de soulignement. Vérifier aussi qu'un objet comme « P1-Fall Cottage » (sans crochets) ne reçoit aucune catégorie.
6. **Mise en forme conditionnelle** : traitement du crochet dans le filtre `LIKE '%[P1%'`.
7. **Noms des cmdlets Graph** des catégories (module Microsoft.Graph.Users).

## Développement local (facultatif)

Pour tester des modifications sans publier sur GitHub, servir le dossier en HTTPS sur le port 3000 avec un certificat de développement Office, puis générer `manifest.dev.xml` :

```powershell
npx office-addin-dev-certs install
npx http-server -S -C "$env:USERPROFILE\.office-addin-dev-certs\localhost.crt" -K "$env:USERPROFILE\.office-addin-dev-certs\localhost.key" -p 3000
.\scripts\Generer-Manifeste.ps1 -Dev
```

Charger ensuite `manifest.dev.xml` comme à l'étape 3. Avec le retrait d'EWS pour Exchange Online, le débogage F5 de Visual Studio ne fonctionne plus pour ce type de complément : utiliser le chargement manuel.

## Tests

```powershell
node tests/tag.test.js
```
