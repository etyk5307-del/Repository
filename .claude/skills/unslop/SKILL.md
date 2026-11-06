---
name: unslop
description: Réécrit un texte en français pour le débarrasser du style générique « généré par IA » (formules creuses, emphase, tournures répétitives, mise en forme surchargée), sans en changer le fond ni les faits. À utiliser dès que l'utilisateur demande de « désloper », « rendre plus humain », « enlever le style ChatGPT/IA », « rendre plus naturel », « alléger », « rendre plus professionnel », de relire ou corriger le style d'un rapport, d'une note, d'un courriel, d'un post LinkedIn ou d'un rapport de stage, et aussi quand Claude rédige lui-même un livrable destiné à être lu par des professionnels (notes de renseignement, notes de sûreté, rapports, synthèses), même si le mot « unslop » n'est pas prononcé.
---

# Unslop — écrire comme un professionnel, pas comme un générateur

Un texte « slop » se reconnaît en trois lignes : il promet beaucoup, dit peu et sonne pareil quel que soit le sujet. Le lecteur professionnel (un chef de service, un jury, un décideur sécurité) le repère vite et décroche, ou pire, doute du fond. Le but de ce skill est de rendre au texte sa densité et sa voix, **sans toucher aux faits, aux chiffres, aux nuances d'incertitude ni à l'intention de l'auteur**.

## Méthode

1. **Lire tout le texte d'abord.** Repérer le message principal, le destinataire et le registre (note de renseignement, rapport académique, courriel, publication). Le registre décide du niveau de formalité, pas le goût du réécrivain.
2. **Repérer les tics** de la liste ci-dessous. Ne pas tout corriger mécaniquement : une formule peut être justifiée à un endroit précis.
3. **Réécrire** en privilégiant : phrases courtes à moyennes, verbes précis, sujet concret, une idée par phrase, le concret avant l'abstrait.
4. **Vérifier la fidélité** : chaque fait, chiffre, date, nom propre, source et degré de certitude de l'original doit se retrouver dans la version réécrite. Si une phrase vide ne contient aucune information, la supprimer est permis. Si une phrase floue semble cacher une information, la signaler plutôt que l'inventer.
5. **Livrer** selon la demande (voir « Formats de sortie »).

## Les tics à traquer

### Formules creuses et emphase
- Ouvertures de chatbot : « Bien sûr ! », « Excellente question », « Voici… », « Je serais ravi de… ».
- Clôtures de chatbot : « N'hésitez pas à… », « J'espère que cela vous aide », « En espérant avoir répondu… ».
- Remplissage : « Il est important de noter que », « Il convient de souligner », « Force est de constater », « Dans un monde en constante évolution », « À l'ère du numérique », « Plus que jamais ».
- Vocabulaire gonflé : *crucial, essentiel, incontournable, véritable, clé, majeur, fondamental, révolutionnaire, levier, catalyseur, pierre angulaire, synergie, écosystème, paradigme, holistique, robuste*. Garder le mot seulement s'il est exact.
- Verbes « IA » : *plonger dans, naviguer, explorer, débloquer, exploiter tout le potentiel, mettre en lumière, souligner l'importance, jouer un rôle clé, s'inscrire dans une dynamique*.
- Intensifs empilés : *véritablement, particulièrement, extrêmement, profondément, résolument*.

### Constructions répétitives
- « Non seulement… mais aussi… » et « Ce n'est pas X, c'est Y » en série.
- Les triplets systématiques (« rapide, efficace et durable »). Une énumération de trois n'est légitime que s'il y a vraiment trois éléments.
- Connecteurs enchaînés en tête de chaque phrase : *Par ailleurs, En outre, De plus, Ainsi, En effet, Dès lors*. En garder un quand la logique l'exige.
- Phrase-slogan ou morale en fin de paragraphe (« La sécurité est l'affaire de tous. »).
- Questions rhétoriques en guise de titres (« Pourquoi est-ce important ? »).
- Le deux-points dramatique (« Le constat est clair : … ») et les tirets cadratins en cascade.
- Nominalisations lourdes : « procéder à la mise en œuvre de la réalisation » → « réaliser ».
- Précautions empilées : « il semblerait que cela pourrait potentiellement » → une seule marque d'incertitude, la bonne.

### Mise en forme surchargée
- Gras sur la moitié des mots, émojis, titres pour chaque paragraphe, listes à puces pour des idées qui forment un raisonnement.
- Une liste convient aux éléments parallèles et indépendants ; un raisonnement (cause → effet → conséquence) s'écrit en phrases.
- Majuscules inutiles (« la Sécurité », « l'Équipe Projet ») et guillemets d'emphase.

### Anglicismes et calques
- *impacter* → toucher, affecter ; *adresser un problème* → traiter ; *faire du sens* → avoir du sens ; *challenger* → remettre en question ; *drive / driver* → conduire, piloter ; *insight* → enseignement ; *process* → processus, procédure.

## Règles propres aux écrits de renseignement et de sûreté

Ces textes ont des conventions qui ne sont pas du « slop » et qu'il faut **préserver** :
- **L'essentiel d'abord** (BLUF) : la conclusion ou l'alerte en tête, puis les éléments qui la fondent.
- **Langage estimatif calibré** : *presque certain, très probable, probable, possible, peu probable*. Ne jamais remplacer « probable » par « certain » ni supprimer un niveau de confiance. Ne pas ajouter d'incertitude là où l'auteur affirme un fait sourcé.
- **Distinguer le fait de l'évaluation** : « Le 18/09, un poste FAMa a été attaqué à Kassela » (fait) / « Cette attaque indique très probablement… » (évaluation).
- **Précision** : dates au format court, lieux avec région ou cercle, bilans chiffrés avec leur source (« selon AzZalaqa », « bilan revendiqué »). Remplacer « récemment » ou « plusieurs » par la date ou le nombre quand l'original les donne ailleurs.
- **Ton neutre** : pas de sensationnalisme (« carnage », « déferlante ») sauf citation, pas de jugement moral. Les acteurs sont désignés par leur nom ou sigle.
- **Recommandations actionnables** : un verbe à l'infinitif, un destinataire, un horizon (« Renforcer l'escorte des convois sur l'axe Ségou–Markala, court terme »).

## Règles propres aux écrits académiques (rapport de stage, mémoire)

- Registre soutenu mais sobre ; la première personne est acceptable pour le vécu du stage (« J'ai participé à… »), le « nous » de modestie pour l'analyse, sans mélanger les deux dans un même paragraphe.
- Chaque affirmation générale s'appuie sur une source, un chiffre ou une observation du terrain (par exemple : l'entreprise, l'équipe, le projet observé).
- Une citation est attribuée prudemment quand sa source est incertaine (« pensée attribuée à… »).
- L'introduction annonce la problématique et le plan en quelques phrases, sans formules du type « De nos jours, la formation est devenue un enjeu majeur ».

## Formats de sortie

Adapter à la demande ; par défaut, choisir **Réécriture** :
- **Réécriture** (par défaut) : le texte réécrit, prêt à l'emploi, puis, sous le texte, une courte liste « Principales modifications » (3 à 6 puces) et, s'il y en a, les **points à vérifier** (information ambiguë, chiffre incohérent, attribution douteuse).
- **Audit** (« dis-moi ce qui ne va pas », « relis ») : tableau *Passage original → Problème → Proposition*, limité aux 10 passages les plus utiles, puis une appréciation d'ensemble en deux phrases.
- **Version courte** (« plus court », « résume ») : viser 40 à 60 % de la longueur d'origine en gardant tous les faits et chiffres essentiels.

Garder la langue de l'original. Respecter la longueur approximative sauf demande contraire : désloper n'est pas tout résumer.

## Exemples

**Exemple 1 — note de sûreté**
Original : « Il est important de souligner que la situation sécuritaire dans la région de Ségou connaît une dégradation véritablement préoccupante. En effet, les groupes armés jouent un rôle clé dans la déstabilisation de cet écosystème fragile, ce qui impacte non seulement les populations mais aussi les activités économiques. »
Réécrit : « La situation se dégrade dans la région de Ségou : 10 incidents en deux mois, dont des embuscades répétées sur l'axe Niono–Dioura. Le JNIM en est l'auteur principal. Les convois commerciaux et les populations riveraines des axes sont les plus touchés. »
(Les chiffres viennent du reste du document ; sans eux, écrire « La situation se dégrade dans la région de Ségou » et signaler le manque de chiffres dans les points à vérifier.)

**Exemple 2 — rapport de stage**
Original : « Dans un monde en constante évolution, la formation continue est devenue un levier incontournable de performance. Elle permet non seulement de développer les compétences, mais aussi de renforcer la cohésion des équipes. »
Réécrit : « Les méthodes et les outils de gestion de projet changent vite ; une équipe qui ne se forme pas accumule un retard que l'expérience seule ne comble pas. Chez GObeLab, la formation continue vise deux effets : actualiser les compétences et renforcer la cohésion des équipes projet. »

**Exemple 3 — courriel**
Original : « Bonjour, j'espère que vous allez bien ! Je me permets de revenir vers vous concernant notre échange afin de faire le point sur l'avancement. N'hésitez pas à me faire part de vos retours. »
Réécrit : « Bonjour, pouvez-vous me dire où en est la cartographie des sites du Mali ? J'ai besoin de l'export avant vendredi pour l'intégrer au point de situation. Merci. »
(Ajouter une précision comme « avant vendredi » seulement si l'auteur l'a fournie ; sinon la laisser entre crochets à compléter.)

## Ce qu'il ne faut pas faire
- Inventer un fait, un chiffre, une date ou une source pour rendre le texte plus concret : laisser un [à préciser] à la place.
- Transformer le style personnel de l'auteur en un autre style uniforme : garder ses tournures quand elles sont claires.
- Remplacer un tic par un autre (supprimer « crucial » pour mettre « déterminant » partout).
- Ajouter des commentaires sur le fait que le texte « semblait écrit par une IA » : l'utilisateur veut un texte meilleur, pas un diagnostic sur son auteur.
