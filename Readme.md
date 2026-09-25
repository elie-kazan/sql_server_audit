# Invoke-SqlPerfAudit.ps1 : audit de performance SQL Server

Script PowerShell qui analyse une ou plusieurs instances SQL Server et produit, pour chacune, un **rapport Word (.docx)** listant les problèmes de performance qui méritent votre attention, classés par gravité, avec pour chacun une explication et une recommandation.

L'objectif est simple : savoir en quelques minutes **s'il y a quelque chose qui ne va pas sur une instance, et quoi faire en priorité**.

---

## Sommaire

1. [Points clés](#1-points-clés)
2. [Prérequis](#2-prérequis)
3. [Installation](#3-installation)
4. [Utilisation](#4-utilisation)
5. [Durée d'exécution et impact sur le serveur](#5-durée-dexécution-et-impact-sur-le-serveur)
6. [Contenu du rapport](#6-contenu-du-rapport)
7. [Niveaux de gravité](#7-niveaux-de-gravité)
8. [Détail des vérifications et des seuils](#8-détail-des-vérifications-et-des-seuils)
9. [Interpréter les résultats](#9-interpréter-les-résultats)
10. [Ce que le script ne vérifie pas](#10-ce-que-le-script-ne-vérifie-pas)
11. [Dépannage](#11-dépannage)
12. [Sécurité et confidentialité](#12-sécurité-et-confidentialité)

---

## 1. Points clés

- **Lecture seule.** Le script interroge uniquement des vues système (DMV et catalogues). Il ne modifie rien sur le serveur : aucune table créée, aucun paramètre changé, aucune procédure installée.
- **Aucune installation.** Un seul fichier `.ps1`. Il n'utilise que des composants .NET fournis avec Windows PowerShell 5.1 : pas de module à installer, pas de Python, pas de Microsoft Office.
- **Exécution à distance.** Le script n'a pas besoin de tourner sur le serveur SQL. N'importe quel poste Windows capable de se connecter à l'instance suffit (poste d'administration, serveur de rebond).
- **Plusieurs instances d'un coup.** Un rapport Word distinct est produit pour chaque instance.
- **Résilient.** Si une vérification échoue (version trop ancienne, droit manquant), le rapport l'indique dans la section concernée et l'audit continue.
- **Rapide et léger.** En général de 10 à 60 secondes par instance, avec une charge comparable à l'ouverture du Moniteur d'activité.

---

## 2. Prérequis

### Poste d'exécution

| Élément | Exigence |
|---|---|
| Système | Windows (poste ou serveur) |
| PowerShell | Windows PowerShell 5.1 (inclus dans Windows 10/11 et Windows Server 2016+). PowerShell 7 fonctionne également. |
| Réseau | Accès au port SQL Server de l'instance (1433 par défaut) |
| Word | Non requis pour générer le rapport, seulement pour l'ouvrir |

### Versions de SQL Server

Le script est conçu pour **SQL Server 2012 à 2025** sous Windows ou Linux. SQL Server 2008 / 2008 R2 devrait fonctionner pour l'essentiel. Deux informations nécessitent une version récente et sont simplement omises sinon :

| Information | Version minimale | Comportement si absente |
|---|---|---|
| Initialisation instantanée des fichiers | 2016 SP1 | Affichée « Inconnu », aucun constat |
| Nombre de VLF (fichiers journaux virtuels) | 2016 SP2 / 2017 | Vérification ignorée |

Azure SQL Database n'est pas pris en charge. Azure SQL Managed Instance n'a pas été testé.

### Droits SQL Server

Le compte utilisé n'a **pas besoin d'être sysadmin**. Il lui faut :

| Droit | Pourquoi |
|---|---|
| `VIEW SERVER STATE` | Lire les statistiques d'attente, CPU, mémoire, E/S, requêtes, index manquants, blocages |
| `VIEW ANY DEFINITION` | Lire la liste des fichiers de bases (`sys.master_files`) |
| `SELECT` sur `msdb.dbo.backupset` | Vérifier la présence de sauvegardes du journal |

Exemple de création d'un compte d'audit dédié (authentification Windows) :

```sql
USE master;
CREATE LOGIN [DOMAINE\svc_audit_sql] FROM WINDOWS;
GRANT VIEW SERVER STATE TO [DOMAINE\svc_audit_sql];
GRANT VIEW ANY DEFINITION TO [DOMAINE\svc_audit_sql];

USE msdb;
CREATE USER [DOMAINE\svc_audit_sql] FOR LOGIN [DOMAINE\svc_audit_sql];
GRANT SELECT ON dbo.backupset TO [DOMAINE\svc_audit_sql];
```

---

## 3. Installation

1. Copiez `Invoke-SqlPerfAudit.ps1` dans un dossier, par exemple `C:\Scripts`.
2. Si le fichier a été téléchargé ou reçu par e-mail, débloquez-le :

   ```powershell
   Unblock-File C:\Scripts\Invoke-SqlPerfAudit.ps1
   ```

3. Si la stratégie d'exécution PowerShell bloque les scripts, lancez-le ainsi, sans modifier la stratégie de la machine :

   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File C:\Scripts\Invoke-SqlPerfAudit.ps1 -SqlInstance SQLPROD01
   ```

---

## 4. Utilisation

### Exemples

```powershell
# Instance par défaut, authentification Windows, rapport dans le dossier courant
.\Invoke-SqlPerfAudit.ps1 -SqlInstance SQLPROD01

# Instance nommée et port spécifique
.\Invoke-SqlPerfAudit.ps1 -SqlInstance 'SQLPROD01\INST1'
.\Invoke-SqlPerfAudit.ps1 -SqlInstance 'SQLPROD01,1450'

# Plusieurs instances, rapports dans un dossier donné
.\Invoke-SqlPerfAudit.ps1 -SqlInstance 'SQLPROD01','SQLPROD02\INST1' -OutputFolder D:\Audits

# Authentification SQL (une fenêtre demande l'identifiant et le mot de passe)
.\Invoke-SqlPerfAudit.ps1 -SqlInstance SQLPROD01 -Credential (Get-Credential)

# Certificat du serveur non reconnu par le poste (certificat auto-signé)
.\Invoke-SqlPerfAudit.ps1 -SqlInstance SQLPROD01 -TrustServerCertificate

# Liste d'instances depuis un fichier texte (une instance par ligne)
.\Invoke-SqlPerfAudit.ps1 -SqlInstance (Get-Content .\serveurs.txt) -OutputFolder \\partage\audits
```

### Paramètres

| Paramètre | Obligatoire | Défaut | Description |
|---|---|---|---|
| `-SqlInstance` | Oui | | Une ou plusieurs instances : `SERVEUR`, `SERVEUR\INSTANCE` ou `SERVEUR,PORT`. |
| `-Credential` | Non | Authentification Windows | Identifiants d'un login SQL, via `Get-Credential`. |
| `-OutputFolder` | Non | Dossier courant | Dossier des rapports, créé s'il n'existe pas. |
| `-Top` | Non | 10 | Nombre de lignes des tableaux « top » (attentes, requêtes, fichiers, index manquants). Entre 5 et 50. |
| `-Encrypt` | Non | | Force le chiffrement de la connexion. |
| `-TrustServerCertificate` | Non | | Accepte le certificat du serveur sans le valider. |

### Fichiers produits

Un fichier par instance, nommé :

```
SQLPerfAudit_<instance>_<AAAAMMJJ_HHMM>.docx
```

Les caractères interdits dans un nom de fichier sont remplacés par `_`. Par exemple, `SQLPROD01\INST1` audité le 25/09/2026 à 10h12 donne `SQLPerfAudit_SQLPROD01_INST1_20260925_1012.docx`.

### Affichage console et synthèse multi-instances

Pendant l'exécution, la console affiche chaque vérification au fur et à mesure, puis le chemin du rapport et le nombre de constats par gravité.

Le script renvoie aussi un objet par instance (`Instance`, `Report`, `High`, `Medium`, `Low`, `Info`). Cela permet de produire une synthèse de tout un parc, exploitable dans Excel :

```powershell
.\Invoke-SqlPerfAudit.ps1 -SqlInstance (Get-Content .\serveurs.txt) -OutputFolder D:\Audits |
    Export-Csv D:\Audits\synthese.csv -NoTypeInformation -Delimiter ';' -Encoding UTF8
```

### Exécution planifiée

Le script peut être lancé par une tâche planifiée Windows, par exemple chaque lundi matin. En authentification Windows, c'est le compte de la tâche qui se connecte à SQL Server : il doit donc disposer des droits décrits plus haut.

```
Programme : powershell.exe
Arguments : -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Invoke-SqlPerfAudit.ps1" -SqlInstance SQLPROD01 -OutputFolder D:\Audits
```

---

## 5. Durée d'exécution et impact sur le serveur

### Durée

Comptez **10 à 60 secondes par instance** dans la plupart des cas. Le script lit des vues que SQL Server maintient déjà en mémoire : il n'y a ni échantillonnage dans le temps, ni lecture des tables utilisateur. La génération du fichier Word prend moins d'une seconde.

Trois vérifications peuvent être plus longues sur certains serveurs :

| Vérification | Cas où elle ralentit |
|---|---|
| Requêtes les plus coûteuses | Cache de plans très volumineux (centaines de milliers de plans, fréquent avec du SQL dynamique non paramétré) : de 10 secondes à quelques minutes. |
| Sauvegardes du journal | Historique de sauvegarde dans `msdb` jamais purgé (millions de lignes). |
| Fichiers journaux virtuels | Centaines de bases ou journaux très volumineux. |

Chaque requête a un délai maximal de 5 minutes. S'il est dépassé, la section concernée l'indique et l'audit continue. Les instances sont traitées l'une après l'autre.

### Impact

L'impact est négligeable et le script peut être lancé en journée. La session d'audit :

- lit les données sans poser de verrous (`READ UNCOMMITTED`) ;
- abandonne une requête bloquée plus de 10 secondes plutôt que d'attendre (`LOCK_TIMEOUT`) ;
- se désigne comme victime en cas de deadlock (`DEADLOCK_PRIORITY LOW`), pour ne jamais gêner une transaction applicative.

La connexion apparaît sous le nom d'application **« SQL Perf Audit (read-only) »**, ce qui permet de l'identifier dans les outils de supervision.

---

## 6. Contenu du rapport

Le rapport fait généralement de 6 à 10 pages au format A4 et contient les parties suivantes.

**1. En-tête.** Nom du serveur, version, édition, date de l'audit, date du dernier redémarrage et compte utilisé.

**2. Synthèse.** C'est la partie à lire en priorité :

- un **verdict en une phrase** (par exemple : « Il y a 2 problèmes de gravité haute qui méritent une attention rapide ») ;
- le **nombre de constats par gravité**, avec un code couleur ;
- le **tableau de tous les constats**, du plus important au moins important. Chaque ligne indique la gravité, le domaine, le problème constaté avec les chiffres réels, et l'action recommandée.

**3. Sections détaillées.** Une section numérotée par vérification. Chacune reprend les constats de son domaine avec une ligne « What to do », ou affiche « No problems found in this area », puis présente un tableau des données mesurées.

**4. À propos du rapport.** Rappel de la méthode, des limites et de la durée de l'audit.

Pied de page : nom du serveur et « Page X of Y ».

> **Langue du rapport :** le rapport Word est rédigé en **anglais**, ce qui facilite son partage avec des éditeurs, des prestataires ou le support Microsoft. Ce README est en français pour votre équipe.

---

## 7. Niveaux de gravité

| Gravité | Couleur | Signification | Délai d'action suggéré |
|---|---|---|---|
| **High** | Rouge | Pénalise probablement déjà les performances ou la stabilité. | Rapidement |
| **Medium** | Orange | Vrai problème ou risque à corriger. | Prochaine fenêtre de maintenance |
| **Low** | Bleu | Amélioration de bonne pratique. | Quand l'occasion se présente |
| **Info** | Gris | Contexte utile, pas d'action nécessaire. | |

---

## 8. Détail des vérifications et des seuils

Les seuils sont issus des recommandations Microsoft et des pratiques courantes de la communauté SQL Server. Ce sont des repères, pas des règles absolues (voir [section 9](#9-interpréter-les-résultats)).

### 8.1 Instance

Présente la version, l'édition, les processeurs et nœuds NUMA utilisés, la mémoire physique, la virtualisation et la date de démarrage.

| Constat | Gravité | Pourquoi c'est important |
|---|---|---|
| Version hors support étendu Microsoft (2016 et antérieures) | Medium | Plus aucun correctif de sécurité ou de performance. Le support de SQL Server 2016 a pris fin en juillet 2026. |
| SQL Server 2017 | Info | Fin du support étendu en octobre 2027 : à anticiper. |
| Processeurs inutilisables par SQL Server | High | Des processeurs payés ne servent à rien, souvent à cause d'une limite de licence de l'édition ou de la topologie sockets/cœurs de la VM. |
| Initialisation instantanée des fichiers désactivée | Medium | Chaque croissance de fichier de données et chaque restauration doivent remettre l'espace à zéro, ce qui bloque l'activité. |
| Redémarrage il y a moins de 7 jours | Info | Les statistiques cumulées ne couvrent pas une semaine type. |

### 8.2 Configuration du serveur

| Constat | Gravité | Pourquoi c'est important |
|---|---|---|
| `max server memory` non configuré (illimité) | High | SQL Server peut priver Windows de mémoire, ce qui provoque de la pagination et de l'instabilité. Le rapport propose une valeur de départ adaptée à la RAM du serveur. |
| `max server memory` trop proche de la RAM totale | Medium | Marge insuffisante pour le système d'exploitation. |
| `MAXDOP` illimité avec plus de 8 processeurs | Medium | Une seule requête peut monopoliser tous les processeurs. Le rapport indique la valeur recommandée par Microsoft pour la topologie CPU/NUMA du serveur. |
| `MAXDOP` supérieur à la recommandation | Low | Parallélisme excessif possible. |
| `MAXDOP` = 1 | Info | Parallélisme désactivé. Parfois exigé par l'éditeur (SharePoint par exemple), mais pénalisant pour les requêtes lourdes. |
| `cost threshold for parallelism` à 5 (défaut) | Medium | Même des requêtes légères partent en parallèle et gaspillent du CPU. Valeur de départ conseillée : 50. |
| `optimize for ad hoc workloads` désactivé | Low | Les plans à usage unique encombrent le cache. |
| `priority boost` activé | High | Paramètre obsolète qui peut affamer le système et le cluster. |
| `lightweight pooling` (mode fibre) activé | Medium | Rarement bénéfique et incompatible avec plusieurs fonctionnalités. |
| `max worker threads` modifié | Low | Masque souvent un problème de blocage ou de parallélisme au lieu de le résoudre. |
| Modifications de configuration en attente | Low | Valeurs configurées mais pas encore actives : `RECONFIGURE` ou redémarrage nécessaire. |

### 8.3 Statistiques d'attente (wait statistics)

C'est souvent la vérification la plus révélatrice : elle montre **sur quoi SQL Server passe son temps à attendre** depuis le dernier redémarrage. Les attentes internes sans signification (tâches de fond en sommeil) sont filtrées. Le tableau présente les 10 principales attentes avec leur signification en langage clair.

| Constat | Gravité |
|---|---|
| Une des 5 premières attentes représente 10 % ou plus du total | Low (10 %), Medium (25 %), High (50 %) |
| Attente de mémoire pour requêtes (`RESOURCE_SEMAPHORE`) parmi les principales | High |
| Attente réseau côté client (`ASYNC_NETWORK_IO`) ou de sauvegarde | Toujours Low : cause généralement extérieure au serveur |
| Pénurie de threads de travail (`THREADPOOL`) au-delà de 10 secondes cumulées | High : les utilisateurs ont probablement subi des délais d'attente ou un serveur figé |
| Attentes de signal à 20 % ou plus du temps d'attente | Medium : signe de pression CPU |

Signification des attentes les plus fréquentes :

| Attente | Indique généralement |
|---|---|
| `PAGEIOLATCH_*` | Lecture de données depuis le disque : stockage lent ou mémoire insuffisante |
| `CXPACKET`, `CXSYNC_*` | Parallélisme |
| `SOS_SCHEDULER_YIELD` | Pression CPU, souvent due à des parcours de tables en mémoire |
| `WRITELOG` | Écriture lente du journal de transactions |
| `LCK_M_*` | Blocages entre sessions |
| `PAGELATCH_*` | Contention en mémoire, souvent dans tempdb |
| `RESOURCE_SEMAPHORE` | Requêtes en attente de mémoire pour trier ou joindre |
| `ASYNC_NETWORK_IO` | L'application lit les résultats trop lentement |

### 8.4 Processeur

Analyse environ les **4 dernières heures**, à raison d'une mesure par minute issue du ring buffer de SQL Server.

| Constat | Gravité |
|---|---|
| CPU moyen de SQL Server à 80 % ou plus | High |
| CPU moyen à 60 % ou plus, ou supérieur à 80 % pendant 20 % du temps ou plus | Medium |
| Autres processus du serveur utilisant 20 % du CPU ou plus (antivirus, autres services, autre instance) | Medium |
| Compilations à 15 % ou plus des requêtes (moyenne depuis le démarrage) : les plans ne sont pas réutilisés | Medium |

### 8.5 Mémoire

| Constat | Gravité |
|---|---|
| Serveur à court de mémoire (état signalé par Windows ou moins de 512 Mo disponibles) | High |
| Page life expectancy inférieure au seuil adapté à la taille mémoire (300 s par tranche de 4 Go) | Medium |
| Requêtes en attente d'allocation mémoire au moment de l'audit | Medium |

La page life expectancy est une mesure instantanée, prise au moment de l'audit. Mieux vaut la vérifier à plusieurs moments de la journée avant de conclure.

### 8.6 Latence du stockage

Mesure le temps moyen de lecture et d'écriture de chaque fichier de base depuis le dernier redémarrage. Seuls les fichiers ayant au moins 1 000 lectures ou écritures sont évalués, pour éviter les faux positifs.

| Constat | Seuils |
|---|---|
| Lectures lentes sur les fichiers de données | 20 ms ou plus : Medium ; 50 ms ou plus : High |
| Écritures lentes sur le journal de transactions | 5 ms ou plus : Low ; 10 ms ou plus : Medium ; 20 ms ou plus : High |

Le journal est jugé plus sévèrement car **chaque validation de transaction (COMMIT) attend son écriture**.

### 8.7 TempDB

| Constat | Gravité |
|---|---|
| Moins de fichiers de données que recommandé (un par processeur, jusqu'à 8) | Medium |
| Fichiers de données de tailles différentes | Low |
| Croissance automatique en pourcentage | Low |

### 8.8 Bases de données

Seules les bases présentant au moins un problème figurent dans le tableau ; le nombre de bases saines est indiqué en dessous.

| Constat | Gravité | Pourquoi c'est important |
|---|---|---|
| `AUTO_SHRINK` activé | High | Cycles de réduction et de croissance qui fragmentent les index et consomment CPU et E/S. |
| `AUTO_CLOSE` activé | Medium | La base est fermée et rouverte en permanence et perd son cache. |
| Statistiques automatiques désactivées | Medium | Des statistiques obsolètes produisent de mauvais plans d'exécution. |
| `PAGE_VERIFY` différent de `CHECKSUM` | Medium | Une corruption peut passer inaperçue. |
| Croissance automatique de 1 Mo ou moins | Medium | Des milliers de petites croissances, chacune bloquant l'activité. |
| Croissance automatique en pourcentage | Low | Croissances de plus en plus longues à mesure que le fichier grossit. |
| Mode de récupération FULL sans sauvegarde du journal depuis 24 h | Medium | Le journal grossit sans limite. À ignorer si les sauvegardes sont faites sur un autre réplica d'un groupe de disponibilité. |
| Plus de 300 fichiers journaux virtuels (VLF) | Low ; Medium au-delà de 1 000 | Ralentit la récupération, les restaurations et les sauvegardes du journal. |
| Niveau de compatibilité en retard de 3 versions ou plus | Low | La base ne profite pas des améliorations de l'optimiseur. |

### 8.9 Requêtes les plus coûteuses

Liste les instructions ayant consommé le plus de CPU d'après le cache de plans. Les requêtes identiques qui ne diffèrent que par leurs valeurs sont regroupées. Pour chacune : base, nombre d'exécutions, CPU total et moyen, lectures moyennes, durée moyenne, part du CPU total et début du texte SQL.

| Constat | Gravité |
|---|---|
| Une seule instruction consomme 25 % ou plus du CPU total du cache | Medium : l'optimiser aura un effet important |
| Les 5 premières instructions consomment 60 % ou plus du CPU | Info : quelques optimisations ciblées suffiront à faire une différence visible |

### 8.10 Index manquants

Liste les index que l'optimiseur de requêtes aurait voulu utiliser, classés par bénéfice estimé. Un constat de gravité **Medium** est émis pour les suggestions à fort bénéfice (score de 100 000 ou plus).

> **Attention :** ces suggestions se recoupent souvent et sont parfois trop larges. Examinez-les, regroupez-les et testez-les ; ne les créez jamais telles quelles.

### 8.11 Blocages et transactions

| Constat | Gravité |
|---|---|
| Requêtes bloquées au moment de l'audit | Medium |
| Transactions ouvertes depuis plus de 10 minutes | Medium. Le rapport signale les sessions inactives (« sleeping ») avec une transaction ouverte, ce qui indique en général une application qui n'a pas validé sa transaction. |
| Deadlocks : 10 par jour ou plus en moyenne depuis le démarrage | Medium |
| Deadlocks : au moins 1 par jour | Low |

---

## 9. Interpréter les résultats

**Un constat signifie « à examiner », pas forcément « à corriger ».** Le contexte décide. Par exemple :

- `MAXDOP = 1` peut être une exigence de l'éditeur de l'application ;
- une forte attente de parallélisme la nuit peut venir d'un traitement de reporting prévu ;
- un index manquant peut être couvert par un index existant légèrement différent.

**Méthode recommandée :**

1. Lisez le verdict et le tableau de synthèse de la première page.
2. Traitez d'abord les constats **High**, puis les **Medium**.
3. Pour chaque constat, consultez la section détaillée et son tableau de données.
4. Testez toute modification sur un environnement hors production.
5. Relancez l'audit après les corrections, idéalement après une semaine d'activité normale, et comparez les rapports.

**Deux types de chiffres cohabitent dans le rapport :**

| Type | Exemples | Portée |
|---|---|---|
| Cumulés depuis le dernier redémarrage | Attentes, latences d'E/S, requêtes coûteuses, deadlocks, compilations | Représentatifs si le serveur tourne depuis au moins une semaine |
| Instantanés au moment de l'audit | Page life expectancy, blocages, transactions ouvertes, allocations mémoire en attente | À confirmer à plusieurs moments de la journée |

Un redémarrage récent, ou une journée inhabituelle (migration, gros traitement ponctuel), fausse l'analyse. Le rapport signale un redémarrage de moins de 7 jours.

**Le cache de plans n'est pas exhaustif.** Les requêtes recompilées ou évincées du cache n'apparaissent pas dans les requêtes les plus coûteuses.

---

## 10. Ce que le script ne vérifie pas

Pour rester simple et rapide, le script n'analyse pas :

- la fragmentation des index ;
- les index inutilisés ou en double ;
- la fraîcheur des statistiques table par table ;
- la couverture des sauvegardes complètes et des contrôles d'intégrité (`DBCC CHECKDB`) ;
- les travaux de l'Agent SQL Server ;
- la sécurité (logins, permissions, surface d'exposition) ;
- l'état des groupes de disponibilité (Always On) ;
- les tendances dans le temps : chaque audit est une photographie. Pour suivre l'évolution, conservez les rapports successifs.

---

## 11. Dépannage

| Symptôme | Cause probable | Solution |
|---|---|---|
| « L'exécution de scripts est désactivée sur ce système » | Stratégie d'exécution PowerShell | Lancer avec `powershell.exe -ExecutionPolicy Bypass -File ...` (voir [section 3](#3-installation)). |
| « Le fichier n'est pas signé numériquement » | Fichier téléchargé marqué comme provenant d'Internet | `Unblock-File .\Invoke-SqlPerfAudit.ps1` |
| « A network-related or instance-specific error » | Nom d'instance incorrect, pare-feu, service SQL Browser arrêté (instances nommées) | Vérifier le nom ; essayer `SERVEUR,PORT` ; tester avec `Test-NetConnection SERVEUR -Port 1433`. |
| « The certificate chain was issued by an authority that is not trusted » | Certificat auto-signé sur le serveur | Ajouter `-TrustServerCertificate`. |
| « Login failed for user » | Compte inconnu ou mot de passe erroné | Vérifier le login ; utiliser `-Credential` pour un login SQL. |
| « VIEW SERVER STATE permission was denied » | Droits insuffisants | Accorder les droits de la [section 2](#droits-sql-server). |
| Une section indique « This check could not run » | Droit manquant, version ancienne ou délai dépassé | Le message d'erreur est affiché dans la section. Le reste du rapport reste valable. |
| « Backup history not checked » | Pas d'accès à `msdb.dbo.backupset` | Accorder `SELECT` sur `msdb.dbo.backupset`. |
| « Performance counters unavailable » | Compteurs de performance désactivés sur l'instance | Les constats mémoire et deadlocks sont alors partiels ; le reste fonctionne. |
| Le fichier .docx ne s'ouvre pas | Fichier en cours d'écriture ou audit interrompu | Relancer l'audit ; vérifier l'espace disque et les droits d'écriture sur le dossier de sortie. |

---

## 12. Sécurité et confidentialité

- **Mots de passe.** Avec `-Credential`, le mot de passe n'est ni affiché, ni journalisé, ni écrit dans le rapport. L'authentification Windows reste préférable.
- **Contenu du rapport.** Le rapport contient les noms des bases, des tables et des fichiers, ainsi que des **extraits du texte des requêtes SQL**, qui peuvent inclure des valeurs métier (noms de clients, montants) lorsque l'application n'utilise pas de requêtes paramétrées. **Traitez le rapport comme un document confidentiel** et partagez-le avec discernement, en particulier à l'extérieur de l'entreprise.
- **Traçabilité.** La connexion est identifiable dans SQL Server sous le nom d'application « SQL Perf Audit (read-only) ». Le compte utilisé est indiqué dans l'en-tête du rapport.

---

*Invoke-SqlPerfAudit.ps1, version 1.0. Pour le premier déploiement, testez le script sur une instance non critique avant de l'utiliser sur l'ensemble du parc.*
