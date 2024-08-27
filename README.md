# VBBrainBox
Un système expert d'ordre 0+
---

VBBrainBox est issu de Turbo-Expert 1.2 en Visual Basic 6 de Philippe LARVET (également l'auteur du mini système expert à vocation ludique dont est issu [IAVB](https://github.com/PatriceDargenton/IAVB)).

Un système expert (SE) est un logiciel qui, à partir d'une base de règles (BR) et d'une base de faits (BF), cherche à établir des conclusions grâce à son moteur d'inférence (MI). Le principe du MI de VBBrainBox est simplement de chercher à appliquer chaque règle une fois seulement. C'est un véritable système de programmation déclarative, où les données sont séparées du code de l'application (le MI), et traitées dans un ordre quelconque, contrairement à la programmation procédurale.

Exemple du fameux syllogisme avec Socrate (logique d'ordre 1), qui illustre le classique modus ponens :
```
TOUT HOMME EST MORTEL
OR SOCRATE EST UN HOMME
DONC ?
DONC SOCRATE EST MORTEL
```
VBBrainBox est capable de manipuler des expressions logiques d'ordre 0+, c'est-à-dire du type :
Si Distance < 2 km Alors AllerAPied.

Les variables et les règles constituent la base de connaissance (BC), laquelle représente la mémoire à long terme du SE, tandis que les sessions forment la base de faits (BF), ou la mémoire à court terme (par exemple, la session "Socrate" contient bHomme = Vrai). Voici un exemple plus compliqué de rapport d'expertise qui est généré à la fin : [RapportExemple.txt](Applications/RapportExemple.txt) (si aucune session n'est sélectionnée, un rapport sur les faits initiaux de l'ensemble des sessions peut aussi être généré).

Le programme VB6 a été converti en VB .NET et un calcul de logique floue y a été ajouté, du type de celui de MYCIN conçu en... 1975 !!! Selon la configuration, la logique floue ne modifie pas le déroulement du programme, on ajoute seulement un degré de fiabilité aux règles et aux faits initiaux, et on en déduit des indices de vraisemblance pour les conclusions obtenues. Il y a cependant un mode de fonctionnement plus cohérent dans lequel l'interprétation de la logique floue peut changer le déroulement de l'expertise.

Une base de données a été ajoutée pour simplifier la création d'application, et on peut échanger des applications en exportant des petits fichiers textes de la base.

En somme, VBBrainBox = Turbo-Expert + Logique Floue + Base de données.

## Table des matières
- [Limitations](#limitations)
- [Versions](#versions)
- [Liens](#liens)

## Limitations
- Le code est une reprise du code en VB6, il n'est pas encore complètement au standard du .Net.

## Versions

Voir le [Changelog.md](Changelog.md)

## Liens

Documentation d'origine complète : [VBBrainBox : index.html](http://patrice.dargenton.free.fr/ia/vbbrainbox/index.html "La doc complète inclut des liens, notamment vers l'ancien dépôt où trouver le code VB6 de Turbo-Expert")