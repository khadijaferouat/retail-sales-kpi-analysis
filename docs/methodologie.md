# Méthodologie

## 1. Définition des indicateurs

| KPI | Formule |
|---|---|
| CA HT | Somme de CA_HT |
| Marge brute | Somme de (CA_HT moins Coût_achat) |
| Taux de marge | Marge brute / CA HT |
| Transactions | Somme de Nb_transactions |
| Panier moyen | CA HT / Transactions |
| Articles par ticket | Quantité / Transactions |
| Variation M/M | CA du mois / CA du mois précédent - 1 |

## 2. Exploration avec les tableaux croisés dynamiques

Le premier TCD (CA par mois et par magasin) montre une saisonnalité commune à tout le réseau. Le passage en « % du total de la ligne » (TCD 3) neutralise cette saisonnalité : chaque magasin garde une part du CA quasi constante d'un mois à l'autre. Lyon Part-Dieu pèse habituellement autour de 24,7 % du CA du réseau, mais seulement 21,9 % en octobre. C'est ce décrochage qui a attiré mon attention.

## 3. Quantification

Pour un magasin m et un mois t :

- ratio(m, t) = CA(m, t) / CA(reste du réseau, t)
- référence(m, t) = moyenne de ratio(m, ·) sur les 11 autres mois
- écart(m, t) = ratio(m, t) / référence(m, t) - 1

Le mois testé est exclu de sa propre référence pour qu'une anomalie ne « tire » pas la moyenne vers elle.

Résultat : écart de -15,4 % pour Lyon en octobre, alors que les 47 autres couples magasin et mois restent entre -4 % et +4 %.

## 4. Diagnostic

CA = Transactions × Panier moyen, donc (1 + écart CA) = (1 + écart transactions) × (1 + écart panier).

Avec -15,4 % sur le CA et -15,4 % sur les transactions, l'écart de panier moyen est nul. La baisse est donc un problème de fréquentation.

## 5. Limites

Les données sont synthétiques. Sur des données réelles, il faudrait compléter l'analyse avec le détail journalier (pour dater précisément le début de l'incident), les informations terrain du magasin et vérifier qu'il ne s'agit pas d'un défaut de collecte des données.
