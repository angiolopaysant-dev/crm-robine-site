# CRM Robine — site

Application web statique (PWA) servie par GitHub Pages.

- Toutes les données sont **chiffrées** (AES-256-GCM) : `data/data.json` et `data/f/*`.
  Elles sont produites et poussées automatiquement par une GitHub Action d’un dépôt privé ; ne pas les modifier à la main.
- `users.json` ne contient que la clé de données enveloppée pour chaque compte (PBKDF2-SHA256, 600 000 itérations, sel unique)
  et un annuaire chiffré. Aucun mot de passe ni aucune donnée n’y est lisible.
- `outils/amorcage.html` : outil local de première installation (clé de données + premier compte).

Aucune donnée en clair ne doit jamais être ajoutée à ce dépôt public.
