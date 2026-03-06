# Memory Game Pro

Version refactorisee et stabilisee du jeu de memoire.

## Lancement

```powershell
cd memory_game
python run.py
```

## Fonctionnalites

- Interface modernisee et plus ergonomique
- Grilles configurables (`4x4`, `4x5`, `5x6`, `6x6`)
- Themes d'images (`Animaux`, `Themes`, `Vehicules`, `Mixte`)
- Musique de fond MP3 (toggle son + volume)
- Statistiques en direct (temps, coups, erreurs, precision)
- Leaderboard persistant (SQLite local)
- Parametres sauvegardes (joueur, grille, theme, son, vitesse de flip)

## Structure du projet

```text
memory_game/
  app/
    game.py
    main.py
    storage.py
  assets/
    audio/
    images/
      animals/
      topics/
      vehicles/
      ui/
  data/
    memory_game.db
  legacy/
    code/
    db/
    misc/
    output/
  run.py
  requirements.txt
```

## Notes

- Les anciens scripts ont ete deplaces vers `legacy/` pour conserver l'historique.
- La base de scores est creee automatiquement dans `data/memory_game.db`.
