# Développement

## Environnement

```bat
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python main.py
```

## Moteur gelé

`engine/engine_v1.py` contient le moteur Factur-X validé. Il ne doit pas être modifié lors des évolutions d'interface.

## Build exécutable

```bat
build_exe.bat
```
