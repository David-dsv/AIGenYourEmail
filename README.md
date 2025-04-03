# AIGenYourEmail

## Description
AIGenYourEmail est un outil d'automatisation basé sur l'API Azure OpenAI qui génère des e-mails professionnels personnalisés pour les clients d'IBM. En utilisant un modèle GPT-4o, cet outil permet d'adapter dynamiquement un template d'e-mail en fonction des informations spécifiques de chaque client.

## Fonctionnalités
- **Lecture des clients depuis un fichier Excel** : Extraction des informations des clients à partir d'un fichier `clients_ibm.xlsx`.
- **Génération automatique d'e-mails personnalisés** : Utilisation de GPT-4o pour adapter le template en fonction du secteur, du pays et des besoins du client.
- **Sauvegarde des e-mails générés** : Stockage des e-mails personnalisés sous forme de fichiers `.txt`.
- **Exportation des données clients en format Excel** : Extraction et structuration des données à partir d'un fichier `.txt` en `.xlsx` avec mise en forme avancée.

## Technologies utilisées
- **Python**
- **Azure OpenAI API (GPT-4o)**
- **Pandas** pour la manipulation de fichiers Excel
- **Requests** pour interagir avec l'API Azure
- **Dotenv** pour la gestion des clés API
- **Openpyxl** pour le formatage Excel avancé

## Installation
1. **Cloner le dépôt**
   ```sh
   git clone https://github.com/David-dsv/AIGenYourEmail.git
   cd AIGenYourEmail
   ```
2. **Créer un environnement virtuel et l'activer**
   ```sh
   python3 -m venv env
   source env/bin/activate  # Pour macOS/Linux
   env\Scripts\activate    # Pour Windows
   ```
3. **Installer les dépendances**
   ```sh
   pip install -r requirements.txt
   ```
4. **Créer un fichier `.env` et ajouter vos clés API Azure**
   ```ini
   AZURE_OPENAI_KEY=VotreCléAPI
   AZURE_OPENAI_ENDPOINT=VotreEndpointAzure
   ```

## Utilisation
1. **Exécuter le script principal**
   ```sh
   python main.py
   ```
2. **Les e-mails générés seront enregistrés dans le dossier `mails_personnalises/`**
3. **Les informations clients seront exportées en `clients_ibm.xlsx`**

## Exemples d'e-mails générés
Un e-mail typique pourrait ressembler à ceci :
```
Bonjour M. Dupont,
J'espère que ce message vous trouve en pleine forme...
...
Bien cordialement,
David Vuong
Responsable commercial | IBM
```

## Contribuer
1. **Forker le projet**
2. **Créer une branche feature**
   ```sh
   git checkout -b ma-nouvelle-feature
   ```
3. **Faire une pull request**

## Auteurs
- **David Vuong** - [GitHub](https://github.com/David-dsv)

## Licence
Ce projet est sous licence MIT - voir le fichier `LICENSE` pour plus de détails.
