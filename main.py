import os
import requests
import time
import pandas as pd
from dotenv import load_dotenv

# Charger les variables d'environnement du fichier .env
load_dotenv()

# Récupérer les clés API depuis les variables d'environnement
AZURE_OPENAI_KEY = os.getenv("AZURE_OPENAI_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")

# Template du mail
MAIL_TEMPLATE = """Bonjour [Prénom du client],
J'espère que ce message vous trouve en pleine forme.
En tant que [Votre poste] chez IBM, je consulte régulièrement les comptes de nos clients pour identifier les opportunités où notre expertise pourrait apporter une réelle valeur ajoutée. J'ai noté dans votre dossier une opportunité en cours en [Pays], et je souhaiterais vous proposer mon soutien pour vous aider à la concrétiser.
**Comment pouvons-nous collaborer ?**
* **Coordination avec nos équipes locales** : Grâce à notre présence mondiale, nos experts en [secteur/métier pertinent, ex: "cloud" ou "transformation digitale"] basés en [Pays/Région] peuvent intervenir pour renforcer votre proposition.
* **Connaissances locales** : Nous partagerons des insights sur les attentes du marché, les spécificités réglementaires ou culturelles, et les bonnes pratiques pour maximiser vos chances de succès.
* **Ressources sur mesure** : Que ce soit une démo technique, une étude de cas locale ou un accompagnement en négociation, nous adaptons notre soutien à vos besoins.
Seriez-vous disponible pour un échange téléphonique ou visio cette semaine afin d'affiner une stratégie commune ? Je suis également en copie de ce mail [collègue local, ex: "notre directeur commercial en Allemagne"] pour une coordination fluide.
En vous remerciant pour la confiance que vous accordez à IBM, je reste à votre disposition pour toute question.
Bien cordialement, [Votre nom complet] [Votre poste] | IBM [Téléphone] | [Email] [LinkedIn – optionnel]"""

def lire_clients_excel(file_path="clients_ibm.xlsx"):
    """
    Lit le fichier Excel et retourne une liste de dictionnaires avec les infos clients
    
    Args:
        file_path (str): Chemin vers le fichier Excel
    
    Returns:
        list: Liste de dictionnaires contenant les informations des clients
    """
    clients = []
    try:
        # Lecture du fichier Excel
        df = pd.read_excel(file_path)
        
        # Conversion en liste de dictionnaires
        for _, row in df.iterrows():
            # Extraire le prénom du contact (si possible)
            nom_complet = str(row.get('Contact', ''))
            prenom = nom_complet.split(' ')[0] if nom_complet else ''
            
            client = {
                "nom": row.get('Entreprise', ''),
                "prenom": prenom,
                "nom_contact": nom_complet,
                "poste_contact": row.get('Poste', ''),
                "pays": row.get('Pays', ''),
                "ville": row.get('Ville', ''),
                "secteur": row.get('Secteur', ''),
                "description": row.get('Description', '')
            }
            clients.append(client)
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier Excel: {e}")
    
    return clients

def generer_mail_personnalise(client, utilisateur):
    """
    Utilise Azure OpenAI pour personnaliser le template de mail pour un client spécifique
    
    Args:
        client (dict): Informations du client
        utilisateur (dict): Informations de l'expéditeur
    
    Returns:
        str: Le mail personnalisé, ou None en cas d'erreur
    """
    # Configuration pour l'API Azure OpenAI
    api_version = "2023-12-01-preview"  # Version récente de l'API
    deployment_id = "gpt-4o"  # Nom du modèle déployé
    
    headers = {
        "Content-Type": "application/json",
        "api-key": AZURE_OPENAI_KEY
    }
    
    # Prompt pour GPT-4o - enrichi avec les nouvelles informations
    prompt = f"""
    Je dois personnaliser un email professionnel pour un client. Voici les informations détaillées:

    Informations sur le client:
    - Entreprise: {client['nom']}
    - Secteur d'activité: {client['secteur']}
    - Pays: {client['pays']}
    - Ville: {client['ville']}
    - Contact principal: {client['nom_contact']} ({client['poste_contact']})
    - Description de l'entreprise/besoins: {client.get('description', 'Non disponible')}

    Informations sur moi (expéditeur):
    - Nom complet: {utilisateur['nom_complet']}
    - Poste: {utilisateur['poste']}
    - Téléphone: {utilisateur['telephone']}
    - Email: {utilisateur['email']}
    - LinkedIn: {utilisateur.get('linkedin', '')}

    Voici le template du mail à personnaliser en remplaçant les éléments entre crochets:
    {MAIL_TEMPLATE}

    Remplace uniquement le contenu des crochets avec des informations pertinentes et personnalisées pour ce client.
    Assure-toi d'adapter le message au secteur d'activité du client et à ses besoins spécifiques mentionnés dans la description.
    Pour [secteur/métier pertinent], choisis quelque chose qui correspond vraiment à son secteur d'activité.
    Pour [Pays/Région], utilise le pays du client mentionné ci-dessus.
    Pour [collègue local], invente un nom plausible qui pourrait être un directeur commercial dans le pays du client.
    
    Retourne uniquement le mail personnalisé, sans commentaires supplémentaires.
    """
    
    data = {
        "messages": [
            {"role": "system", "content": "Tu es un assistant spécialisé dans la personnalisation d'emails commerciaux pour IBM."},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.7,
        "max_tokens": 1000
    }
    
    try:
        response = requests.post(
            f"{AZURE_OPENAI_ENDPOINT}/openai/deployments/{deployment_id}/chat/completions?api-version={api_version}",
            headers=headers,
            json=data
        )
        
        if response.status_code == 200:
            result = response.json()
            mail_personnalise = result["choices"][0]["message"]["content"]
            return mail_personnalise
        else:
            print(f"Erreur API Azure OpenAI: {response.status_code} - {response.text}")
            return None
    except Exception as e:
        print(f"Erreur lors de la génération du mail personnalisé: {e}")
        return None

def sauvegarder_mail(client, mail_contenu, dossier_sortie="mails_personnalises"):
    """
    Sauvegarde le mail personnalisé dans un fichier
    
    Args:
        client (dict): Informations du client
        mail_contenu (str): Contenu du mail personnalisé
        dossier_sortie (str): Dossier de destination pour les mails
    """
    # Créer le dossier de sortie s'il n'existe pas
    if not os.path.exists(dossier_sortie):
        os.makedirs(dossier_sortie)
    
    # Créer un nom de fichier unique basé sur le nom de l'entreprise
    nom_fichier = f"{dossier_sortie}/{client['nom'].replace(' ', '_')}_{int(time.time())}.txt"
    
    try:
        with open(nom_fichier, 'w', encoding='utf-8') as f:
            f.write(mail_contenu)
        print(f"Mail sauvegardé avec succès dans {nom_fichier}")
        return True
    except Exception as e:
        print(f"Erreur lors de la sauvegarde du mail: {e}")
        return False

def main():
    # Informations sur l'expéditeur (à personnaliser)
    expediteur = {
        "nom_complet": "David Vuong",
        "poste": "Responsable commercial",
        "telephone": "+33 7 83 48 45 99",
        "email": "david.sv04@ibm.com",
        "linkedin": "linkedin.com/j'aipasmonlinkedn"
    }
    
    # Lire les informations des clients depuis Excel
    clients = lire_clients_excel("clients_ibm.xlsx")
    
    if not clients:
        print("Aucun client trouvé dans le fichier Excel ou fichier non accessible.")
        return
    
    print(f"Génération de mails personnalisés pour {len(clients)} clients...")
    
    # Générer un mail personnalisé pour chaque client
    for i, client in enumerate(clients):
        print(f"Traitement du client {i+1}/{len(clients)}: {client['nom']}")
        
        # Générer le mail personnalisé
        mail_personnalise = generer_mail_personnalise(client, expediteur)
        
        if mail_personnalise:
            # Sauvegarder le mail
            sauvegarder_mail(client, mail_personnalise)
        else:
            print(f"Échec de la génération du mail pour {client['nom']}")
        
        # Petite pause pour éviter de surcharger l'API
        time.sleep(1)
    
    print("Traitement terminé.")

if __name__ == "__main__":
    main()