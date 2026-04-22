# MCP Outlook — Emails & Calendrier Microsoft 365

Connecte Claude Desktop / Claude Code à Microsoft Outlook via Microsoft Graph API (OAuth2). Permet de lire, envoyer des emails et gérer le calendrier directement depuis Claude.

## Prérequis

- Node.js 18+ ([nodejs.org](https://nodejs.org))
- Un compte Microsoft 365 (Outlook professionnel)
- Les credentials Azure fournis par Anne — services@serumandco.com

---

## Installation

### Étape 1 — Cloner le repo

```bash
git clone https://github.com/serumandco/outlook-mcp.git %USERPROFILE%\.claude\mcp-servers\outlook
```

> Sur Mac/Linux : remplacer `%USERPROFILE%` par `~`

### Étape 2 — Installer les dépendances

```bash
cd %USERPROFILE%\.claude\mcp-servers\outlook
npm install
```

### Étape 3 — Créer le fichier `.env`

Créer un fichier `.env` dans le dossier `outlook/` (demander les valeurs à Anne) :

```env
CLIENT_ID=your_client_id
TENANT_ID=your_tenant_id
CLIENT_SECRET=your_client_secret
REDIRECT_URI=http://localhost:3333/callback
SCOPES=Mail.Read,Mail.Send,User.Read,Calendars.ReadWrite
```

---

## Configuration Claude Desktop

Ouvrir le fichier de config Claude Desktop :

- **Windows** : `%APPDATA%\Claude\claude_desktop_config.json`
- **Mac** : `~/Library/Application Support/Claude/claude_desktop_config.json`

Ajouter la section `mcpServers` :

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["C:/Users/VOTRE_NOM/.claude/mcp-servers/outlook/index.js"]
    }
  }
}
```

> ⚠️ Remplacer `VOTRE_NOM` par votre nom d'utilisateur Windows.  
> Sur Mac : `"/Users/VOTRE_NOM/.claude/mcp-servers/outlook/index.js"`

Redémarrer Claude Desktop — le MCP `outlook` doit apparaître dans les outils disponibles (icône 🔌 en bas de la fenêtre).

---

## Configuration Claude Code (CLI)

```bash
claude mcp add outlook node %USERPROFILE%/.claude/mcp-servers/outlook/index.js
```

---

## Première utilisation — Authentification

Au premier démarrage, se connecter avec le tool `authenticate` :

```
Utilise le tool authenticate pour te connecter à Outlook
```

Un navigateur s'ouvre sur la page de login Microsoft. Se connecter avec le compte pro Serum & Co. Le token est ensuite sauvegardé localement (pas besoin de ré-authentifier à chaque session).

---

## Outils disponibles

| Outil | Description |
|---|---|
| `authenticate` | Démarre l'authentification OAuth2 — ouvre le navigateur |
| `complete_authentication` | Finalise l'auth avec le code de retour |
| `list_emails` | Liste les emails (filtrable par dossier, nombre, recherche) |
| `read_email` | Lit le contenu complet d'un email par son ID |
| `search_emails` | Recherche dans tout Outlook par mots-clés |
| `send_email` | Envoie un email depuis Outlook |
| `create_event` | Crée un événement dans le calendrier |
| `list_events` | Liste les événements du calendrier |
| `delete_event` | Supprime un événement du calendrier |

---

## Exemples d'utilisation

```
Lis mes 10 derniers emails non lus
Cherche les emails de pascal@cefri.fr
Envoie un email à john@example.com : objet "Suivi projet", message "Bonjour..."
Crée un RDV "Réunion client MICHALAK" vendredi à 14h, durée 1h
Liste mes événements de cette semaine
```

---

## Dépannage

| Problème | Solution |
|---|---|
| `ERREUR : Variables manquantes dans le fichier .env` | Vérifier que le fichier `.env` existe dans le dossier `outlook/` |
| `Cannot find module` | Relancer `npm install` dans le dossier |
| Le MCP n'apparaît pas dans Claude Desktop | Vérifier le chemin dans `claude_desktop_config.json` + redémarrer Claude |
| Tokens expirés | Relancer `authenticate` pour se reconnecter |
| Le navigateur ne s'ouvre pas | Copier l'URL affichée dans les logs Claude et l'ouvrir manuellement |
