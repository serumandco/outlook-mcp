# MCP Outlook — Emails & Calendrier Microsoft 365

Connecte Claude Code à Microsoft Outlook via Microsoft Graph API (OAuth2). Permet de lire, envoyer des emails et gérer le calendrier directement depuis Claude Code.

## Prérequis

- Node.js 18+
- Un compte Microsoft 365 (Outlook professionnel ou personnel)
- Les credentials Azure fournis par Serum & Co (fichier `.env`)

## Installation

### 1. Copier le dossier

Copier le dossier `outlook-deploy/` à l'emplacement suivant :

```
~/.claude/mcp-servers/outlook/
```

### 2. Installer les dépendances

```bash
cd ~/.claude/mcp-servers/outlook
npm install
```

### 3. Créer le fichier `.env`

Créer un fichier `.env` dans le dossier `outlook/` avec les credentials Azure :

```env
CLIENT_ID=your_client_id
TENANT_ID=your_tenant_id
CLIENT_SECRET=your_client_secret
REDIRECT_URI=http://localhost:3333/callback
SCOPES=Mail.Read,Mail.Send,User.Read,Calendars.ReadWrite
```

> Les credentials Azure sont fournis par Anne (services@serumandco.com) — ne pas les mettre dans un fichier versionné.

### 4. Ajouter le MCP à Claude Code

```bash
claude mcp add outlook node ~/.claude/mcp-servers/outlook/index.js
```

Ou ajouter manuellement dans `~/.claude/settings.json` :

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["~/.claude/mcp-servers/outlook/index.js"]
    }
  }
}
```

### 5. Authentification (première utilisation)

Au premier démarrage, utiliser l'outil `authenticate` depuis Claude Code :

```
Utilise le tool authenticate pour te connecter à Outlook
```

Un navigateur s'ouvre pour le login Microsoft — se connecter avec son compte professionnel Serum & Co.

## Outils disponibles

| Outil | Description |
|---|---|
| `authenticate` | Authentification OAuth2 — ouvre le navigateur pour se connecter |
| `complete_authentication` | Finalise l'auth avec le code de retour |
| `list_emails` | Liste les emails (filtrable par dossier, nombre, recherche) |
| `read_email` | Lit le contenu complet d'un email par son ID |
| `search_emails` | Recherche dans tout Outlook par mots-clés |
| `send_email` | Envoie un email depuis Outlook |
| `create_event` | Crée un événement dans le calendrier |
| `list_events` | Liste les événements du calendrier |
| `delete_event` | Supprime un événement du calendrier |

## Utilisation

```
Lis mes 10 derniers emails non lus
Envoie un email à john@example.com avec pour objet "Test" et le message "Bonjour"
Crée un RDV "Réunion client" demain à 14h pendant 1 heure
Liste mes événements de la semaine prochaine
```

## Dépannage

- **"Tokens expirés"** : relancer `authenticate` pour se reconnecter
- **"CLIENT_ID manquant"** : vérifier que le fichier `.env` existe dans le bon dossier
- **Le navigateur ne s'ouvre pas** : copier l'URL affichée dans le terminal et l'ouvrir manuellement
