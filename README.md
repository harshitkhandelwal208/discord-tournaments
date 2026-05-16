# 🏁 HCR2 Tournament Tracker — Discord Bot

A Discord bot for managing **Hill Climb Racing 2** tournament registrations.
Players upload their in-game **driver's license / profile screenshot** and the bot
automatically extracts their name and team using **Google Gemini Vision AI**,
then renames their Discord nickname to match their in-game identity.

---

## Features

| Feature | Description |
|---|---|
| `/register` | Player uploads a driver's license screenshot → Gemini extracts name + team → nickname auto-set |
| `/unregister` | Player withdraws from the tournament; original nickname restored |
| `/mystats` | Player views their own registration info (private) |
| `/players` | Public list of all registered players |
| `/stats` | Server-wide registration statistics |
| `/export` | **[Organizer]** Download all registrations as a formatted `.xlsx` Excel file |
| `/remove_player` | **[Organizer]** Remove a player and restore their nickname |
| `/add_note` | **[Organizer]** Attach a note to a player's record |
| `/lookup` | **[Organizer]** View full registration details for any member |
| `/tournament_create` | **[Organizer]** Create a named tournament |
| `/tournament_list` | List all tournaments |

---

## Setup

### 1. Prerequisites

- Python 3.10 or newer
- A [Discord Bot Token](https://discord.com/developers/applications)
- A [Google Gemini API Key](https://aistudio.google.com/app/apikey) (free tier available)

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure environment variables

```bash
cp .env.example .env
# Edit .env and fill in your DISCORD_TOKEN and GEMINI_API_KEY
```

**Linux / macOS:**
```bash
export DISCORD_TOKEN="your_token_here"
export GEMINI_API_KEY="your_key_here"
```

**Windows (PowerShell):**
```powershell
$env:DISCORD_TOKEN = "your_token_here"
$env:GEMINI_API_KEY = "your_key_here"
```

Or use [python-dotenv](https://pypi.org/project/python-dotenv/) — add `from dotenv import load_dotenv; load_dotenv()` at the top of `bot.py`.

### 4. Discord Bot Permissions

When inviting your bot, ensure the following permissions are enabled:

**Bot Permissions:**
- `Manage Nicknames`
- `Send Messages`
- `Embed Links`
- `Attach Files`
- `Read Message History`
- `View Channels`
- `Use Application Commands`

**Privileged Gateway Intents (enable in the Developer Portal):**
- `Server Members Intent`
- `Message Content Intent`

### 5. Configure the bot (top of `bot.py`)

```python
# Roles that can use organizer-only commands
ORGANIZER_ROLE_NAMES = {"Organizer", "Admin", "Moderator", "Tournament Host"}

# Channel where /register is allowed (set to None to allow anywhere)
REGISTER_CHANNEL_NAME = "tournament-registration"
```

Create a `#tournament-registration` channel and a `#tournament-log` channel in your server.

### 6. Run

```bash
python bot.py
```

---

## Discord Server Setup Checklist

- [ ] Create role: `Organizer` (or whatever matches `ORGANIZER_ROLE_NAMES`)
- [ ] Create channel: `#tournament-registration` (restrict who can send here if desired)
- [ ] Create channel: `#tournament-log` (bot posts public registration announcements here)
- [ ] Invite the bot with the required permissions
- [ ] Make sure the bot's role is **above** regular member roles so it can rename nicknames

---

## How Registration Works

1. Player goes to `#tournament-registration` and runs `/register`
2. They attach their **in-game profile screenshot** (driver's license card)
3. The bot downloads the image and sends it to **Gemini 1.5 Flash Vision**
4. Gemini returns the player's in-game name and team name as JSON
5. The data is saved to a local **SQLite** database (`hcr2_tournament.db`)
6. The bot renames the player's Discord nickname to their in-game name
7. A confirmation embed is sent to the player (ephemeral) and a summary to `#tournament-log`

---

## Excel Export

Organizers use `/export` to download a `.xlsx` file containing:

- **Registrations sheet** — styled table with all player data, auto-filter, frozen header
- **Summary sheet** — total count, active players, unique teams, export timestamp

---

## Database

The bot uses a local **SQLite** file (`hcr2_tournament.db`). Tables:

- `registrations` — one row per player (unique by Discord ID), with status `active | withdrawn | removed`
- `tournaments` — named tournament records
- `tournament_entries` — many-to-many between tournaments and registrations

---

## Customisation Tips

- **Multiple tournaments**: Use `/tournament_create` then `/tournament_enter` (extend the code) to let players join specific tournaments
- **Role assignment on registration**: Add `await member.add_roles(registered_role)` inside `/register`
- **Hosting**: Works on any VPS, Railway, Render, or a Raspberry Pi. Use `screen` or `systemd` to keep it alive
- **python-dotenv**: `pip install python-dotenv` and add `load_dotenv()` for automatic `.env` loading
