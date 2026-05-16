"""
HCR2 Tournament Tracker Bot
============================
A Discord bot for managing Hill Climb Racing 2 tournament registrations.
Uses Gemini Vision API to auto-extract player info from driver's license screenshots.
"""

import discord
from discord import app_commands
from discord.ext import commands
import sqlite3
import os
import io
import aiohttp
import json
import base64
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import logging
import asyncio
from typing import Optional
import re

# ─── Logging ────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("hcr2_bot.log", encoding="utf-8"),
    ],
)
log = logging.getLogger("HCR2Bot")

# ─── Config ─────────────────────────────────────────────────────────────────
DISCORD_TOKEN = os.getenv("DISCORD_TOKEN", "YOUR_DISCORD_BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY")
GEMINI_API_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "gemini-1.5-flash:generateContent"
)

# Role names — change to match your server's roles
ORGANIZER_ROLE_NAMES = {"Organizer", "Admin", "Moderator", "Tournament Host"}

# Registration channel name (optional gate: only allow /register in this channel)
REGISTER_CHANNEL_NAME = "tournament-registration"  # set to None to allow anywhere

# Role assigned to every successfully registered player.
# Right-click the role in Discord → Copy Role ID, then paste it here.
# Set to None to disable automatic role assignment.
REGISTERED_ROLE_ID: int | None = 123456789012345678  # ← REPLACE WITH YOUR ROLE ID

DB_PATH = "hcr2_tournament.db"

# ─── Database ────────────────────────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS registrations (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            discord_id       TEXT    NOT NULL UNIQUE,
            discord_username TEXT    NOT NULL,
            ingame_name      TEXT    NOT NULL,
            team_name        TEXT,
            original_nickname TEXT,
            registered_at    TEXT    NOT NULL,
            updated_at       TEXT,
            status           TEXT    NOT NULL DEFAULT 'active',
            notes            TEXT
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS tournaments (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            name         TEXT    NOT NULL,
            created_at   TEXT    NOT NULL,
            created_by   TEXT    NOT NULL,
            status       TEXT    NOT NULL DEFAULT 'open',
            description  TEXT
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS tournament_entries (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            tournament_id   INTEGER NOT NULL,
            registration_id INTEGER NOT NULL,
            entered_at      TEXT    NOT NULL,
            FOREIGN KEY (tournament_id)   REFERENCES tournaments(id),
            FOREIGN KEY (registration_id) REFERENCES registrations(id),
            UNIQUE(tournament_id, registration_id)
        )
    """)

    conn.commit()
    conn.close()
    log.info("Database initialised.")


def db_connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

# ─── Gemini Vision Helper ────────────────────────────────────────────────────

async def extract_player_info_from_image(image_bytes: bytes, mime_type: str = "image/jpeg") -> dict:
    """
    Send the driver's license screenshot to Gemini Vision and parse the result.
    Returns: {"ingame_name": str, "team_name": str|None, "raw": str, "success": bool}
    """
    b64_image = base64.b64encode(image_bytes).decode("utf-8")

    prompt = """You are analyzing a Hill Climb Racing 2 (HCR2) in-game player profile / driver's license screenshot.

Your task is to extract:
1. The player's exact in-game username / display name (shown prominently at the top of the profile card).
2. The team/club name (shown below the username, sometimes with a ™ symbol or in a coloured banner). If the player has no team, return null.

Rules:
- The in-game name may contain clan/team tags like "DC|", "TM-", "[TAG]" as a prefix — include the full name exactly as shown.
- Team/club name is the name of the club banner below the username (e.g. "Discord 3™"). Do NOT confuse rank names (Legendary, Challenger) with team names.
- Return ONLY a valid JSON object with keys "ingame_name" and "team_name". No markdown, no extra text.

Example output:
{"ingame_name": "DC|BlackWing", "team_name": "Discord 3™"}

If you cannot read the image clearly, return:
{"ingame_name": null, "team_name": null}
"""

    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt},
                    {
                        "inline_data": {
                            "mime_type": mime_type,
                            "data": b64_image,
                        }
                    },
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.1,
            "maxOutputTokens": 256,
        },
    }

    async with aiohttp.ClientSession() as session:
        url = f"{GEMINI_API_URL}?key={GEMINI_API_KEY}"
        async with session.post(url, json=payload) as resp:
            if resp.status != 200:
                text = await resp.text()
                log.error("Gemini API error %s: %s", resp.status, text)
                return {"success": False, "ingame_name": None, "team_name": None, "raw": text}

            data = await resp.json()

    try:
        raw_text = data["candidates"][0]["content"]["parts"][0]["text"].strip()
        # Strip potential markdown code fences
        raw_text = re.sub(r"^```[a-z]*\n?", "", raw_text)
        raw_text = re.sub(r"\n?```$", "", raw_text).strip()
        parsed = json.loads(raw_text)
        return {
            "success": parsed.get("ingame_name") is not None,
            "ingame_name": parsed.get("ingame_name"),
            "team_name": parsed.get("team_name"),
            "raw": raw_text,
        }
    except Exception as e:
        log.error("Failed to parse Gemini response: %s — raw: %s", e, data)
        return {"success": False, "ingame_name": None, "team_name": None, "raw": str(data)}

# ─── Excel Export Helper ─────────────────────────────────────────────────────

def build_excel_export(rows, sheet_title="Registrations") -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title

    # Colours
    HEADER_BG  = "1E3A5F"
    HEADER_FG  = "FFFFFF"
    ALT_ROW_BG = "EBF2FA"
    BORDER_CLR = "AAAAAA"

    thin = Side(style="thin", color=BORDER_CLR)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = [
        "#", "Discord Username", "Discord ID", "In-Game Name",
        "Team / Club", "Registered At", "Status", "Notes",
    ]

    # ── Header row ──────────────────────────────────────────────────────────
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True, color=HEADER_FG, name="Calibri", size=11)
        cell.fill = PatternFill("solid", fgColor=HEADER_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    ws.row_dimensions[1].height = 24

    # ── Data rows ────────────────────────────────────────────────────────────
    for row_idx, r in enumerate(rows, 2):
        values = [
            row_idx - 1,
            r["discord_username"],
            r["discord_id"],
            r["ingame_name"],
            r["team_name"] or "—",
            r["registered_at"],
            r["status"].capitalize(),
            r["notes"] or "",
        ]
        alt = (row_idx % 2 == 0)
        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=(col_idx == len(headers)))
            cell.border = border
            if alt:
                cell.fill = PatternFill("solid", fgColor=ALT_ROW_BG)

    # ── Column widths ────────────────────────────────────────────────────────
    col_widths = [5, 24, 20, 24, 22, 22, 10, 30]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Freeze header ────────────────────────────────────────────────────────
    ws.freeze_panes = "A2"

    # ── Auto-filter ──────────────────────────────────────────────────────────
    ws.auto_filter.ref = ws.dimensions

    # ── Summary sheet ────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "HCR2 Tournament — Registration Summary"
    ws2["A1"].font = Font(bold=True, size=14, name="Calibri", color=HEADER_BG)
    ws2["A3"] = "Generated at:"
    ws2["B3"] = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
    ws2["A4"] = "Total registrations:"
    ws2["B4"] = len(rows)
    ws2["A5"] = "Active players:"
    ws2["B5"] = sum(1 for r in rows if r["status"] == "active")
    ws2["A6"] = "Unique teams:"
    teams = set(r["team_name"] for r in rows if r["team_name"])
    ws2["B6"] = len(teams)
    for cell in ["A3", "A4", "A5", "A6"]:
        ws2[cell].font = Font(bold=True, name="Calibri")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ─── Bot Setup ───────────────────────────────────────────────────────────────

intents = discord.Intents.default()
intents.members = True
intents.message_content = True

bot = commands.Bot(command_prefix="!", intents=intents)

def is_organizer(interaction: discord.Interaction) -> bool:
    if not interaction.guild:
        return False
    member = interaction.guild.get_member(interaction.user.id)
    if not member:
        return False
    return any(r.name in ORGANIZER_ROLE_NAMES for r in member.roles) or member.guild_permissions.administrator

def organizer_check():
    async def predicate(interaction: discord.Interaction) -> bool:
        if not is_organizer(interaction):
            await interaction.response.send_message(
                "❌ You need the **Organizer** role to use this command.", ephemeral=True
            )
            return False
        return True
    return app_commands.check(predicate)

# ─── /register ───────────────────────────────────────────────────────────────

@bot.tree.command(name="register", description="Register for the HCR2 tournament by uploading your driver's license screenshot.")
@app_commands.describe(screenshot="Your in-game driver's license / profile screenshot")
async def register(interaction: discord.Interaction, screenshot: discord.Attachment):
    # Optional channel gate
    if REGISTER_CHANNEL_NAME and interaction.channel.name != REGISTER_CHANNEL_NAME:
        await interaction.response.send_message(
            f"❌ Please use <#{discord.utils.get(interaction.guild.channels, name=REGISTER_CHANNEL_NAME).id}> to register.",
            ephemeral=True,
        )
        return

    if not screenshot.content_type or not screenshot.content_type.startswith("image/"):
        await interaction.response.send_message("❌ Please attach a valid image file (PNG or JPG).", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True, thinking=True)

    # Download the attachment
    try:
        image_bytes = await screenshot.read()
    except Exception as e:
        await interaction.followup.send(f"❌ Failed to download your screenshot: {e}", ephemeral=True)
        return

    # Gemini extraction
    await interaction.followup.send("🔍 Analysing your driver's license with Gemini AI...", ephemeral=True)
    mime = screenshot.content_type or "image/jpeg"
    result = await extract_player_info_from_image(image_bytes, mime)

    if not result["success"] or not result["ingame_name"]:
        await interaction.followup.send(
            "❌ Could not read your in-game profile from the screenshot.\n\n"
            "**Tips:**\n"
            "• Make sure the screenshot shows the **driver's licence / player profile card** clearly.\n"
            "• The image must show your **in-game name** at the top.\n"
            "• Try a cleaner, higher-resolution screenshot.\n\n"
            f"*(Raw Gemini response: `{result.get('raw', 'none')[:200]}`)*",
            ephemeral=True,
        )
        return

    ingame_name = result["ingame_name"].strip()
    team_name   = (result["team_name"] or "").strip() or None

    # Save to DB
    now = datetime.utcnow().isoformat(timespec="seconds")
    discord_id  = str(interaction.user.id)
    discord_tag = str(interaction.user)

    conn = db_connect()
    c = conn.cursor()

    # Check existing registration
    existing = c.execute(
        "SELECT * FROM registrations WHERE discord_id = ?", (discord_id,)
    ).fetchone()

    original_nick = interaction.user.display_name if not existing else existing["original_nickname"]

    if existing:
        c.execute("""
            UPDATE registrations
            SET ingame_name = ?, team_name = ?, discord_username = ?, updated_at = ?, status = 'active'
            WHERE discord_id = ?
        """, (ingame_name, team_name, discord_tag, now, discord_id))
        action = "updated"
    else:
        c.execute("""
            INSERT INTO registrations
                (discord_id, discord_username, ingame_name, team_name, original_nickname, registered_at, status)
            VALUES (?, ?, ?, ?, ?, ?, 'active')
        """, (discord_id, discord_tag, ingame_name, team_name, original_nick, now))
        action = "registered"

    conn.commit()
    conn.close()

    # Rename Discord nickname
    rename_status = ""
    try:
        new_nick = ingame_name[:32]  # Discord nick limit
        await interaction.user.edit(nick=new_nick)
        rename_status = f"✅ Your server nickname has been set to **{new_nick}**."
    except discord.Forbidden:
        rename_status = "⚠️ I couldn't rename your nickname (missing permissions or you're a server owner)."
    except Exception as e:
        rename_status = f"⚠️ Nickname update failed: {e}"

    # Assign registered role
    role_status = ""
    if REGISTERED_ROLE_ID:
        role = interaction.guild.get_role(REGISTERED_ROLE_ID)
        if role is None:
            role_status = f"⚠️ Registered role (ID `{REGISTERED_ROLE_ID}`) not found in this server."
            log.warning("REGISTERED_ROLE_ID %s not found in guild %s", REGISTERED_ROLE_ID, interaction.guild.id)
        elif role in interaction.user.roles:
            role_status = f"✅ You already have the **{role.name}** role."
        else:
            try:
                await interaction.user.add_roles(role, reason="HCR2 tournament registration")
                role_status = f"✅ You've been given the **{role.name}** role."
            except discord.Forbidden:
                role_status = f"⚠️ I don't have permission to assign the **{role.name}** role."
            except Exception as e:
                role_status = f"⚠️ Role assignment failed: {e}"
                log.error("Role assignment error for %s: %s", interaction.user, e)

    # Success embed
    embed = discord.Embed(
        title="🏁 Registration Successful!",
        color=discord.Color.green(),
        timestamp=datetime.utcnow(),
    )
    embed.add_field(name="In-Game Name", value=f"`{ingame_name}`", inline=True)
    embed.add_field(name="Team / Club",  value=f"`{team_name}`" if team_name else "*None*", inline=True)
    embed.add_field(name="Status",       value=f"{'Updated ✏️' if action == 'updated' else 'New entry ✨'}", inline=True)
    embed.add_field(name="Nickname",     value=rename_status, inline=False)
    if role_status:
        embed.add_field(name="Role",     value=role_status,   inline=False)
    embed.set_thumbnail(url=screenshot.url)
    embed.set_footer(text=f"Discord: {discord_tag} | {action.capitalize()} at {now} UTC")

    await interaction.followup.send(embed=embed, ephemeral=True)

    # Log to a public channel if it exists
    log_ch = discord.utils.get(interaction.guild.text_channels, name="tournament-log")
    if log_ch:
        pub = discord.Embed(
            title="🎮 New Tournament Registration",
            color=discord.Color.blue(),
            timestamp=datetime.utcnow(),
        )
        pub.add_field(name="Player",      value=interaction.user.mention, inline=True)
        pub.add_field(name="In-Game Name", value=f"`{ingame_name}`",       inline=True)
        pub.add_field(name="Team",         value=team_name or "—",         inline=True)
        await log_ch.send(embed=pub)

# ─── /unregister ─────────────────────────────────────────────────────────────

@bot.tree.command(name="unregister", description="Remove yourself from the tournament.")
async def unregister(interaction: discord.Interaction):
    discord_id = str(interaction.user.id)
    conn = db_connect()
    row = conn.execute("SELECT * FROM registrations WHERE discord_id = ?", (discord_id,)).fetchone()
    if not row:
        conn.close()
        await interaction.response.send_message("❌ You are not registered in any tournament.", ephemeral=True)
        return

    conn.execute("UPDATE registrations SET status = 'withdrawn' WHERE discord_id = ?", (discord_id,))
    conn.commit()
    conn.close()

    # Restore original nickname
    try:
        original = row["original_nickname"] or ""
        await interaction.user.edit(nick=original if original else None)
    except discord.Forbidden:
        pass

    # Remove registered role
    if REGISTERED_ROLE_ID:
        role = interaction.guild.get_role(REGISTERED_ROLE_ID)
        if role and role in interaction.user.roles:
            try:
                await interaction.user.remove_roles(role, reason="HCR2 tournament withdrawal")
            except discord.Forbidden:
                pass

    await interaction.response.send_message(
        "✅ You have been **withdrawn** from the tournament. Your nickname has been restored.",
        ephemeral=True,
    )

# ─── /mystats ────────────────────────────────────────────────────────────────

@bot.tree.command(name="mystats", description="View your current registration info.")
async def mystats(interaction: discord.Interaction):
    discord_id = str(interaction.user.id)
    conn = db_connect()
    row = conn.execute("SELECT * FROM registrations WHERE discord_id = ?", (discord_id,)).fetchone()
    conn.close()

    if not row:
        await interaction.response.send_message("❌ You are not registered. Use `/register` to join.", ephemeral=True)
        return

    embed = discord.Embed(title="📋 Your Registration", color=discord.Color.blue(), timestamp=datetime.utcnow())
    embed.add_field(name="In-Game Name", value=f"`{row['ingame_name']}`",           inline=True)
    embed.add_field(name="Team / Club",  value=row["team_name"] or "—",             inline=True)
    embed.add_field(name="Status",       value=row["status"].capitalize(),           inline=True)
    embed.add_field(name="Registered",  value=row["registered_at"],                 inline=True)
    if row["updated_at"]:
        embed.add_field(name="Last Updated", value=row["updated_at"],               inline=True)
    await interaction.response.send_message(embed=embed, ephemeral=True)

# ─── /players ────────────────────────────────────────────────────────────────

@bot.tree.command(name="players", description="List all registered tournament players.")
@app_commands.describe(status="Filter by status (active / withdrawn / all)")
async def players(interaction: discord.Interaction, status: Optional[str] = "active"):
    status = (status or "active").lower()
    conn = db_connect()
    if status == "all":
        rows = conn.execute("SELECT * FROM registrations ORDER BY registered_at").fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM registrations WHERE status = ? ORDER BY registered_at", (status,)
        ).fetchall()
    conn.close()

    if not rows:
        await interaction.response.send_message(f"No registrations found (status: `{status}`).", ephemeral=True)
        return

    # Paginate at 20 per embed
    lines = []
    for i, r in enumerate(rows):
        team_part = f"• *{r['team_name']}*" if r['team_name'] else ''
        lines.append(f"`{i+1:02d}.` **{r['ingame_name']}** {team_part} — <@{r['discord_id']}>")
    chunks = [lines[i:i+20] for i in range(0, len(lines), 20)]

    embeds = []
    for page, chunk in enumerate(chunks, 1):
        e = discord.Embed(
            title=f"🏆 Registered Players ({status.capitalize()}) — Page {page}/{len(chunks)}",
            description="\n".join(chunk),
            color=discord.Color.gold(),
            timestamp=datetime.utcnow(),
        )
        e.set_footer(text=f"Total: {len(rows)} player(s)")
        embeds.append(e)

    # Send first page (paginator not included for brevity; extend if needed)
    await interaction.response.send_message(embed=embeds[0])
    for e in embeds[1:]:
        await interaction.channel.send(embed=e)

# ─── /export ─────────────────────────────────────────────────────────────────

@bot.tree.command(name="export", description="[Organizer] Export all registrations to an Excel file.")
@app_commands.describe(status="Filter by status (active / withdrawn / all)")
@organizer_check()
async def export(interaction: discord.Interaction, status: Optional[str] = "all"):
    await interaction.response.defer(ephemeral=True, thinking=True)

    status = (status or "all").lower()
    conn = db_connect()
    if status == "all":
        rows = conn.execute("SELECT * FROM registrations ORDER BY registered_at").fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM registrations WHERE status = ? ORDER BY registered_at", (status,)
        ).fetchall()
    conn.close()

    if not rows:
        await interaction.followup.send("❌ No registrations found.", ephemeral=True)
        return

    dicts = [dict(r) for r in rows]
    excel_bytes = build_excel_export(dicts, sheet_title="Registrations")

    filename = f"HCR2_Tournament_{datetime.utcnow().strftime('%Y%m%d_%H%M')}.xlsx"
    file = discord.File(io.BytesIO(excel_bytes), filename=filename)

    embed = discord.Embed(
        title="📊 Export Ready",
        description=f"**{len(rows)}** registration(s) exported (status: `{status}`).",
        color=discord.Color.green(),
        timestamp=datetime.utcnow(),
    )
    await interaction.followup.send(embed=embed, file=file, ephemeral=True)

# ─── /remove_player ──────────────────────────────────────────────────────────

@bot.tree.command(name="remove_player", description="[Organizer] Remove a player's registration.")
@app_commands.describe(member="The Discord member to remove", reason="Reason for removal")
@organizer_check()
async def remove_player(interaction: discord.Interaction, member: discord.Member, reason: Optional[str] = "Removed by organizer"):
    discord_id = str(member.id)
    conn = db_connect()
    row = conn.execute("SELECT * FROM registrations WHERE discord_id = ?", (discord_id,)).fetchone()
    if not row:
        conn.close()
        await interaction.response.send_message(f"❌ {member.mention} is not registered.", ephemeral=True)
        return

    conn.execute(
        "UPDATE registrations SET status = 'removed', notes = ? WHERE discord_id = ?",
        (reason, discord_id),
    )
    conn.commit()
    conn.close()

    # Restore nickname
    try:
        original = row["original_nickname"] or ""
        await member.edit(nick=original if original else None)
    except discord.Forbidden:
        pass

    # Remove registered role
    if REGISTERED_ROLE_ID:
        role = interaction.guild.get_role(REGISTERED_ROLE_ID)
        if role and role in member.roles:
            try:
                await member.remove_roles(role, reason=f"HCR2 removal: {reason}")
            except discord.Forbidden:
                pass

    await interaction.response.send_message(
        f"✅ **{member.mention}** (`{row['ingame_name']}`) has been removed. Reason: *{reason}*",
        ephemeral=True,
    )

# ─── /tournament_create ──────────────────────────────────────────────────────

@bot.tree.command(name="tournament_create", description="[Organizer] Create a new tournament.")
@app_commands.describe(name="Tournament name", description="Short description")
@organizer_check()
async def tournament_create(interaction: discord.Interaction, name: str, description: Optional[str] = None):
    now = datetime.utcnow().isoformat(timespec="seconds")
    conn = db_connect()
    conn.execute(
        "INSERT INTO tournaments (name, created_at, created_by, status, description) VALUES (?,?,?,?,?)",
        (name, now, str(interaction.user), "open", description),
    )
    conn.commit()
    tid = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.close()

    embed = discord.Embed(
        title="🏁 Tournament Created!",
        color=discord.Color.green(),
        timestamp=datetime.utcnow(),
    )
    embed.add_field(name="Name",        value=name,                   inline=True)
    embed.add_field(name="ID",          value=f"`{tid}`",             inline=True)
    embed.add_field(name="Status",      value="Open 🟢",              inline=True)
    embed.add_field(name="Description", value=description or "—",     inline=False)
    await interaction.response.send_message(embed=embed)

# ─── /tournament_list ────────────────────────────────────────────────────────

@bot.tree.command(name="tournament_list", description="List all tournaments.")
async def tournament_list(interaction: discord.Interaction):
    conn = db_connect()
    rows = conn.execute("SELECT * FROM tournaments ORDER BY created_at DESC").fetchall()
    conn.close()

    if not rows:
        await interaction.response.send_message("No tournaments have been created yet.", ephemeral=True)
        return

    embed = discord.Embed(title="📋 Tournaments", color=discord.Color.blurple(), timestamp=datetime.utcnow())
    for t in rows:
        status_icon = {"open": "🟢", "closed": "🔴", "ongoing": "🟡"}.get(t["status"], "⚪")
        embed.add_field(
            name=f"{status_icon} [{t['id']}] {t['name']}",
            value=f"Status: **{t['status']}** | Created: {t['created_at']} by {t['created_by']}",
            inline=False,
        )
    await interaction.response.send_message(embed=embed)

# ─── /add_note ───────────────────────────────────────────────────────────────

@bot.tree.command(name="add_note", description="[Organizer] Add a note to a player's registration.")
@app_commands.describe(member="Discord member", note="Note to add")
@organizer_check()
async def add_note(interaction: discord.Interaction, member: discord.Member, note: str):
    conn = db_connect()
    result = conn.execute(
        "UPDATE registrations SET notes = ? WHERE discord_id = ?",
        (note, str(member.id)),
    )
    conn.commit()
    conn.close()

    if result.rowcount == 0:
        await interaction.response.send_message(f"❌ {member.mention} is not registered.", ephemeral=True)
    else:
        await interaction.response.send_message(
            f"✅ Note added for {member.mention}: *{note}*", ephemeral=True
        )

# ─── /lookup ─────────────────────────────────────────────────────────────────

@bot.tree.command(name="lookup", description="[Organizer] Look up a player's registration by Discord member.")
@app_commands.describe(member="Discord member to look up")
@organizer_check()
async def lookup(interaction: discord.Interaction, member: discord.Member):
    conn = db_connect()
    row = conn.execute("SELECT * FROM registrations WHERE discord_id = ?", (str(member.id),)).fetchone()
    conn.close()

    if not row:
        await interaction.response.send_message(f"❌ {member.mention} is not registered.", ephemeral=True)
        return

    embed = discord.Embed(title=f"🔎 Player Lookup: {member.display_name}", color=discord.Color.blue(), timestamp=datetime.utcnow())
    embed.set_thumbnail(url=member.display_avatar.url)
    embed.add_field(name="In-Game Name",      value=f"`{row['ingame_name']}`",       inline=True)
    embed.add_field(name="Team / Club",       value=row["team_name"] or "—",         inline=True)
    embed.add_field(name="Status",            value=row["status"].capitalize(),       inline=True)
    embed.add_field(name="Discord Username",  value=row["discord_username"],          inline=True)
    embed.add_field(name="Discord ID",        value=row["discord_id"],               inline=True)
    embed.add_field(name="Registered At",     value=row["registered_at"],            inline=True)
    if row["updated_at"]:
        embed.add_field(name="Last Updated", value=row["updated_at"],               inline=True)
    if row["notes"]:
        embed.add_field(name="Notes",        value=row["notes"],                    inline=False)
    await interaction.response.send_message(embed=embed, ephemeral=True)

# ─── /stats ──────────────────────────────────────────────────────────────────

@bot.tree.command(name="stats", description="Show overall tournament registration statistics.")
async def stats(interaction: discord.Interaction):
    conn = db_connect()
    total    = conn.execute("SELECT COUNT(*) FROM registrations").fetchone()[0]
    active   = conn.execute("SELECT COUNT(*) FROM registrations WHERE status='active'").fetchone()[0]
    withdrawn= conn.execute("SELECT COUNT(*) FROM registrations WHERE status='withdrawn'").fetchone()[0]
    removed  = conn.execute("SELECT COUNT(*) FROM registrations WHERE status='removed'").fetchone()[0]
    teams    = conn.execute("SELECT COUNT(DISTINCT team_name) FROM registrations WHERE team_name IS NOT NULL AND status='active'").fetchone()[0]
    top_teams= conn.execute("""
        SELECT team_name, COUNT(*) as cnt FROM registrations
        WHERE status='active' AND team_name IS NOT NULL
        GROUP BY team_name ORDER BY cnt DESC LIMIT 5
    """).fetchall()
    conn.close()

    embed = discord.Embed(
        title="📊 Tournament Registration Statistics",
        color=discord.Color.gold(),
        timestamp=datetime.utcnow(),
    )
    embed.add_field(name="Total Registrations", value=str(total),    inline=True)
    embed.add_field(name="Active Players",       value=str(active),   inline=True)
    embed.add_field(name="Unique Teams",         value=str(teams),    inline=True)
    embed.add_field(name="Withdrawn",            value=str(withdrawn),inline=True)
    embed.add_field(name="Removed",              value=str(removed),  inline=True)

    if top_teams:
        top_str = "\n".join(f"`{t['team_name']}` — {t['cnt']} player(s)" for t in top_teams)
        embed.add_field(name="Top Teams", value=top_str, inline=False)

    await interaction.response.send_message(embed=embed)

# ─── Events ──────────────────────────────────────────────────────────────────

@bot.event
async def on_ready():
    log.info("Logged in as %s (ID: %s)", bot.user, bot.user.id)
    init_db()
    try:
        synced = await bot.tree.sync()
        log.info("Synced %d slash commands.", len(synced))
    except Exception as e:
        log.error("Failed to sync commands: %s", e)
    await bot.change_presence(
        activity=discord.Activity(
            type=discord.ActivityType.watching,
            name="🏁 HCR2 Tournament | /register"
        )
    )

@bot.event
async def on_app_command_error(interaction: discord.Interaction, error: app_commands.AppCommandError):
    if isinstance(error, app_commands.CheckFailure):
        return  # Already handled inside the check
    log.error("Slash command error: %s", error, exc_info=True)
    msg = f"❌ An error occurred: `{error}`"
    if interaction.response.is_done():
        await interaction.followup.send(msg, ephemeral=True)
    else:
        await interaction.response.send_message(msg, ephemeral=True)

# ─── Entry Point ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if DISCORD_TOKEN == "YOUR_DISCORD_BOT_TOKEN":
        print("ERROR: Set the DISCORD_TOKEN environment variable before running.")
    elif GEMINI_API_KEY == "YOUR_GEMINI_API_KEY":
        print("ERROR: Set the GEMINI_API_KEY environment variable before running.")
    else:
        bot.run(DISCORD_TOKEN, log_handler=None)
