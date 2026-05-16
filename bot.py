"""
HCR2 Tournament Tracker Bot
============================
A Discord bot for managing Hill Climb Racing 2 tournament registrations.
Uses Gemini Vision API to auto-extract player info from driver's license screenshots.
Includes a web dashboard for viewing registrations and configuring roles.
"""

import discord
from discord import app_commands
from discord.ext import commands
import sqlite3
import os
import io
import aiohttp
from aiohttp import web as aio_web
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
DISCORD_TOKEN   = os.getenv("DISCORD_TOKEN",   "YOUR_DISCORD_BOT_TOKEN")
GEMINI_API_KEY  = os.getenv("GEMINI_API_KEY",  "YOUR_GEMINI_API_KEY")
GEMINI_API_URL  = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "gemini-1.5-flash:generateContent"
)

DASHBOARD_PORT = int(os.getenv("DASHBOARD_PORT", "8080"))
DASHBOARD_HOST = os.getenv("DASHBOARD_HOST", "0.0.0.0")

# Fallback organizer role names used when no DB config exists for a guild
ORGANIZER_ROLE_NAMES = {"Organizer", "Admin", "Moderator", "Tournament Host"}

# Registration channel name — set to None to allow /register anywhere
REGISTER_CHANNEL_NAME = "tournament-registration"

DB_PATH = "hcr2_tournament.db"

# ─── Database ────────────────────────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS registrations (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            discord_id        TEXT    NOT NULL UNIQUE,
            discord_username  TEXT    NOT NULL,
            ingame_name       TEXT    NOT NULL,
            team_name         TEXT,
            original_nickname TEXT,
            registered_at     TEXT    NOT NULL,
            updated_at        TEXT,
            status            TEXT    NOT NULL DEFAULT 'active'
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

    c.execute("""
        CREATE TABLE IF NOT EXISTS guild_config (
            guild_id            TEXT PRIMARY KEY,
            organizer_role_ids  TEXT NOT NULL DEFAULT '[]',
            participant_role_id TEXT
        )
    """)

    conn.commit()
    conn.close()
    log.info("Database initialised.")


def db_connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def get_guild_config(guild_id: str) -> dict:
    conn = db_connect()
    row = conn.execute(
        "SELECT * FROM guild_config WHERE guild_id = ?", (guild_id,)
    ).fetchone()
    conn.close()
    if row:
        return {
            "organizer_role_ids": json.loads(row["organizer_role_ids"] or "[]"),
            "participant_role_id": row["participant_role_id"],
        }
    return {"organizer_role_ids": [], "participant_role_id": None}

# ─── Gemini Vision Helper ────────────────────────────────────────────────────

async def extract_player_info_from_image(image_bytes: bytes, mime_type: str = "image/jpeg") -> dict:
    """Send driver's license screenshot to Gemini Vision and parse the result."""
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
                    {"inline_data": {"mime_type": mime_type, "data": b64_image}},
                ]
            }
        ],
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": 256},
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

    HEADER_BG  = "1E3A5F"
    HEADER_FG  = "FFFFFF"
    ALT_ROW_BG = "EBF2FA"
    BORDER_CLR = "AAAAAA"

    thin   = Side(style="thin", color=BORDER_CLR)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["#", "Discord Username", "Discord ID", "In-Game Name", "Team / Club", "Registered At", "Status"]

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font      = Font(bold=True, color=HEADER_FG, name="Calibri", size=11)
        cell.fill      = PatternFill("solid", fgColor=HEADER_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = border

    ws.row_dimensions[1].height = 24

    for row_idx, r in enumerate(rows, 2):
        values = [
            row_idx - 1,
            r["discord_username"],
            r["discord_id"],
            r["ingame_name"],
            r["team_name"] or "—",
            r["registered_at"],
            r["status"].capitalize(),
        ]
        alt = (row_idx % 2 == 0)
        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font      = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="center")
            cell.border    = border
            if alt:
                cell.fill = PatternFill("solid", fgColor=ALT_ROW_BG)

    col_widths = [5, 24, 20, 24, 22, 22, 10]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

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
    ws2["B6"] = len(set(r["team_name"] for r in rows if r["team_name"]))
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
    if member.guild_permissions.administrator:
        return True
    config = get_guild_config(str(interaction.guild.id))
    configured_ids = set(config["organizer_role_ids"])
    if configured_ids:
        return any(str(r.id) in configured_ids for r in member.roles)
    # Fallback to static role names when guild has no DB config
    return any(r.name in ORGANIZER_ROLE_NAMES for r in member.roles)


def organizer_check():
    async def predicate(interaction: discord.Interaction) -> bool:
        if not is_organizer(interaction):
            await interaction.response.send_message(
                "❌ You need an **Organizer** role to use this command.", ephemeral=True
            )
            return False
        return True
    return app_commands.check(predicate)

# ─── /register ───────────────────────────────────────────────────────────────

@bot.tree.command(name="register", description="Register for the HCR2 tournament by uploading your driver's license screenshot.")
@app_commands.describe(screenshot="Your in-game driver's license / profile screenshot")
async def register(interaction: discord.Interaction, screenshot: discord.Attachment):
    if REGISTER_CHANNEL_NAME and interaction.channel.name != REGISTER_CHANNEL_NAME:
        ch = discord.utils.get(interaction.guild.channels, name=REGISTER_CHANNEL_NAME)
        await interaction.response.send_message(
            f"❌ Please use <#{ch.id}> to register.", ephemeral=True
        )
        return

    if not screenshot.content_type or not screenshot.content_type.startswith("image/"):
        await interaction.response.send_message("❌ Please attach a valid image file (PNG or JPG).", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True, thinking=True)

    try:
        image_bytes = await screenshot.read()
    except Exception as e:
        await interaction.followup.send(f"❌ Failed to download your screenshot: {e}", ephemeral=True)
        return

    await interaction.followup.send("🔍 Analysing your driver's license with Gemini AI...", ephemeral=True)
    mime   = screenshot.content_type or "image/jpeg"
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

    now        = datetime.utcnow().isoformat(timespec="seconds")
    discord_id = str(interaction.user.id)
    discord_tag= str(interaction.user)

    conn = db_connect()
    c    = conn.cursor()

    existing      = c.execute("SELECT * FROM registrations WHERE discord_id = ?", (discord_id,)).fetchone()
    original_nick = interaction.user.display_name if not existing else existing["original_nickname"]

    if existing:
        c.execute("""
            UPDATE registrations
            SET ingame_name=?, team_name=?, discord_username=?, updated_at=?, status='active'
            WHERE discord_id=?
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

    rename_status = ""
    try:
        await interaction.user.edit(nick=ingame_name[:32])
        rename_status = f"✅ Your server nickname has been set to **{ingame_name[:32]}**."
    except discord.Forbidden:
        rename_status = "⚠️ I couldn't rename your nickname (missing permissions or you're a server owner)."
    except Exception as e:
        rename_status = f"⚠️ Nickname update failed: {e}"

    role_status = ""
    config      = get_guild_config(str(interaction.guild.id))
    p_role_id   = config.get("participant_role_id")
    if p_role_id:
        role = interaction.guild.get_role(int(p_role_id))
        if role:
            try:
                await interaction.user.add_roles(role, reason="HCR2 tournament registration")
                role_status = f"✅ You have been given the **{role.name}** role."
            except discord.Forbidden:
                role_status = "⚠️ I couldn't assign your role (missing permissions)."
            except Exception as e:
                role_status = f"⚠️ Role assignment failed: {e}"
                log.error("Role assignment error for %s: %s", interaction.user, e)

    embed = discord.Embed(
        title="🏁 Registration Successful!",
        color=discord.Color.green(),
        timestamp=datetime.utcnow(),
    )
    embed.add_field(name="In-Game Name", value=f"`{ingame_name}`", inline=True)
    embed.add_field(name="Team / Club",  value=f"`{team_name}`" if team_name else "*None*", inline=True)
    embed.add_field(name="Status",       value="Updated ✏️" if action == "updated" else "New entry ✨", inline=True)
    embed.add_field(name="Nickname",     value=rename_status, inline=False)
    if role_status:
        embed.add_field(name="Role", value=role_status, inline=False)
    embed.set_thumbnail(url=screenshot.url)
    embed.set_footer(text=f"Discord: {discord_tag} | {action.capitalize()} at {now} UTC")

    await interaction.followup.send(embed=embed, ephemeral=True)

    log_ch = discord.utils.get(interaction.guild.text_channels, name="tournament-log")
    if log_ch:
        pub = discord.Embed(title="🎮 New Tournament Registration", color=discord.Color.blue(), timestamp=datetime.utcnow())
        pub.add_field(name="Player",       value=interaction.user.mention, inline=True)
        pub.add_field(name="In-Game Name", value=f"`{ingame_name}`",       inline=True)
        pub.add_field(name="Team",         value=team_name or "—",         inline=True)
        await log_ch.send(embed=pub)

# ─── /unregister ─────────────────────────────────────────────────────────────

@bot.tree.command(name="unregister", description="Remove yourself from the tournament.")
async def unregister(interaction: discord.Interaction):
    discord_id = str(interaction.user.id)
    conn = db_connect()
    row  = conn.execute("SELECT * FROM registrations WHERE discord_id = ?", (discord_id,)).fetchone()
    if not row:
        conn.close()
        await interaction.response.send_message("❌ You are not registered in any tournament.", ephemeral=True)
        return

    conn.execute("UPDATE registrations SET status='withdrawn' WHERE discord_id=?", (discord_id,))
    conn.commit()
    conn.close()

    try:
        original = row["original_nickname"] or ""
        await interaction.user.edit(nick=original or None)
    except discord.Forbidden:
        pass

    config    = get_guild_config(str(interaction.guild.id))
    p_role_id = config.get("participant_role_id")
    if p_role_id:
        role = interaction.guild.get_role(int(p_role_id))
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
    conn = db_connect()
    row  = conn.execute("SELECT * FROM registrations WHERE discord_id=?", (str(interaction.user.id),)).fetchone()
    conn.close()

    if not row:
        await interaction.response.send_message("❌ You are not registered. Use `/register` to join.", ephemeral=True)
        return

    embed = discord.Embed(title="📋 Your Registration", color=discord.Color.blue(), timestamp=datetime.utcnow())
    embed.add_field(name="In-Game Name", value=f"`{row['ingame_name']}`", inline=True)
    embed.add_field(name="Team / Club",  value=row["team_name"] or "—",   inline=True)
    embed.add_field(name="Status",       value=row["status"].capitalize(), inline=True)
    embed.add_field(name="Registered",   value=row["registered_at"],       inline=True)
    if row["updated_at"]:
        embed.add_field(name="Last Updated", value=row["updated_at"], inline=True)
    await interaction.response.send_message(embed=embed, ephemeral=True)

# ─── /players ────────────────────────────────────────────────────────────────

@bot.tree.command(name="players", description="List all registered tournament players.")
@app_commands.describe(status="Filter by status (active / withdrawn / all)")
async def players(interaction: discord.Interaction, status: Optional[str] = "active"):
    status = (status or "active").lower()
    conn   = db_connect()
    if status == "all":
        rows = conn.execute("SELECT * FROM registrations ORDER BY registered_at").fetchall()
    else:
        rows = conn.execute("SELECT * FROM registrations WHERE status=? ORDER BY registered_at", (status,)).fetchall()
    conn.close()

    if not rows:
        await interaction.response.send_message(f"No registrations found (status: `{status}`).", ephemeral=True)
        return

    lines  = []
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
    conn   = db_connect()
    rows   = conn.execute(
        "SELECT * FROM registrations ORDER BY registered_at"
        if status == "all"
        else "SELECT * FROM registrations WHERE status=? ORDER BY registered_at",
        () if status == "all" else (status,)
    ).fetchall()
    conn.close()

    if not rows:
        await interaction.followup.send("❌ No registrations found.", ephemeral=True)
        return

    excel_bytes = build_excel_export([dict(r) for r in rows])
    filename    = f"HCR2_Tournament_{datetime.utcnow().strftime('%Y%m%d_%H%M')}.xlsx"
    file        = discord.File(io.BytesIO(excel_bytes), filename=filename)

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
async def remove_player(
    interaction: discord.Interaction,
    member: discord.Member,
    reason: Optional[str] = "Removed by organizer"
):
    discord_id = str(member.id)
    conn = db_connect()
    row  = conn.execute("SELECT * FROM registrations WHERE discord_id=?", (discord_id,)).fetchone()
    if not row:
        conn.close()
        await interaction.response.send_message(f"❌ {member.mention} is not registered.", ephemeral=True)
        return

    conn.execute("UPDATE registrations SET status='removed' WHERE discord_id=?", (discord_id,))
    conn.commit()
    conn.close()

    try:
        original = row["original_nickname"] or ""
        await member.edit(nick=original or None)
    except discord.Forbidden:
        pass

    config    = get_guild_config(str(interaction.guild.id))
    p_role_id = config.get("participant_role_id")
    if p_role_id:
        role = interaction.guild.get_role(int(p_role_id))
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

    embed = discord.Embed(title="🏁 Tournament Created!", color=discord.Color.green(), timestamp=datetime.utcnow())
    embed.add_field(name="Name",        value=name,               inline=True)
    embed.add_field(name="ID",          value=f"`{tid}`",         inline=True)
    embed.add_field(name="Status",      value="Open 🟢",          inline=True)
    embed.add_field(name="Description", value=description or "—", inline=False)
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
        icon = {"open": "🟢", "closed": "🔴", "ongoing": "🟡"}.get(t["status"], "⚪")
        embed.add_field(
            name=f"{icon} [{t['id']}] {t['name']}",
            value=f"Status: **{t['status']}** | Created: {t['created_at']} by {t['created_by']}",
            inline=False,
        )
    await interaction.response.send_message(embed=embed)

# ─── /lookup ─────────────────────────────────────────────────────────────────

@bot.tree.command(name="lookup", description="[Organizer] Look up a player's registration.")
@app_commands.describe(member="Discord member to look up")
@organizer_check()
async def lookup(interaction: discord.Interaction, member: discord.Member):
    conn = db_connect()
    row  = conn.execute("SELECT * FROM registrations WHERE discord_id=?", (str(member.id),)).fetchone()
    conn.close()

    if not row:
        await interaction.response.send_message(f"❌ {member.mention} is not registered.", ephemeral=True)
        return

    embed = discord.Embed(title=f"🔎 Player Lookup: {member.display_name}", color=discord.Color.blue(), timestamp=datetime.utcnow())
    embed.set_thumbnail(url=member.display_avatar.url)
    embed.add_field(name="In-Game Name",     value=f"`{row['ingame_name']}`",  inline=True)
    embed.add_field(name="Team / Club",      value=row["team_name"] or "—",    inline=True)
    embed.add_field(name="Status",           value=row["status"].capitalize(),  inline=True)
    embed.add_field(name="Discord Username", value=row["discord_username"],     inline=True)
    embed.add_field(name="Discord ID",       value=row["discord_id"],           inline=True)
    embed.add_field(name="Registered At",    value=row["registered_at"],        inline=True)
    if row["updated_at"]:
        embed.add_field(name="Last Updated", value=row["updated_at"], inline=True)
    await interaction.response.send_message(embed=embed, ephemeral=True)

# ─── /stats ──────────────────────────────────────────────────────────────────

@bot.tree.command(name="stats", description="Show overall tournament registration statistics.")
async def stats(interaction: discord.Interaction):
    conn      = db_connect()
    total     = conn.execute("SELECT COUNT(*) FROM registrations").fetchone()[0]
    active    = conn.execute("SELECT COUNT(*) FROM registrations WHERE status='active'").fetchone()[0]
    withdrawn = conn.execute("SELECT COUNT(*) FROM registrations WHERE status='withdrawn'").fetchone()[0]
    removed   = conn.execute("SELECT COUNT(*) FROM registrations WHERE status='removed'").fetchone()[0]
    teams     = conn.execute(
        "SELECT COUNT(DISTINCT team_name) FROM registrations WHERE team_name IS NOT NULL AND status='active'"
    ).fetchone()[0]
    top_teams = conn.execute("""
        SELECT team_name, COUNT(*) as cnt FROM registrations
        WHERE status='active' AND team_name IS NOT NULL
        GROUP BY team_name ORDER BY cnt DESC LIMIT 5
    """).fetchall()
    conn.close()

    embed = discord.Embed(title="📊 Tournament Registration Statistics", color=discord.Color.gold(), timestamp=datetime.utcnow())
    embed.add_field(name="Total Registrations", value=str(total),    inline=True)
    embed.add_field(name="Active Players",       value=str(active),   inline=True)
    embed.add_field(name="Unique Teams",         value=str(teams),    inline=True)
    embed.add_field(name="Withdrawn",            value=str(withdrawn),inline=True)
    embed.add_field(name="Removed",              value=str(removed),  inline=True)
    if top_teams:
        embed.add_field(
            name="Top Teams",
            value="\n".join(f"`{t['team_name']}` — {t['cnt']} player(s)" for t in top_teams),
            inline=False,
        )
    await interaction.response.send_message(embed=embed)

# ─── Dashboard HTML ───────────────────────────────────────────────────────────

DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>HCR2 Tournament Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;500;600;700&family=IBM+Plex+Mono:ital,wght@0,400;0,500;1,400&display=swap" rel="stylesheet">
<style>
:root {
  --bg:      #09090b;
  --surface: #111115;
  --card:    #16161a;
  --border:  #26262e;
  --accent:  #f59e0b;
  --accent2: #b45309;
  --text:    #e4e0d8;
  --muted:   #5a5a68;
  --green:   #22c55e;
  --red:     #ef4444;
  --yellow:  #eab308;
  --blue:    #60a5fa;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
html, body { height: 100%; }
body {
  font-family: 'Rajdhani', sans-serif;
  background: var(--bg);
  color: var(--text);
  display: flex;
  min-height: 100vh;
}
/* ── Sidebar ── */
#sidebar {
  width: 220px;
  min-height: 100vh;
  background: var(--surface);
  border-right: 1px solid var(--border);
  display: flex;
  flex-direction: column;
  padding: 0;
  position: fixed;
  top: 0; left: 0; bottom: 0;
  z-index: 10;
}
.logo {
  padding: 24px 20px 20px;
  border-bottom: 1px solid var(--border);
}
.logo-mark {
  font-size: 26px;
  font-weight: 700;
  letter-spacing: 0.04em;
  color: var(--accent);
  line-height: 1;
}
.logo-sub {
  font-size: 10px;
  color: var(--muted);
  letter-spacing: 0.14em;
  text-transform: uppercase;
  margin-top: 5px;
  font-family: 'IBM Plex Mono', monospace;
}
.nav-section {
  padding: 16px 0 8px;
}
.nav-label {
  font-size: 10px;
  color: var(--muted);
  text-transform: uppercase;
  letter-spacing: 0.14em;
  padding: 0 20px;
  margin-bottom: 4px;
  font-family: 'IBM Plex Mono', monospace;
}
.nav-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 9px 20px;
  cursor: pointer;
  color: var(--muted);
  font-size: 15px;
  font-weight: 500;
  letter-spacing: 0.03em;
  transition: all 0.12s;
  border-left: 3px solid transparent;
  user-select: none;
}
.nav-item:hover { color: var(--text); background: rgba(255,255,255,0.03); }
.nav-item.active {
  color: var(--accent);
  border-left-color: var(--accent);
  background: rgba(245,158,11,0.07);
}
.nav-icon { font-size: 16px; opacity: 0.8; }
.sidebar-footer {
  margin-top: auto;
  padding: 14px 20px;
  border-top: 1px solid var(--border);
  font-size: 11px;
  color: var(--muted);
  font-family: 'IBM Plex Mono', monospace;
  line-height: 1.6;
}
/* ── Main ── */
#main {
  margin-left: 220px;
  flex: 1;
  padding: 36px 44px;
  min-width: 0;
}
.section { display: none; animation: fadeIn 0.18s ease; }
.section.active { display: block; }
@keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: none; } }
/* ── Page header ── */
.page-header {
  margin-bottom: 28px;
  padding-bottom: 20px;
  border-bottom: 1px solid var(--border);
  display: flex;
  align-items: baseline;
  gap: 16px;
}
.page-title {
  font-size: 30px;
  font-weight: 700;
  letter-spacing: 0.03em;
}
.page-subtitle {
  font-size: 12px;
  color: var(--muted);
  font-family: 'IBM Plex Mono', monospace;
}
/* ── Stat cards ── */
.stats-grid {
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 14px;
  margin-bottom: 28px;
}
.stat-card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 7px;
  padding: 18px 16px;
  position: relative;
  overflow: hidden;
}
.stat-card::after {
  content: '';
  position: absolute;
  bottom: 0; left: 0; right: 0;
  height: 2px;
  background: linear-gradient(90deg, var(--accent), transparent);
}
.stat-value {
  font-size: 32px;
  font-weight: 700;
  color: var(--accent);
  line-height: 1;
  font-family: 'IBM Plex Mono', monospace;
  letter-spacing: -0.02em;
}
.stat-label {
  font-size: 11px;
  color: var(--muted);
  text-transform: uppercase;
  letter-spacing: 0.1em;
  margin-top: 7px;
  font-weight: 600;
}
/* ── Cards ── */
.overview-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
.card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 22px;
}
.card-title {
  font-size: 11px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.12em;
  color: var(--muted);
  margin-bottom: 14px;
  font-family: 'IBM Plex Mono', monospace;
}
.list-row {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 9px 0;
  border-bottom: 1px solid var(--border);
}
.list-row:last-child { border-bottom: none; }
.list-num { font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: var(--muted); width: 18px; flex-shrink: 0; }
.list-name { flex: 1; font-size: 15px; font-weight: 600; }
.list-val { font-family: 'IBM Plex Mono', monospace; font-size: 12px; color: var(--accent); }
.list-date { font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: var(--muted); }
/* ── Table ── */
.table-bar {
  display: flex;
  gap: 8px;
  margin-bottom: 14px;
  align-items: center;
  flex-wrap: wrap;
}
.filter-btn {
  padding: 5px 14px;
  border-radius: 4px;
  border: 1px solid var(--border);
  background: transparent;
  color: var(--muted);
  cursor: pointer;
  font-family: 'Rajdhani', sans-serif;
  font-size: 14px;
  font-weight: 600;
  letter-spacing: 0.04em;
  transition: all 0.12s;
}
.filter-btn:hover { border-color: var(--accent); color: var(--text); }
.filter-btn.active { border-color: var(--accent); color: var(--accent); background: rgba(245,158,11,0.08); }
.search-box {
  margin-left: auto;
  padding: 5px 12px;
  border-radius: 4px;
  border: 1px solid var(--border);
  background: var(--card);
  color: var(--text);
  font-family: 'IBM Plex Mono', monospace;
  font-size: 12px;
  outline: none;
  width: 210px;
  transition: border-color 0.12s;
}
.search-box:focus { border-color: var(--accent); }
.search-box::placeholder { color: var(--muted); }
.table-wrap {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 8px;
  overflow: hidden;
}
table { width: 100%; border-collapse: collapse; }
thead th {
  padding: 10px 16px;
  text-align: left;
  font-size: 10px;
  text-transform: uppercase;
  letter-spacing: 0.12em;
  color: var(--muted);
  font-weight: 700;
  font-family: 'IBM Plex Mono', monospace;
  border-bottom: 1px solid var(--border);
  background: var(--surface);
}
tbody tr { border-bottom: 1px solid rgba(38,38,46,0.6); transition: background 0.08s; }
tbody tr:last-child { border-bottom: none; }
tbody tr:hover { background: rgba(255,255,255,0.02); }
tbody td {
  padding: 10px 16px;
  font-size: 13px;
  font-family: 'IBM Plex Mono', monospace;
  color: var(--text);
  vertical-align: middle;
}
td.name-cell {
  font-family: 'Rajdhani', sans-serif;
  font-size: 15px;
  font-weight: 600;
  color: var(--text);
}
.badge {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 3px;
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  font-family: 'IBM Plex Mono', monospace;
}
.badge-active    { background: rgba(34,197,94,0.14); color: #4ade80; border: 1px solid rgba(34,197,94,0.28); }
.badge-withdrawn { background: rgba(234,179,8,0.12); color: #fbbf24; border: 1px solid rgba(234,179,8,0.28); }
.badge-removed   { background: rgba(239,68,68,0.12); color: #f87171; border: 1px solid rgba(239,68,68,0.28); }
.empty { text-align: center; padding: 48px; color: var(--muted); font-size: 13px; }
/* ── Settings ── */
.settings-row {
  display: grid;
  grid-template-columns: 320px 1fr;
  gap: 20px;
  margin-bottom: 20px;
}
.form-label {
  display: block;
  font-size: 10px;
  text-transform: uppercase;
  letter-spacing: 0.12em;
  color: var(--muted);
  margin-bottom: 8px;
  font-weight: 700;
  font-family: 'IBM Plex Mono', monospace;
}
.form-select {
  width: 100%;
  padding: 8px 34px 8px 12px;
  border-radius: 5px;
  border: 1px solid var(--border);
  background: var(--card);
  color: var(--text);
  font-family: 'Rajdhani', sans-serif;
  font-size: 15px;
  font-weight: 500;
  outline: none;
  cursor: pointer;
  transition: border-color 0.12s;
  appearance: none;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='10' fill='%235a5a68' viewBox='0 0 16 16'%3E%3Cpath d='M7.247 11.14L2.451 5.658C1.885 5.013 2.345 4 3.204 4h9.592a1 1 0 0 1 .753 1.659l-4.796 5.48a1 1 0 0 1-1.506 0z'/%3E%3C/svg%3E");
  background-repeat: no-repeat;
  background-position: right 12px center;
}
.form-select:focus { border-color: var(--accent); }
.settings-cards { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
.settings-card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 24px;
}
.settings-card h3 {
  font-size: 17px;
  font-weight: 700;
  margin-bottom: 6px;
  letter-spacing: 0.03em;
}
.settings-hint {
  font-size: 12px;
  color: var(--muted);
  margin-bottom: 18px;
  font-family: 'IBM Plex Mono', monospace;
  line-height: 1.55;
}
.roles-list {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 5px;
  max-height: 230px;
  overflow-y: auto;
  padding: 4px;
  scrollbar-width: thin;
  scrollbar-color: var(--border) transparent;
}
.role-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 8px 10px;
  border-radius: 4px;
  cursor: pointer;
  transition: background 0.08s;
  user-select: none;
}
.role-item:hover { background: rgba(255,255,255,0.04); }
.role-check {
  width: 14px; height: 14px;
  border: 1.5px solid var(--border);
  border-radius: 3px;
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  transition: all 0.12s;
  font-size: 9px;
  font-weight: 800;
}
.role-item.checked .role-check {
  background: var(--accent);
  border-color: var(--accent);
  color: #000;
}
.role-item.checked .role-check::after { content: '✓'; }
.role-dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }
.role-name { font-size: 13px; font-family: 'IBM Plex Mono', monospace; color: var(--text); }
.save-btn {
  margin-top: 20px;
  padding: 9px 22px;
  background: var(--accent);
  color: #000;
  font-family: 'Rajdhani', sans-serif;
  font-size: 14px;
  font-weight: 700;
  letter-spacing: 0.06em;
  text-transform: uppercase;
  border: none;
  border-radius: 5px;
  cursor: pointer;
  transition: all 0.12s;
}
.save-btn:hover { background: var(--accent2); color: var(--text); }
/* ── Toast ── */
.toast {
  position: fixed;
  bottom: 24px; right: 24px;
  padding: 11px 18px;
  background: var(--surface);
  border: 1px solid var(--green);
  border-radius: 6px;
  color: var(--green);
  font-size: 13px;
  font-family: 'IBM Plex Mono', monospace;
  opacity: 0;
  transform: translateY(8px);
  transition: all 0.25s;
  pointer-events: none;
  z-index: 999;
}
.toast.show { opacity: 1; transform: none; }
.toast.err  { border-color: var(--red); color: var(--red); }
.placeholder-text { color: var(--muted); font-size: 13px; font-family: 'IBM Plex Mono', monospace; padding: 12px 0; }
</style>
</head>
<body>

<nav id="sidebar">
  <div class="logo">
    <div class="logo-mark">🏁 HCR2</div>
    <div class="logo-sub">Tournament Dashboard</div>
  </div>
  <div class="nav-section">
    <div class="nav-label">Menu</div>
    <div class="nav-item active" onclick="nav('overview',this)">
      <span class="nav-icon">▣</span> Overview
    </div>
    <div class="nav-item" onclick="nav('registrations',this)">
      <span class="nav-icon">◈</span> Registrations
    </div>
    <div class="nav-item" onclick="nav('settings',this)">
      <span class="nav-icon">◎</span> Settings
    </div>
  </div>
  <div class="sidebar-footer" id="sidebar-status">Connecting...</div>
</nav>

<main id="main">

  <!-- OVERVIEW -->
  <section class="section active" id="sec-overview">
    <div class="page-header">
      <div class="page-title">Overview</div>
      <div class="page-subtitle" id="overview-ts">—</div>
    </div>
    <div class="stats-grid" id="stats-grid">
      <div class="placeholder-text">Loading stats...</div>
    </div>
    <div class="overview-grid">
      <div class="card">
        <div class="card-title">Top Teams</div>
        <div id="top-teams"><div class="placeholder-text">Loading...</div></div>
      </div>
      <div class="card">
        <div class="card-title">Recent Registrations</div>
        <div id="recent-list"><div class="placeholder-text">Loading...</div></div>
      </div>
    </div>
  </section>

  <!-- REGISTRATIONS -->
  <section class="section" id="sec-registrations">
    <div class="page-header">
      <div class="page-title">Registrations</div>
      <div class="page-subtitle" id="reg-subtitle">—</div>
    </div>
    <div class="table-bar">
      <button class="filter-btn active" onclick="setFilter('active',this)">Active</button>
      <button class="filter-btn" onclick="setFilter('withdrawn',this)">Withdrawn</button>
      <button class="filter-btn" onclick="setFilter('removed',this)">Removed</button>
      <button class="filter-btn" onclick="setFilter('all',this)">All</button>
      <input class="search-box" id="reg-search" type="text" placeholder="Search name / team..." oninput="renderTable()">
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>In-Game Name</th>
            <th>Team / Club</th>
            <th>Discord</th>
            <th>Registered</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody id="reg-tbody">
          <tr><td colspan="6" class="empty">Loading...</td></tr>
        </tbody>
      </table>
    </div>
  </section>

  <!-- SETTINGS -->
  <section class="section" id="sec-settings">
    <div class="page-header">
      <div class="page-title">Settings</div>
      <div class="page-subtitle">Per-server role configuration</div>
    </div>

    <div class="settings-row">
      <div>
        <label class="form-label">Server</label>
        <select class="form-select" id="guild-select" onchange="onGuildChange()">
          <option value="">Select a server...</option>
        </select>
      </div>
    </div>

    <div id="guild-cfg" style="display:none;">
      <div class="settings-cards">
        <div class="settings-card">
          <h3>Organizer Roles</h3>
          <p class="settings-hint">Members with any selected role can run organizer-only commands
            (/export, /remove_player, /lookup, etc.)</p>
          <div class="roles-list" id="org-roles">
            <div class="placeholder-text">&nbsp;&nbsp;Select a server first.</div>
          </div>
          <button class="save-btn" onclick="saveConfig()">Save Settings</button>
        </div>
        <div class="settings-card">
          <h3>Participant Role</h3>
          <p class="settings-hint">Automatically assigned to players when they successfully register
            via /register, and removed on withdrawal or removal.</p>
          <label class="form-label">Role</label>
          <select class="form-select" id="participant-select">
            <option value="">— None —</option>
          </select>
        </div>
      </div>
    </div>
  </section>

</main>

<div class="toast" id="toast"></div>

<script>
// ── State ────────────────────────────────────────────────────────────────────
let allRegs       = [];
let currentFilter = 'active';
let guildsData    = [];
let guildRoles    = [];
let orgRoles      = new Set();

// ── Navigation ───────────────────────────────────────────────────────────────
function nav(name, el) {
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.getElementById('sec-' + name).classList.add('active');
  el.classList.add('active');
  if (name === 'overview')       loadOverview();
  if (name === 'registrations')  loadRegistrations();
  if (name === 'settings')       loadGuilds();
}

// ── Toast ────────────────────────────────────────────────────────────────────
function toast(msg, err) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className   = 'toast show' + (err ? ' err' : '');
  setTimeout(() => t.classList.remove('show'), 2800);
}

function esc(s) {
  return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ── Overview ─────────────────────────────────────────────────────────────────
async function loadOverview() {
  try {
    const d = await fetch('/api/stats').then(r => r.json());

    const statsMap = [
      ['Total',     d.total],
      ['Active',    d.active],
      ['Withdrawn', d.withdrawn],
      ['Removed',   d.removed],
      ['Teams',     d.unique_teams],
    ];
    document.getElementById('stats-grid').innerHTML = statsMap.map(([l,v]) =>
      `<div class="stat-card"><div class="stat-value">${v}</div><div class="stat-label">${l}</div></div>`
    ).join('');

    document.getElementById('overview-ts').textContent = 'Updated ' + new Date().toLocaleTimeString();
    document.getElementById('sidebar-status').innerHTML =
      `<span style="color:var(--green)">● </span> Online &mdash; ${d.guilds} server${d.guilds !== 1 ? 's' : ''}`;

    const teams = document.getElementById('top-teams');
    teams.innerHTML = (d.top_teams||[]).length
      ? d.top_teams.map((t,i) =>
          `<div class="list-row">
            <span class="list-num">${i+1}</span>
            <span class="list-name">${esc(t.name)}</span>
            <span class="list-val">${t.count}p</span>
          </div>`).join('')
      : '<div class="placeholder-text">No teams yet.</div>';

    const rec = document.getElementById('recent-list');
    rec.innerHTML = (d.recent||[]).length
      ? d.recent.map(r =>
          `<div class="list-row">
            <span class="list-name">${esc(r.ingame_name)}</span>
            <span class="list-date">${r.registered_at.split('T')[0]}</span>
          </div>`).join('')
      : '<div class="placeholder-text">No registrations yet.</div>';

  } catch(e) {
    document.getElementById('stats-grid').innerHTML = '<div class="placeholder-text">Failed to load stats.</div>';
  }
}

// ── Registrations ─────────────────────────────────────────────────────────────
async function loadRegistrations() {
  document.getElementById('reg-tbody').innerHTML =
    '<tr><td colspan="6" class="empty">Loading...</td></tr>';
  try {
    allRegs = await fetch('/api/registrations').then(r => r.json());
    renderTable();
  } catch(e) {
    document.getElementById('reg-tbody').innerHTML =
      '<tr><td colspan="6" class="empty">Failed to load.</td></tr>';
  }
}

function setFilter(f, el) {
  currentFilter = f;
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  el.classList.add('active');
  renderTable();
}

function renderTable() {
  const q = document.getElementById('reg-search').value.toLowerCase();
  let rows = allRegs;
  if (currentFilter !== 'all') rows = rows.filter(r => r.status === currentFilter);
  if (q) rows = rows.filter(r =>
    r.ingame_name.toLowerCase().includes(q) ||
    (r.team_name || '').toLowerCase().includes(q) ||
    r.discord_username.toLowerCase().includes(q)
  );
  document.getElementById('reg-subtitle').textContent =
    rows.length + ' player' + (rows.length !== 1 ? 's' : '');
  const tbody = document.getElementById('reg-tbody');
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="6" class="empty">No results.</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map((r, i) => `
    <tr>
      <td style="color:var(--muted);font-size:11px">${i+1}</td>
      <td class="name-cell">${esc(r.ingame_name)}</td>
      <td style="color:var(--muted)">${esc(r.team_name || '—')}</td>
      <td style="font-size:11px;color:var(--muted)">${esc(r.discord_username)}</td>
      <td style="font-size:11px;color:var(--muted)">${(r.registered_at||'').split('T')[0]}</td>
      <td><span class="badge badge-${r.status}">${r.status}</span></td>
    </tr>`).join('');
}

// ── Settings ──────────────────────────────────────────────────────────────────
async function loadGuilds() {
  const sel = document.getElementById('guild-select');
  try {
    guildsData = await fetch('/api/guilds').then(r => r.json());
    sel.innerHTML = '<option value="">Select a server...</option>' +
      guildsData.map(g => `<option value="${esc(g.id)}">${esc(g.name)}</option>`).join('');
  } catch(e) {
    sel.innerHTML = '<option value="">Failed to load servers</option>';
  }
}

async function onGuildChange() {
  const guildId = document.getElementById('guild-select').value;
  const cfg     = document.getElementById('guild-cfg');
  if (!guildId) { cfg.style.display = 'none'; return; }
  cfg.style.display = 'block';

  const guild = guildsData.find(g => g.id === guildId);
  guildRoles  = guild ? guild.roles : [];

  let config = { organizer_role_ids: [], participant_role_id: null };
  try { config = await fetch('/api/config?guild_id=' + guildId).then(r => r.json()); } catch(e){}

  orgRoles = new Set(config.organizer_role_ids || []);
  renderOrgRoles();

  // Participant dropdown
  const pSel = document.getElementById('participant-select');
  pSel.innerHTML = '<option value="">— None —</option>' +
    guildRoles.map(ro =>
      `<option value="${esc(ro.id)}" ${config.participant_role_id === ro.id ? 'selected' : ''}>${esc(ro.name)}</option>`
    ).join('');
}

function renderOrgRoles() {
  const list = document.getElementById('org-roles');
  if (!guildRoles.length) {
    list.innerHTML = '<div class="placeholder-text">&nbsp;&nbsp;No roles found.</div>';
    return;
  }
  list.innerHTML = guildRoles.map(ro => {
    const checked = orgRoles.has(ro.id);
    const raw     = parseInt(ro.color);
    const color   = raw ? '#' + raw.toString(16).padStart(6,'0') : '#5a5a68';
    return `<div class="role-item${checked ? ' checked' : ''}" onclick="toggleOrgRole('${ro.id}',this)">
      <div class="role-check"></div>
      <div class="role-dot" style="background:${color}"></div>
      <span class="role-name">${esc(ro.name)}</span>
    </div>`;
  }).join('');
}

function toggleOrgRole(id, el) {
  orgRoles.has(id) ? orgRoles.delete(id) : orgRoles.add(id);
  el.classList.toggle('checked');
}

async function saveConfig() {
  const guildId = document.getElementById('guild-select').value;
  if (!guildId) return;
  const pRoleId = document.getElementById('participant-select').value;
  try {
    const res = await fetch('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        guild_id: guildId,
        organizer_role_ids: [...orgRoles],
        participant_role_id: pRoleId || null,
      }),
    });
    res.ok ? toast('✓ Settings saved') : toast('Failed to save', true);
  } catch(e) { toast('Error: ' + e.message, true); }
}

// Init
loadOverview();
</script>
</body>
</html>
"""

# ─── Dashboard Route Handlers ────────────────────────────────────────────────

async def handle_index(request):
    return aio_web.Response(text=DASHBOARD_HTML, content_type="text/html")


async def handle_api_stats(request):
    conn      = db_connect()
    total     = conn.execute("SELECT COUNT(*) FROM registrations").fetchone()[0]
    active    = conn.execute("SELECT COUNT(*) FROM registrations WHERE status='active'").fetchone()[0]
    withdrawn = conn.execute("SELECT COUNT(*) FROM registrations WHERE status='withdrawn'").fetchone()[0]
    removed   = conn.execute("SELECT COUNT(*) FROM registrations WHERE status='removed'").fetchone()[0]
    u_teams   = conn.execute(
        "SELECT COUNT(DISTINCT team_name) FROM registrations WHERE team_name IS NOT NULL AND status='active'"
    ).fetchone()[0]
    top_teams = conn.execute("""
        SELECT team_name, COUNT(*) as cnt FROM registrations
        WHERE status='active' AND team_name IS NOT NULL
        GROUP BY team_name ORDER BY cnt DESC LIMIT 6
    """).fetchall()
    recent    = conn.execute("""
        SELECT ingame_name, registered_at FROM registrations ORDER BY registered_at DESC LIMIT 8
    """).fetchall()
    conn.close()

    return aio_web.json_response({
        "total":       total,
        "active":      active,
        "withdrawn":   withdrawn,
        "removed":     removed,
        "unique_teams": u_teams,
        "guilds":      len(bot.guilds),
        "top_teams":   [{"name": r["team_name"], "count": r["cnt"]} for r in top_teams],
        "recent":      [{"ingame_name": r["ingame_name"], "registered_at": r["registered_at"]} for r in recent],
    })


async def handle_api_registrations(request):
    conn = db_connect()
    rows = conn.execute("SELECT * FROM registrations ORDER BY registered_at DESC").fetchall()
    conn.close()
    return aio_web.json_response([dict(r) for r in rows])


async def handle_api_guilds(request):
    result = []
    for guild in bot.guilds:
        roles = [
            {"id": str(r.id), "name": r.name, "color": str(r.color.value)}
            for r in guild.roles
            if not r.is_default() and not r.managed
        ]
        result.append({"id": str(guild.id), "name": guild.name, "roles": roles})
    return aio_web.json_response(result)


async def handle_api_get_config(request):
    guild_id = request.rel_url.query.get("guild_id", "")
    if not guild_id:
        return aio_web.json_response({"error": "guild_id required"}, status=400)
    return aio_web.json_response(get_guild_config(guild_id))


async def handle_api_set_config(request):
    try:
        data          = await request.json()
        guild_id      = data.get("guild_id")
        org_ids_json  = json.dumps(data.get("organizer_role_ids", []))
        p_role_id     = data.get("participant_role_id")
        if not guild_id:
            return aio_web.json_response({"error": "guild_id required"}, status=400)
        conn = db_connect()
        conn.execute("""
            INSERT INTO guild_config (guild_id, organizer_role_ids, participant_role_id)
            VALUES (?, ?, ?)
            ON CONFLICT(guild_id) DO UPDATE SET
                organizer_role_ids  = excluded.organizer_role_ids,
                participant_role_id = excluded.participant_role_id
        """, (guild_id, org_ids_json, p_role_id))
        conn.commit()
        conn.close()
        log.info("Guild config updated for %s", guild_id)
        return aio_web.json_response({"ok": True})
    except Exception as e:
        log.error("Config save error: %s", e)
        return aio_web.json_response({"error": str(e)}, status=500)


async def start_dashboard():
    app = aio_web.Application()
    app.router.add_get("/",                    handle_index)
    app.router.add_get("/api/stats",           handle_api_stats)
    app.router.add_get("/api/registrations",   handle_api_registrations)
    app.router.add_get("/api/guilds",          handle_api_guilds)
    app.router.add_get("/api/config",          handle_api_get_config)
    app.router.add_post("/api/config",         handle_api_set_config)

    runner = aio_web.AppRunner(app)
    await runner.setup()
    site = aio_web.TCPSite(runner, DASHBOARD_HOST, DASHBOARD_PORT)
    await site.start()
    log.info("Dashboard running at http://%s:%s", DASHBOARD_HOST, DASHBOARD_PORT)

# ─── Events ──────────────────────────────────────────────────────────────────

@bot.event
async def on_ready():
    log.info("Logged in as %s (ID: %s)", bot.user, bot.user.id)
    init_db()
    asyncio.create_task(start_dashboard())
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
        return
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