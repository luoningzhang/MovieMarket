#!/usr/bin/env python3
"""
MovieMarket Database Builder
============================================================
Pipeline:
  1. import_excel  – parse 05-27美国电影.xlsx → SQLite
  2. enrich_tmdb   – TMDb API: box office revenue, genres, cast/crew,
                     certifications (MPAA rating), production companies
  3. enrich_omdb   – OMDb API: IMDb / Rotten Tomatoes / Metacritic scores,
                     domestic box office

API keys (free tier):
  TMDb : https://www.themoviedb.org/settings/api
  OMDb : https://www.omdbapi.com/apikey.aspx

Usage:
  # Full pipeline (import + both APIs)
  python build_db.py

  # Step by step
  python build_db.py --import-only
  python build_db.py --tmdb-only
  python build_db.py --omdb-only

  # Check progress
  python build_db.py --stats
"""

import argparse
import ast
import configparser
import os
import sqlite3
import time
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import requests

# ── Config ──────────────────────────────────────────────────────────────────

_cfg = configparser.ConfigParser()
_cfg.read(Path(__file__).with_name("config.ini"))

EXCEL_FILE   = _cfg.get("database", "excel_file", fallback="05-27美国电影.xlsx")
DB_FILE      = _cfg.get("database", "db_file",    fallback="moviemarket.db")
TMDB_API_KEY = _cfg.get("api", "tmdb_api_key",    fallback="")
OMDB_API_KEY = _cfg.get("api", "omdb_api_key",    fallback="")

TMDB_BASE = "https://api.themoviedb.org/3"
OMDB_BASE = "https://www.omdbapi.com"

# TMDb free tier: 50 req / 10s  →  keep at ~0.26s between calls
TMDB_SLEEP = 0.26
OMDB_SLEEP = 0.12


# ── Database Schema ──────────────────────────────────────────────────────────

SCHEMA = """
-- Core movie record (one row per film)
CREATE TABLE IF NOT EXISTS movies (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    title             TEXT    NOT NULL,
    year              INTEGER,
    release_date      TEXT,           -- YYYY-MM-DD
    studio            TEXT,           -- from Excel (raw distribution info)
    tmdb_id           INTEGER UNIQUE,
    imdb_id           TEXT    UNIQUE,
    overview          TEXT,           -- plot summary
    tagline           TEXT,
    runtime           INTEGER,        -- minutes
    budget            INTEGER,        -- USD
    status            TEXT,           -- "Released", "Post Production", …
    original_language TEXT,
    original_title    TEXT,
    popularity        REAL,           -- TMDb popularity score
    enriched          INTEGER DEFAULT 0
                                      -- 0 = raw Excel only
                                      -- 1 = TMDb done
                                      -- 2 = OMDb done
);

-- Lookup table for genres
CREATE TABLE IF NOT EXISTS genres (
    id   INTEGER PRIMARY KEY,         -- TMDb genre id
    name TEXT    UNIQUE NOT NULL
);

-- Movie ↔ Genre (many-to-many)
CREATE TABLE IF NOT EXISTS movie_genres (
    movie_id INTEGER NOT NULL REFERENCES movies(id) ON DELETE CASCADE,
    genre_id INTEGER NOT NULL REFERENCES genres(id) ON DELETE CASCADE,
    PRIMARY KEY (movie_id, genre_id)
);

-- Box office figures (all USD)
CREATE TABLE IF NOT EXISTS box_office (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    movie_id      INTEGER UNIQUE NOT NULL REFERENCES movies(id) ON DELETE CASCADE,
    budget        INTEGER,        -- production budget (TMDb)
    domestic      INTEGER,        -- US/Canada gross (OMDb)
    international INTEGER,        -- derived: worldwide - domestic
    worldwide     INTEGER,        -- total global gross (TMDb revenue field)
    opening_wknd  INTEGER         -- US opening weekend (OMDb where available)
);

-- Ratings / review aggregators
CREATE TABLE IF NOT EXISTS ratings (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    movie_id   INTEGER NOT NULL REFERENCES movies(id) ON DELETE CASCADE,
    source     TEXT    NOT NULL,   -- 'TMDb', 'IMDb', 'Rotten Tomatoes', 'Metacritic'
    score      TEXT,               -- kept as text: "8.2", "94%", "72/100"
    vote_count INTEGER,
    UNIQUE (movie_id, source)
);

-- MPAA / content certifications (one per country, highest-priority rating)
CREATE TABLE IF NOT EXISTS certifications (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    movie_id INTEGER NOT NULL REFERENCES movies(id) ON DELETE CASCADE,
    country  TEXT    NOT NULL,     -- ISO 3166-1, e.g. "US"
    rating   TEXT,                 -- G, PG, PG-13, R, NC-17, NR …
    UNIQUE (movie_id, country)
);

-- Per-country release schedule (all release types from TMDb + inferred re-releases)
-- release_type codes: 1=Premiere 2=LimitedTheatrical 3=Theatrical
--                     4=Digital  5=Physical           6=TV
--                     7=Re-release (inferred from Excel duplicate rows)
CREATE TABLE IF NOT EXISTS release_schedule (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    movie_id      INTEGER NOT NULL REFERENCES movies(id) ON DELETE CASCADE,
    country       TEXT    NOT NULL,
    release_date  TEXT,            -- YYYY-MM-DD
    release_type  INTEGER,
    certification TEXT,            -- local age rating for this release
    note          TEXT,            -- e.g. "3D re-release", "Director's Cut"
    UNIQUE (movie_id, country, release_type)
);
CREATE INDEX IF NOT EXISTS idx_schedule_movie   ON release_schedule(movie_id);
CREATE INDEX IF NOT EXISTS idx_schedule_country ON release_schedule(country);

-- Cast & crew
CREATE TABLE IF NOT EXISTS cast_crew (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    movie_id   INTEGER NOT NULL REFERENCES movies(id) ON DELETE CASCADE,
    name       TEXT    NOT NULL,
    role       TEXT,               -- 'actor', 'director', 'writer', 'producer', …
    character  TEXT,               -- character name (actors only)
    sort_order INTEGER             -- billing order
);

-- Production / distribution companies
CREATE TABLE IF NOT EXISTS companies (
    id   INTEGER PRIMARY KEY,      -- TMDb company id
    name TEXT    UNIQUE NOT NULL
);

CREATE TABLE IF NOT EXISTS movie_companies (
    movie_id   INTEGER NOT NULL REFERENCES movies(id)   ON DELETE CASCADE,
    company_id INTEGER NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    PRIMARY KEY (movie_id, company_id)
);

-- Useful indexes
CREATE INDEX IF NOT EXISTS idx_movies_year    ON movies(year);
CREATE INDEX IF NOT EXISTS idx_movies_tmdb    ON movies(tmdb_id);
CREATE INDEX IF NOT EXISTS idx_movies_imdb    ON movies(imdb_id);
CREATE INDEX IF NOT EXISTS idx_cast_movie     ON cast_crew(movie_id);
CREATE INDEX IF NOT EXISTS idx_ratings_movie  ON ratings(movie_id);
"""


# ── Init DB ──────────────────────────────────────────────────────────────────

def init_db(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)
    conn.execute("PRAGMA journal_mode=WAL")   # safe concurrent writes
    conn.execute("PRAGMA foreign_keys=ON")
    conn.commit()
    print(f"[DB] Ready → {db_path}")
    return conn


# ── Step 1: Import Excel ─────────────────────────────────────────────────────

def _excel_serial_to_iso(serial) -> str | None:
    """Convert Excel date serial number (days since 1900-01-00) to ISO string."""
    try:
        n = int(float(serial))
        if n <= 0:
            return None
        # Excel epoch quirk: serial 1 = 1900-01-01, but has the Lotus leap-year bug
        dt = date(1899, 12, 30) + timedelta(days=n)
        return dt.isoformat()
    except (ValueError, TypeError, OverflowError):
        return None


def import_excel(conn: sqlite3.Connection, excel_path: str) -> int:
    """
    Parse 05-27美国电影.xlsx and insert raw records into `movies` + `cast_crew`.
    Safe to re-run — uses INSERT OR IGNORE on title+year.
    """
    print(f"[Import] Reading {excel_path} …")
    df = pd.read_excel(excel_path, sheet_name="Sheet1", header=0, dtype=str)

    # Normalise column names regardless of original order
    df.columns = [
        "year", "date_serial", "title", "studio",
        "director", "writer", "cast_raw", "genre", "notes", "raw_columns"
    ]

    cur = conn.cursor()
    inserted = skipped = 0

    for _, row in df.iterrows():
        title = (row.get("title") or "").strip()
        if not title or title.lower() == "title":
            continue  # skip blank / repeated header rows

        year_raw = row.get("year", "")
        try:
            year = int(float(year_raw)) if year_raw and year_raw != "nan" else None
        except ValueError:
            year = None

        release_date = _excel_serial_to_iso(row.get("date_serial"))
        studio       = str(row["studio"]).strip()     if pd.notna(row.get("studio"))      else None
        cast_raw     = str(row["cast_raw"]).strip()   if pd.notna(row.get("cast_raw"))    else None
        studio       = studio   or None
        cast_raw     = cast_raw or None

        # Try to get a better release date from raw_columns dict
        if not release_date:
            raw = str(row["raw_columns"]).strip() if pd.notna(row.get("raw_columns")) else ""
            if raw and raw != "nan":
                try:
                    d = ast.literal_eval(raw)
                    month = d.get("month", "")
                    day   = d.get("day", "")
                    if month and day and year:
                        try:
                            dt = pd.to_datetime(f"{month} {day} {year}", errors="coerce")
                            if pd.notna(dt):
                                release_date = dt.date().isoformat()
                        except Exception:
                            pass
                except Exception:
                    pass

        try:
            cur.execute(
                """INSERT OR IGNORE INTO movies (title, year, release_date, studio)
                   VALUES (?, ?, ?, ?)""",
                (title, year, release_date, studio)
            )
            if cur.rowcount:
                movie_id = cur.lastrowid
                _insert_cast_raw(cur, movie_id, cast_raw)
                inserted += 1
            else:
                skipped += 1
        except sqlite3.Error as e:
            print(f"  [WARN] '{title}' ({year}): {e}")

    conn.commit()
    print(f"[Import] Done — {inserted} inserted, {skipped} skipped (already exist)")
    return inserted


def _insert_cast_raw(cur: sqlite3.Cursor, movie_id: int, cast_raw: str | None):
    if not cast_raw or cast_raw == "nan":
        return
    for i, name in enumerate(cast_raw.split("|")):
        name = name.strip()
        if name:
            cur.execute(
                """INSERT OR IGNORE INTO cast_crew (movie_id, name, role, sort_order)
                   VALUES (?, ?, 'unknown', ?)""",
                (movie_id, name, i)
            )


# ── Step 2: TMDb Enrichment ──────────────────────────────────────────────────

def _tmdb_get(path: str, params: dict) -> dict | None:
    params["api_key"] = TMDB_API_KEY
    try:
        r = requests.get(f"{TMDB_BASE}{path}", params=params, timeout=10)
        r.raise_for_status()
        return r.json()
    except requests.HTTPError as e:
        if e.response is not None and e.response.status_code == 429:
            print("  [TMDb] Rate limited — sleeping 10s …")
            time.sleep(10)
        return None
    except Exception as e:
        print(f"  [TMDb] Request error: {e}")
        return None


def _tmdb_search(title: str, year: int | None) -> dict | None:
    """Return the best-matching TMDb movie result, or None."""
    params = {"query": title, "language": "en-US", "include_adult": "false"}
    if year:
        params["year"] = year
    data = _tmdb_get("/search/movie", params)
    if data and data.get("results"):
        return data["results"][0]
    # Retry without year constraint if no results
    if year:
        params.pop("year")
        data = _tmdb_get("/search/movie", params)
        if data and data.get("results"):
            return data["results"][0]
    return None


def _tmdb_details(tmdb_id: int) -> dict | None:
    return _tmdb_get(
        f"/movie/{tmdb_id}",
        {"language": "en-US", "append_to_response": "credits,release_dates"}
    )


def _apply_tmdb_data(conn: sqlite3.Connection, cur: sqlite3.Cursor,
                     movie_id: int, d: dict):
    """Write all TMDb fields into the database."""

    # ── movies table ──
    # imdb_id has a UNIQUE constraint; if another row already claimed this
    # imdb_id (duplicate titles in the Excel), skip setting it to avoid crash.
    imdb_id = d.get("imdb_id")
    if imdb_id:
        clash = cur.execute(
            "SELECT id FROM movies WHERE imdb_id=? AND id!=?", (imdb_id, movie_id)
        ).fetchone()
        if clash:
            imdb_id = None   # don't overwrite the other row's claim

    try:
        cur.execute("""
            UPDATE movies SET
                tmdb_id           = ?,
                imdb_id           = COALESCE(imdb_id, ?),
                overview          = ?,
                tagline           = ?,
                runtime           = ?,
                budget            = CASE WHEN ? > 0 THEN ? ELSE budget END,
                status            = ?,
                original_language = ?,
                original_title    = ?,
                popularity        = ?,
                release_date      = COALESCE(NULLIF(release_date, ''), ?)
            WHERE id = ?
        """, (
            d.get("id"),
            imdb_id,
            d.get("overview"),
            d.get("tagline"),
            d.get("runtime"),
            d.get("budget") or 0, d.get("budget"),
            d.get("status"),
            d.get("original_language"),
            d.get("original_title"),
            d.get("popularity"),
            d.get("release_date"),
            movie_id,
        ))
    except sqlite3.IntegrityError:
        # Rare race: retry without imdb_id
        cur.execute("""
            UPDATE movies SET
                tmdb_id=?, overview=?, tagline=?, runtime=?,
                budget=CASE WHEN ? > 0 THEN ? ELSE budget END,
                status=?, original_language=?, original_title=?,
                popularity=?, release_date=COALESCE(NULLIF(release_date,''),?)
            WHERE id=?
        """, (
            d.get("id"), d.get("overview"), d.get("tagline"), d.get("runtime"),
            d.get("budget") or 0, d.get("budget"),
            d.get("status"), d.get("original_language"), d.get("original_title"),
            d.get("popularity"), d.get("release_date"), movie_id,
        ))

    # ── box_office: TMDb revenue = worldwide gross ──
    revenue = d.get("revenue") or 0
    budget  = d.get("budget")  or 0
    if revenue or budget:
        cur.execute("""
            INSERT INTO box_office (movie_id, worldwide, budget)
                VALUES (?, ?, ?)
            ON CONFLICT(movie_id) DO UPDATE SET
                worldwide = CASE WHEN excluded.worldwide > 0
                            THEN excluded.worldwide ELSE worldwide END,
                budget    = CASE WHEN excluded.budget > 0
                            THEN excluded.budget    ELSE budget    END
        """, (movie_id, revenue or None, budget or None))

    # ── genres ──
    for g in d.get("genres", []):
        cur.execute("INSERT OR IGNORE INTO genres (id, name) VALUES (?,?)",
                    (g["id"], g["name"]))
        cur.execute("INSERT OR IGNORE INTO movie_genres VALUES (?,?)",
                    (movie_id, g["id"]))

    # ── TMDb vote score ──
    va = d.get("vote_average")
    vc = d.get("vote_count")
    if va:
        cur.execute("""
            INSERT INTO ratings (movie_id, source, score, vote_count)
                VALUES (?, 'TMDb', ?, ?)
            ON CONFLICT(movie_id, source) DO UPDATE SET
                score=excluded.score, vote_count=excluded.vote_count
        """, (movie_id, str(round(va, 1)), vc))

    # ── Release schedule (all countries) + certifications ──
    for entry in d.get("release_dates", {}).get("results", []):
        country = entry.get("iso_3166_1", "").strip()
        if not country:
            continue
        best_cert = None  # pick first non-empty certification per country
        for rd in entry.get("release_dates", []):
            rdate = (rd.get("release_date") or "")[:10] or None  # trim time part
            rtype = rd.get("type")                                # 1-6
            cert  = (rd.get("certification") or "").strip() or None
            if cert and best_cert is None:
                best_cert = cert
            cur.execute("""
                INSERT INTO release_schedule
                    (movie_id, country, release_date, release_type, certification)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(movie_id, country, release_type) DO UPDATE SET
                    release_date  = COALESCE(excluded.release_date,  release_date),
                    certification = COALESCE(excluded.certification, certification)
            """, (movie_id, country, rdate, rtype, cert))
        # Store best certification in the certifications lookup table
        if best_cert:
            cur.execute("""
                INSERT INTO certifications (movie_id, country, rating)
                    VALUES (?, ?, ?)
                ON CONFLICT(movie_id, country) DO UPDATE SET
                    rating=excluded.rating
            """, (movie_id, country, best_cert))

    # ── Credits ──
    # Delete placeholder rows inserted during Excel import
    cur.execute("DELETE FROM cast_crew WHERE movie_id=? AND role='unknown'",
                (movie_id,))

    credits = d.get("credits", {})
    for p in credits.get("cast", [])[:25]:   # top-billed 25 actors
        cur.execute("""
            INSERT OR IGNORE INTO cast_crew
                (movie_id, name, role, character, sort_order)
            VALUES (?, ?, 'actor', ?, ?)
        """, (movie_id, p["name"], p.get("character"), p.get("order")))

    CREW_JOBS = {"Director", "Screenplay", "Writer", "Story",
                 "Executive Producer", "Producer"}
    for p in credits.get("crew", []):
        if p.get("job") in CREW_JOBS:
            cur.execute("""
                INSERT OR IGNORE INTO cast_crew (movie_id, name, role)
                VALUES (?, ?, ?)
            """, (movie_id, p["name"], p["job"].lower()))

    # ── Production companies ──
    for co in d.get("production_companies", []):
        cur.execute("INSERT OR IGNORE INTO companies (id, name) VALUES (?,?)",
                    (co["id"], co["name"]))
        cur.execute("INSERT OR IGNORE INTO movie_companies VALUES (?,?)",
                    (movie_id, co["id"]))


def enrich_tmdb(conn: sqlite3.Connection, limit: int | None = None):
    """Enrich all unenriched movies with TMDb data."""
    if not TMDB_API_KEY:
        print("[TMDb] TMDB_API_KEY not set — skipping.")
        print("       Get a free key at https://www.themoviedb.org/settings/api")
        return

    cur = conn.cursor()
    sql = "SELECT id, title, year FROM movies WHERE enriched = 0 ORDER BY year, title"
    if limit:
        sql += f" LIMIT {limit}"
    rows = cur.execute(sql).fetchall()
    total = len(rows)
    print(f"[TMDb] Enriching {total} movies …")

    for i, row in enumerate(rows, 1):
        movie_id, title, year = row["id"], row["title"], row["year"]

        match = _tmdb_search(title, year)
        if not match:
            cur.execute("UPDATE movies SET enriched=1 WHERE id=?", (movie_id,))
            time.sleep(TMDB_SLEEP)
            continue

        tmdb_id = match["id"]

        # ── Re-release detection ──────────────────────────────────────────
        # If another movies row already owns this tmdb_id, the current Excel
        # row is a re-release (or duplicate).  Merge it: record the release
        # date in release_schedule, then delete the redundant movies row.
        original = cur.execute(
            "SELECT id FROM movies WHERE tmdb_id=? AND id!=?", (tmdb_id, movie_id)
        ).fetchone()

        if original:
            original_id = original["id"]
            # Get the release date we stored during Excel import
            cur_row = cur.execute(
                "SELECT release_date, year FROM movies WHERE id=?", (movie_id,)
            ).fetchone()
            rdate = cur_row["release_date"]
            ryear = cur_row["year"]
            # Fall back to Jan 1 of the year if we only have a year
            if not rdate and ryear:
                rdate = f"{ryear}-01-01"
            if rdate:
                cur.execute("""
                    INSERT OR IGNORE INTO release_schedule
                        (movie_id, country, release_date, release_type, note)
                    VALUES (?, 'US', ?, 7, ?)
                """, (original_id, rdate, f"Re-release ({title})"))
                print(f"  [Re-release] '{title}' {year} → merged into movie_id={original_id}")
            # Delete the redundant row (CASCADE removes cast_crew etc.)
            cur.execute("DELETE FROM movies WHERE id=?", (movie_id,))
            conn.commit()
            time.sleep(TMDB_SLEEP)
            continue
        # ─────────────────────────────────────────────────────────────────

        details = _tmdb_details(tmdb_id)
        if details:
            _apply_tmdb_data(conn, cur, movie_id, details)

        cur.execute("UPDATE movies SET enriched=1 WHERE id=?", (movie_id,))
        conn.commit()

        if i % 50 == 0 or i == total:
            pct = i / total * 100
            print(f"  {i:>5}/{total}  ({pct:.1f}%)")

        time.sleep(TMDB_SLEEP)

    print("[TMDb] Done.")


# ── Step 3: OMDb Enrichment ──────────────────────────────────────────────────

def enrich_omdb(conn: sqlite3.Connection, limit: int | None = None):
    """
    Pull IMDb, Rotten Tomatoes, Metacritic scores + US box office via OMDb.
    Requires movies to have an imdb_id (set by TMDb step).
    """
    if not OMDB_API_KEY:
        print("[OMDb] OMDB_API_KEY not set — skipping.")
        print("       Get a free key (1 000 req/day) at https://www.omdbapi.com/apikey.aspx")
        return

    cur = conn.cursor()
    sql = """SELECT id, imdb_id FROM movies
             WHERE enriched = 1 AND imdb_id IS NOT NULL
             ORDER BY year"""
    if limit:
        sql += f" LIMIT {limit}"
    rows = cur.execute(sql).fetchall()
    total = len(rows)
    print(f"[OMDb] Fetching ratings for {total} movies …")

    for i, row in enumerate(rows, 1):
        movie_id, imdb_id = row["id"], row["imdb_id"]
        try:
            r = requests.get(
                OMDB_BASE,
                params={"apikey": OMDB_API_KEY, "i": imdb_id, "tomatoes": "true"},
                timeout=10,
            )
            r.raise_for_status()
            data = r.json()
            if data.get("Response") == "True":
                _apply_omdb_data(cur, movie_id, data)
                conn.commit()
        except Exception as e:
            print(f"  [OMDb] {imdb_id}: {e}")

        if i % 100 == 0 or i == total:
            print(f"  {i:>5}/{total}  ({i/total*100:.1f}%)")

        time.sleep(OMDB_SLEEP)

    # Mark as fully enriched
    conn.execute(
        "UPDATE movies SET enriched=2 WHERE enriched=1 AND imdb_id IS NOT NULL"
    )
    conn.commit()
    print("[OMDb] Done.")


_OMDB_SOURCE_MAP = {
    "Internet Movie Database": "IMDb",
    "Rotten Tomatoes":         "Rotten Tomatoes",
    "Metacritic":              "Metacritic",
}


def _apply_omdb_data(cur: sqlite3.Cursor, movie_id: int, d: dict):
    # Ratings list
    for entry in d.get("Ratings", []):
        src   = _OMDB_SOURCE_MAP.get(entry["Source"], entry["Source"])
        score = entry["Value"]
        vote_count = None
        # Parse vote count for IMDb e.g. "8.2/10"
        if src == "IMDb":
            imdb_votes_raw = d.get("imdbVotes", "").replace(",", "")
            try:
                vote_count = int(imdb_votes_raw)
            except ValueError:
                pass
        cur.execute("""
            INSERT INTO ratings (movie_id, source, score, vote_count)
                VALUES (?, ?, ?, ?)
            ON CONFLICT(movie_id, source) DO UPDATE SET
                score=excluded.score,
                vote_count=COALESCE(excluded.vote_count, vote_count)
        """, (movie_id, src, score, vote_count))

    # US domestic box office
    bo_raw = d.get("BoxOffice", "N/A").replace("$", "").replace(",", "").strip()
    if bo_raw and bo_raw != "N/A":
        try:
            domestic = int(bo_raw)
            cur.execute("""
                INSERT INTO box_office (movie_id, domestic) VALUES (?, ?)
                ON CONFLICT(movie_id) DO UPDATE SET
                    domestic = excluded.domestic,
                    international = CASE
                        WHEN worldwide IS NOT NULL AND excluded.domestic IS NOT NULL
                        THEN worldwide - excluded.domestic
                        ELSE international END
            """, (movie_id, domestic))
        except ValueError:
            pass

    # Opening weekend
    ow_raw = d.get("Opening", "N/A").replace("$", "").replace(",", "").strip()
    if ow_raw and ow_raw != "N/A":
        try:
            cur.execute("""
                INSERT INTO box_office (movie_id, opening_wknd) VALUES (?, ?)
                ON CONFLICT(movie_id) DO UPDATE SET opening_wknd=excluded.opening_wknd
            """, (movie_id, int(ow_raw)))
        except ValueError:
            pass


# ── Stats & Reporting ────────────────────────────────────────────────────────

def print_stats(conn: sqlite3.Connection):
    cur = conn.cursor()
    tables = [
        "movies", "genres", "movie_genres",
        "box_office", "ratings", "certifications",
        "cast_crew", "companies", "movie_companies",
    ]
    print("\n── Database Statistics ─────────────────────────────")
    for t in tables:
        n = cur.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
        print(f"  {t:<25} {n:>8,} rows")

    total    = cur.execute("SELECT COUNT(*) FROM movies").fetchone()[0]
    e1       = cur.execute("SELECT COUNT(*) FROM movies WHERE enriched >= 1").fetchone()[0]
    e2       = cur.execute("SELECT COUNT(*) FROM movies WHERE enriched >= 2").fetchone()[0]
    with_rev = cur.execute("SELECT COUNT(*) FROM movies WHERE tmdb_id IS NOT NULL").fetchone()[0]
    print(f"\n  Total movies        : {total:>8,}")
    print(f"  TMDb enriched       : {e1:>8,}")
    print(f"  OMDb enriched       : {e2:>8,}")
    print(f"  Matched to TMDb     : {with_rev:>8,}")

    # Top genres
    print("\n  Top 10 genres:")
    rows = cur.execute("""
        SELECT g.name, COUNT(*) AS cnt
        FROM movie_genres mg JOIN genres g ON mg.genre_id=g.id
        GROUP BY g.id ORDER BY cnt DESC LIMIT 10
    """).fetchall()
    for r in rows:
        print(f"    {r['name']:<20} {r['cnt']:>6,}")

    print("────────────────────────────────────────────────────\n")


# ── CLI Entry Point ──────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description="MovieMarket — build & enrich movie database",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    ap.add_argument("--excel",        default=EXCEL_FILE,
                    help=f"Excel input file (default: {EXCEL_FILE})")
    ap.add_argument("--db",           default=DB_FILE,
                    help=f"SQLite output file (default: {DB_FILE})")
    ap.add_argument("--import-only",  action="store_true",
                    help="Only import Excel, skip API enrichment")
    ap.add_argument("--tmdb-only",    action="store_true",
                    help="Only run TMDb enrichment step")
    ap.add_argument("--omdb-only",    action="store_true",
                    help="Only run OMDb enrichment step")
    ap.add_argument("--stats",        action="store_true",
                    help="Print database statistics and exit")
    ap.add_argument("--limit",        type=int, default=None,
                    help="Process only N movies (for testing)")
    args = ap.parse_args()

    conn = init_db(args.db)

    if args.stats:
        print_stats(conn)
    elif args.import_only:
        import_excel(conn, args.excel)
        print_stats(conn)
    elif args.tmdb_only:
        enrich_tmdb(conn, limit=args.limit)
        print_stats(conn)
    elif args.omdb_only:
        enrich_omdb(conn, limit=args.limit)
        print_stats(conn)
    else:
        # Full pipeline
        import_excel(conn, args.excel)
        enrich_tmdb(conn, limit=args.limit)
        enrich_omdb(conn, limit=args.limit)
        print_stats(conn)

    conn.close()


if __name__ == "__main__":
    main()
