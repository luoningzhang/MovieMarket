#!/usr/bin/env python3
"""
用法:
  python preview.py                        # 随机看 3 部
  python preview.py "Batman Begins"
  python preview.py "Batman Begins" "Inception" "The Dark Knight"
"""
import sqlite3, sys, random

DB = "moviemarket.db"

QUERY = """
SELECT
    m.title, m.year, m.release_date, m.studio,
    m.tmdb_id, m.imdb_id, m.overview, m.tagline,
    m.runtime, m.budget, m.status,
    m.original_language, m.popularity,
    GROUP_CONCAT(DISTINCT g.name)           AS genres,
    b.budget      AS bo_budget,
    b.domestic    AS bo_domestic,
    b.international,
    b.worldwide,
    b.opening_wknd,
    m.enriched
FROM movies m
LEFT JOIN movie_genres mg ON mg.movie_id = m.id
LEFT JOIN genres g        ON g.id = mg.genre_id
LEFT JOIN box_office b    ON b.movie_id = m.id
WHERE m.title LIKE ?
GROUP BY m.id
LIMIT 5
"""

def fmt_money(v):
    if v is None: return "—"
    if v >= 1_000_000_000: return f"${v/1e9:.2f}B"
    if v >= 1_000_000:     return f"${v/1e6:.1f}M"
    return f"${v:,}"

def fmt_min(v):
    if v is None: return "—"
    return f"{v // 60}h {v % 60}m"

def show(conn, title_pattern):
    cur = conn.cursor()
    rows = cur.execute(QUERY, (f"%{title_pattern}%",)).fetchall()
    if not rows:
        print(f'  [未找到] "{title_pattern}"\n')
        return

    for row in rows:
        movie_id = cur.execute(
            "SELECT id FROM movies WHERE title LIKE ?", (f"%{title_pattern}%",)
        ).fetchone()
        mid = movie_id[0] if movie_id else None

        # 评分
        ratings = {}
        if mid:
            for r in cur.execute(
                "SELECT source, score, vote_count FROM ratings WHERE movie_id=?", (mid,)
            ).fetchall():
                ratings[r[0]] = (r[1], r[2])

        # 演员/导演
        cast    = []
        directors = []
        writers   = []
        if mid:
            for p in cur.execute(
                "SELECT name, role, character FROM cast_crew WHERE movie_id=? ORDER BY sort_order",
                (mid,)
            ).fetchall():
                if p[1] == "actor":
                    char = f" ({p[2]})" if p[2] else ""
                    cast.append(p[0] + char)
                elif p[1] == "director":
                    directors.append(p[0])
                elif p[1] in ("screenplay", "writer", "story"):
                    writers.append(p[0])

        # 制作公司
        companies = []
        if mid:
            companies = [r[0] for r in cur.execute(
                "SELECT c.name FROM companies c JOIN movie_companies mc ON mc.company_id=c.id WHERE mc.movie_id=?",
                (mid,)
            ).fetchall()]

        # 分级
        cert = None
        if mid:
            r = cur.execute(
                "SELECT rating FROM certifications WHERE movie_id=? AND country='US'", (mid,)
            ).fetchone()
            cert = r[0] if r else None

        enriched_status = {0: "仅Excel原始数据", 1: "已爬TMDb", 2: "已爬TMDb+OMDb"}

        print("=" * 60)
        print(f"  {row['title']}  ({row['year']})")
        print("=" * 60)
        print(f"  上映日期    : {row['release_date'] or '—'}")
        print(f"  类型        : {row['genres'] or '—'}")
        print(f"  MPAA 分级   : {cert or '—'}")
        print(f"  时长        : {fmt_min(row['runtime'])}")
        print(f"  语言        : {row['original_language'] or '—'}")
        print(f"  状态        : {row['status'] or '—'}")
        print(f"  发行公司(Excel): {row['studio'] or '—'}")
        if companies:
            print(f"  制作公司(TMDb): {' | '.join(companies)}")
        print()
        print(f"  导演        : {' | '.join(directors) or '—'}")
        print(f"  编剧        : {' | '.join(writers) or '—'}")
        print(f"  主演        : {', '.join(cast[:8]) or '—'}")
        print()
        print(f"  预算        : {fmt_money(row['bo_budget'] or row['budget'])}")
        print(f"  美国票房    : {fmt_money(row['bo_domestic'])}")
        print(f"  海外票房    : {fmt_money(row['international'])}")
        print(f"  全球票房    : {fmt_money(row['worldwide'])}")
        print(f"  首周末票房  : {fmt_money(row['opening_wknd'])}")
        print()
        for src in ["IMDb", "Rotten Tomatoes", "Metacritic", "TMDb"]:
            if src in ratings:
                score, votes = ratings[src]
                v = f"  ({votes:,} votes)" if votes else ""
                print(f"  {src:<16}: {score}{v}")
        print()
        if row['overview']:
            # 简介折行
            words = row['overview'].split()
            line, lines = [], []
            for w in words:
                line.append(w)
                if len(" ".join(line)) > 55:
                    lines.append(" ".join(line))
                    line = []
            if line: lines.append(" ".join(line))
            print(f"  简介        : {lines[0]}")
            for l in lines[1:]:
                print(f"               {l}")
        if row['tagline']:
            print(f"  Tagline     : {row['tagline']}")
        print(f"\n  TMDb ID     : {row['tmdb_id'] or '—'}")
        print(f"  IMDb ID     : {row['imdb_id'] or '—'}")
        print(f"  数据状态    : {enriched_status.get(row['enriched'], row['enriched'])}")
        print()


def main():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row

    if len(sys.argv) > 1:
        titles = sys.argv[1:]
    else:
        # 随机挑 3 部已enriched的
        rows = conn.execute(
            "SELECT title FROM movies WHERE enriched > 0 ORDER BY RANDOM() LIMIT 3"
        ).fetchall()
        if not rows:
            rows = conn.execute(
                "SELECT title FROM movies ORDER BY RANDOM() LIMIT 3"
            ).fetchall()
        titles = [r[0] for r in rows]
        print(f"随机抽取: {titles}\n")

    for t in titles:
        show(conn, t)

    conn.close()

if __name__ == "__main__":
    main()
