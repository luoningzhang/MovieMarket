#!/usr/bin/env python3
"""
筛选全球上映电影
用法:
  python global_releases.py               # 默认：上映国家数 >= 10
  python global_releases.py --min 20      # 至少 20 个国家
  python global_releases.py --min 5 --year 2019        # 指定年份
  python global_releases.py --min 10 --type theatrical # 只看院线上映
  python global_releases.py --top 50      # 上映国家最多的前 50 部
  python global_releases.py --export global_movies.csv # 导出 CSV
"""

import argparse
import csv
import sqlite3

DB = "moviemarket.db"

# release_type 3 = 院线, 1 = 首映, 2 = 限定院线, 7 = 重映
THEATRICAL_TYPES = (1, 2, 3, 7)

QUERY_ALL = """
SELECT
    m.id,
    m.title,
    m.year,
    m.release_date,
    m.original_language,
    m.runtime,
    m.status,
    GROUP_CONCAT(DISTINCT g.name ORDER BY g.name)  AS genres,
    b.worldwide,
    b.domestic,
    COUNT(DISTINCT rs.country)                     AS country_count,
    GROUP_CONCAT(DISTINCT rs.country ORDER BY rs.release_date) AS countries,
    r_imdb.score     AS imdb_score,
    r_rt.score       AS rt_score,
    r_mc.score       AS mc_score,
    r_tmdb.score     AS tmdb_score,
    cert_us.rating   AS mpaa_rating
FROM movies m
LEFT JOIN release_schedule rs  ON rs.movie_id = m.id {type_filter}
LEFT JOIN movie_genres mg      ON mg.movie_id = m.id
LEFT JOIN genres g             ON g.id = mg.genre_id
LEFT JOIN box_office b         ON b.movie_id = m.id
LEFT JOIN ratings r_imdb       ON r_imdb.movie_id = m.id AND r_imdb.source = 'IMDb'
LEFT JOIN ratings r_rt         ON r_rt.movie_id   = m.id AND r_rt.source   = 'Rotten Tomatoes'
LEFT JOIN ratings r_mc         ON r_mc.movie_id   = m.id AND r_mc.source   = 'Metacritic'
LEFT JOIN ratings r_tmdb       ON r_tmdb.movie_id = m.id AND r_tmdb.source = 'TMDb'
LEFT JOIN certifications cert_us ON cert_us.movie_id = m.id AND cert_us.country = 'US'
{year_filter}
GROUP BY m.id
HAVING country_count >= {min_countries}
ORDER BY country_count DESC, b.worldwide DESC NULLS LAST
{limit_clause}
"""


def fmt_money(v):
    if v is None:
        return ""
    if v >= 1_000_000_000:
        return f"${v/1e9:.2f}B"
    if v >= 1_000_000:
        return f"${v/1e6:.1f}M"
    return f"${v:,}"


def run(args):
    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row

    # 类型过滤
    if args.type == "theatrical":
        type_filter = f"AND rs.release_type IN {THEATRICAL_TYPES}"
    else:
        type_filter = ""

    # 年份过滤
    year_filter = f"WHERE m.year = {args.year}" if args.year else ""

    limit_clause = f"LIMIT {args.top}" if args.top else ""

    sql = QUERY_ALL.format(
        type_filter=type_filter,
        year_filter=year_filter,
        min_countries=args.min,
        limit_clause=limit_clause,
    )

    rows = conn.execute(sql).fetchall()

    if not rows:
        print("没有符合条件的电影。")
        print(f"（数据库里可能还没爬完，先跑 python build_db.py --tmdb-only）")
        return

    print(f"\n共找到 {len(rows)} 部全球上映电影（上映国家 >= {args.min}）\n")
    print(f"{'#':<4} {'片名':<40} {'年':<5} {'国家数':>5} {'全球票房':>10} "
          f"{'IMDb':>6} {'烂番茄':>6} {'类型'}")
    print("─" * 100)

    for i, r in enumerate(rows, 1):
        countries_preview = r["countries"] or ""
        if len(countries_preview) > 30:
            countries_preview = countries_preview[:30] + "…"

        print(
            f"{i:<4} {(r['title'] or '')[:38]:<40} {r['year'] or '':<5} "
            f"{r['country_count']:>5}  "
            f"{fmt_money(r['worldwide']):>10}  "
            f"{r['imdb_score'] or '—':>6}  "
            f"{r['rt_score'] or '—':>6}  "
            f"{(r['genres'] or '')[:25]}"
        )

    # 简单统计
    total_countries = set()
    for r in rows:
        if r["countries"]:
            total_countries.update(r["countries"].split(","))
    print(f"\n覆盖国家/地区共 {len(total_countries)} 个")

    # 导出 CSV
    if args.export:
        fields = [
            "id", "title", "year", "release_date", "original_language",
            "runtime", "status", "genres", "worldwide", "domestic",
            "country_count", "countries",
            "imdb_score", "rt_score", "mc_score", "tmdb_score", "mpaa_rating",
        ]
        with open(args.export, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=fields)
            writer.writeheader()
            for r in rows:
                writer.writerow({k: r[k] for k in fields})
        print(f"\n已导出 → {args.export}  ({len(rows)} 行)")

    conn.close()


def main():
    ap = argparse.ArgumentParser(description="筛选全球上映电影")
    ap.add_argument("--min",    type=int, default=10,
                    help="最少上映国家数（默认 10）")
    ap.add_argument("--top",    type=int, default=None,
                    help="只显示前 N 部")
    ap.add_argument("--year",   type=int, default=None,
                    help="限定某一年，例如 --year 2019")
    ap.add_argument("--type",   choices=["all", "theatrical"], default="all",
                    help="all=所有类型  theatrical=仅院线（默认 all）")
    ap.add_argument("--export", default=None,
                    help="导出结果为 CSV，例如 --export result.csv")
    ap.add_argument("--db",     default=DB,
                    help=f"数据库路径（默认 {DB}）")
    args = ap.parse_args()
    run(args)


if __name__ == "__main__":
    main()
