import sqlite3
conn = sqlite3.connect(r'E:\CRT\cloudrtech_hr\db.sqlite3')
cur = conn.cursor()
cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
tables = cur.fetchall()
print('Tables:', tables)
for t in tables:
    tname = t[0]
    print(f'\n--- {tname} ---')
    cur.execute(f'PRAGMA table_info({tname})')
    cols = cur.fetchall()
    for c in cols:
        print(f'  {c}')
    cur.execute(f'SELECT COUNT(*) FROM {tname}')
    cnt = cur.fetchone()[0]
    print(f'  Row count: {cnt}')
    if cnt > 0:
        cur.execute(f'SELECT * FROM {tname} LIMIT 5')
        rows = cur.fetchall()
        for r in rows:
            print(f'  -> {r}')
conn.close()
