"""One-off read-only diagnostic: is System/Purpose really 100% NULL across the
WHOLE ita.snowserver table (not just the matched-with-Excel subset), and do a
handful of other mapped columns generally agree between Excel and the DB for
matched servers (sanity spot-check)."""
import os
import pyodbc
from dotenv import load_dotenv

load_dotenv()

conn = pyodbc.connect(os.environ["DB_CONN"])
cur = conn.cursor()

cur.execute("SELECT COUNT(*) FROM ita.snowserver")
total = cur.fetchone()[0]

cols_to_check = ["System", "Purpose", "Primaryapplicationadmin", "OSversion", "Supportedby", "Hostadmin"]
print(f"Total rows in ita.snowserver: {total}")
for c in cols_to_check:
    cur.execute(f"SELECT COUNT([{c}]) FROM ita.snowserver WHERE [{c}] IS NOT NULL")
    non_null = cur.fetchone()[0]
    print(f"  {c}: {non_null} non-NULL out of {total}")

conn.close()
