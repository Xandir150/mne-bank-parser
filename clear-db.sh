#!/bin/bash
# Clear izvod SQLite database (statements + transactions)
SSH_KEY="${SSH_KEY:-$HOME/.ssh/id_tools_eddsa}"
SSH_HOST="${SSH_HOST:-sshuser@10.252.1.47}"

ssh -i "$SSH_KEY" -o ConnectTimeout=10 -T "$SSH_HOST" \
  "docker exec -w /app mne-bank-parser-main-izvod-1 python3 -c \"import sqlite3; c=sqlite3.connect('/data/db/statements.db'); c.execute('DELETE FROM transactions'); c.execute('DELETE FROM statements'); c.commit(); print('Cleared:', c.execute('SELECT count(*) FROM statements').fetchone()[0], 'statements,', c.execute('SELECT count(*) FROM transactions').fetchone()[0], 'transactions'); c.close()\"" </dev/null
