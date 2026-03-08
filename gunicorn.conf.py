# Configuració de Gunicorn per a l'InDret Corrector d'articles
# Referència: https://gunicorn.org/reference/settings/
#
# Arrencada: gunicorn -c gunicorn.conf.py app:app

import multiprocessing

# ─── Workers ────────────────────────────────────────────────
# 2 workers és suficient per a ús editorial de baixa concurrència.
# Fórmula general: (CPU * 2) + 1, però 2 és el límit pràctic per RAM de VPS.
workers = 2

# ─── Xarxa ──────────────────────────────────────────────────
# Port 8090: CloudPanel fa de proxy nginx cap a aquest port.
bind = "127.0.0.1:8090"

# ─── Timeouts ───────────────────────────────────────────────
# CRÍTIC: el corrector pot tardar 10-30 s en documents complexos.
# El timeout per defecte de gunicorn (30 s) mataria el worker a mig processament.
# 120 s és segur. L'nginx hauria de tenir proxy_read_timeout 130s (lleugerament superior).
timeout = 120
graceful_timeout = 30
keepalive = 2

# ─── Logs ───────────────────────────────────────────────────
# stdout/stderr: systemd/journald els recull automàticament.
accesslog = "-"
errorlog = "-"

# ─── Estabilitat de memòria ─────────────────────────────────
# Reinicia workers cada 500 peticions per prevenir fuites de memòria.
max_requests = 500
max_requests_jitter = 50

# ─── Seguretat ──────────────────────────────────────────────
# Només acceptar capçaleres X-Forwarded-For de localhost (nginx).
forwarded_allow_ips = "127.0.0.1"
