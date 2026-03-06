# InDret — Corrector d'articles

Interfície web per al corrector de format d'articles de la revista InDret (UPF).

## Arrencada en local

### Requisits

- Python 3.11 o superior
- pip

### Instal·lació

```bash
cd web/
python3 -m venv .venv
source .venv/bin/activate       # Linux/macOS
# .venv\Scripts\activate        # Windows
pip install -r requirements.txt
```

### Execució

```bash
python app.py
```

Obrir el navegador a: http://localhost:5000

---

## Desplegament al VPS (gunicorn + nginx + systemd)

### 1. Preparació del servidor

```bash
# Instalar dependències del sistema (Ubuntu/Debian)
sudo apt update && sudo apt install python3 python3-venv python3-pip nginx -y
```

### 2. Clonar el repositori i crear l'entorn virtual

```bash
sudo mkdir -p /opt/indret-corrector
sudo chown www-data:www-data /opt/indret-corrector
cd /opt/indret-corrector
git clone <URL_REPO> .
cd web/
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

> **Nota:** Si el VPS no té `libmagic` instal·lat, `puremagic` (dependència pura Python) funciona sense cap paquet de sistema addicional.

### 3. Servei systemd per a Gunicorn

Crear `/etc/systemd/system/indret-corrector.service`:

```ini
[Unit]
Description=InDret Corrector d'articles — Gunicorn
After=network.target

[Service]
User=www-data
Group=www-data
WorkingDirectory=/opt/indret-corrector/web
Environment="PATH=/opt/indret-corrector/web/.venv/bin"
ExecStart=/opt/indret-corrector/web/.venv/bin/gunicorn -c gunicorn.conf.py app:app
Restart=on-failure
RestartSec=5

[Install]
WantedBy=multi-user.target
```

Activar i arrencar:

```bash
sudo systemctl daemon-reload
sudo systemctl enable indret-corrector
sudo systemctl start indret-corrector
sudo systemctl status indret-corrector
```

### 4. Configuració de nginx

Crear `/etc/nginx/sites-available/indret-corrector`:

```nginx
server {
    listen 80;
    server_name _;  # O: corrector.indret.com

    # Límit de pujada: lleugerament superior al límit de Flask (20 MB)
    # IMPORTANT: sense això, nginx rebutja fitxers grans amb 413 abans que Flask pugui
    client_max_body_size 21M;

    location / {
        include proxy_params;
        proxy_pass http://unix:/tmp/indret-corrector.sock;
        # IMPORTANT: lleugerament superior al timeout de gunicorn (120 s)
        # sense això, nginx retorna 502 en documents complexos
        proxy_read_timeout 130s;
        proxy_connect_timeout 10s;
    }
}
```

Activar:

```bash
sudo ln -s /etc/nginx/sites-available/indret-corrector /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
```

### 5. Verificació del desplegament

```bash
# Comprovar que gunicorn és actiu
sudo systemctl status indret-corrector

# Comprovar que el socket Unix existeix
ls -la /tmp/indret-corrector.sock

# Prova des del servidor
curl -s -o /dev/null -w "%{http_code}" http://localhost/
# Ha de retornar 200
```

---

## Notes importants

### Timeout del corrector

El corrector pot trigar 10-30 segons en documents complexos. `gunicorn.conf.py` ja configura `timeout = 120`. Si es veuen errors **502 Bad Gateway** en documents grans, verificar:
- `gunicorn.conf.py`: `timeout = 120` ✓
- nginx: `proxy_read_timeout 130s` ✓ (ha de ser lleugerament superior al de gunicorn)

### Límit de mida de fitxer

El límit de 20 MB s'aplica en dos llocs:
- Flask: `MAX_CONTENT_LENGTH = 20 * 1024 * 1024` (a `app.py`)
- nginx: `client_max_body_size 21M` (a la configuració nginx)

Si s'augmenta el límit de Flask, cal augmentar també el de nginx.

### HTTPS / SSL

Per afegir SSL amb Let's Encrypt:

```bash
sudo apt install certbot python3-certbot-nginx
sudo certbot --nginx -d corrector.indret.com
```

### Clau secreta de Flask

La clau secreta actual (`os.urandom(24)`) es regenera cada cop que l'app arrenca, invalidant les sessions actives. Per a producció, establir una clau fixa:

```bash
# Generar una clau segura
python3 -c "import secrets; print(secrets.token_hex(32))"
```

Afegir com a variable d'entorn `SECRET_KEY` al fitxer de servei systemd i actualitzar `app.py`:
```python
app.secret_key = os.environ.get('SECRET_KEY', os.urandom(24))
```
