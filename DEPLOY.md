# Desplegament al VPS — indret_originals.masellas.info

## Prerequisits locals

Abans de connectar al servidor, fes push del codi:

```bash
git push InDret_Corrector master
```

---

## Al servidor (SSH)

### 1. Instal·lar dependències del sistema

```bash
sudo apt update && sudo apt install -y python3 python3-venv python3-pip
```

### 2. Clonar el repo

```bash
sudo mkdir -p /opt/indret-corrector
sudo chown $USER:$USER /opt/indret-corrector
cd /opt/indret-corrector
git clone https://github.com/rogeraclaro/indret-corrector-standalone.git .
```

### 3. Crear entorn virtual i instal·lar dependències

```bash
cd /opt/indret-corrector/web
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 -m spacy download es_core_news_lg
```

### 4. Crear el servei systemd

```bash
sudo nano /etc/systemd/system/indret-corrector.service
```

Contingut del fitxer:

```ini
[Unit]
Description=InDret Corrector — Gunicorn
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

Activar el servei:

```bash
sudo chown -R www-data:www-data /opt/indret-corrector
sudo systemctl daemon-reload
sudo systemctl enable --now indret-corrector
sudo systemctl status indret-corrector
```

### 5. Configurar nginx

```bash
sudo nano /etc/nginx/sites-available/indret-corrector
```

Contingut del fitxer:

```nginx
server {
    listen 80;
    server_name indret_originals.masellas.info;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl;
    server_name indret_originals.masellas.info;

    ssl_certificate     /etc/letsencrypt/live/indret_originals.masellas.info/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/indret_originals.masellas.info/privkey.pem;

    client_max_body_size 21M;

    location / {
        include proxy_params;
        proxy_pass http://unix:/tmp/indret-corrector.sock;
        proxy_read_timeout 130s;
        proxy_connect_timeout 10s;
    }
}
```

Activar i recarregar nginx:

```bash
sudo ln -s /etc/nginx/sites-available/indret-corrector /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
```

### 6. Verificar

```bash
curl -s -o /dev/null -w "%{http_code}" https://indret_originals.masellas.info/
# Ha de retornar 200
```

---

## Actualitzar el servidor (futures versions)

### Al Mac (local)
```bash
# Des del repo principal, pujar canvis al repo standalone:
cd "/Volumes/1Tera/Local Sites/indret-prod/app/public/wp-content/themes/indret"
git subtree push --prefix=web standalone main
```

### Al servidor
```bash
cd /opt/indret-corrector
git pull
sudo systemctl restart indret-corrector
```
