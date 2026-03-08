# Com actualitzar el servidor

## 1. Al Mac — pujar canvis al repo standalone

```bash
cd "/Volumes/1Tera/Local Sites/indret-prod/app/public/wp-content/themes/indret"
git subtree push --prefix=web standalone main
```

## 2. Al servidor (SSH)

```bash
cd /home/masellas-indret-originals/htdocs/indret-originals.masellas.info
git pull
systemctl restart indret-corrector
```

## Verificar que funciona

```bash
systemctl status indret-corrector
```

O obre el navegador a: https://indret-originals.masellas.info
