#!/bin/bash
set -e

echo "=========================================="
echo "  Деплой Штаб отделки на сервер"
echo "=========================================="

# Переменные
PROJECT_DIR="/var/www/shtab-otdelki"
REPO_URL="https://github.com/qqJonni/app_otdelka.git"
DOMAIN="shtab-otdelki.ru"

echo ""
echo "[1/10] Обновление системы..."
apt update && apt upgrade -y

echo ""
echo "[2/10] Установка зависимостей..."
apt install python3 python3-pip python3-venv nginx git sqlite3 certbot python3-certbot-nginx -y

echo ""
echo "[3/10] Клонирование проекта..."
mkdir -p $PROJECT_DIR
cd $PROJECT_DIR
git clone $REPO_URL .

echo ""
echo "[4/10] Настройка виртуального окружения..."
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
pip install gunicorn

echo ""
echo "[5/10] Создание папок для загрузок..."
mkdir -p static/avatars static/uploads

echo ""
echo "[6/10] Создание .env файла..."
SECRET_KEY=$(python3 -c "import secrets; print(secrets.token_hex(32))")
cat > .env << EOF
FLASK_ENV=production
SECRET_KEY=$SECRET_KEY
EOF
echo "SECRET_KEY сгенерирован: $SECRET_KEY"

echo ""
echo "[7/10] Настройка systemd-сервиса..."
cat > /etc/systemd/system/shtab-otdelki.service << EOF
[Unit]
Description=Штаб отделки
After=network.target

[Service]
User=root
WorkingDirectory=$PROJECT_DIR
ExecStart=$PROJECT_DIR/venv/bin/gunicorn --workers 2 --bind 127.0.0.1:8080 app:app
Restart=always

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl start shtab-otdelki
systemctl enable shtab-otdelki
echo "Сервис запущен"

echo ""
echo "[8/10] Настройка Nginx..."
cat > /etc/nginx/sites-available/shtab-otdelki << EOF
server {
    listen 80;
    server_name $DOMAIN www.$DOMAIN;
    client_max_body_size 50M;

    location /static/ {
        alias $PROJECT_DIR/static/;
    }

    location / {
        proxy_pass http://127.0.0.1:8080;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
    }
}
EOF

ln -sf /etc/nginx/sites-available/shtab-otdelki /etc/nginx/sites-enabled/
rm -f /etc/nginx/sites-enabled/default
nginx -t
systemctl restart nginx
echo "Nginx настроен"

echo ""
echo "[9/10] Настройка SSL-сертификата..."
read -p "Введите email для SSL-сертификата (Let's Encrypt): " SSL_EMAIL
certbot --nginx -d $DOMAIN -d www.$DOMAIN --non-interactive --agree-tos --email "$SSL_EMAIL"
echo "SSL-сертификат получен"

echo ""
echo "[10/10] Проверка..."
systemctl status shtab-otdelki --no-pager
echo ""
echo "=========================================="
echo "  Деплой завершён!"
echo "  Сайт: https://$DOMAIN"
echo "  Логин: admin / admin123"
echo "=========================================="
