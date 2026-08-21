DASHBOARD SUPER ROTA - PUBLICAÇÃO

Login padrão inicial:
Usuário: admin
Senha: 123456

IMPORTANTE: antes de publicar, altere as variáveis de ambiente:
SECRET_KEY
DASHBOARD_USER
DASHBOARD_PASSWORD

Comando de inicialização recomendado:
gunicorn app:app

Para testar localmente:
python app.py
Acesse: http://127.0.0.1:5000
