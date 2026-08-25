DASHBOARD SGPA - PUBLICAÇÃO

ACESSO PROTEGIDO
O dashboard agora exige login antes de liberar o painel e as rotas de exportação.

Usuário configurado:
FNP-ADMIN

Por segurança, a senha NÃO fica gravada no código nem neste arquivo.

VARIÁVEIS DE AMBIENTE NO RENDER
Configure estas três variáveis em Environment:
SECRET_KEY = uma chave longa e aleatória
DASHBOARD_USER = FNP-ADMIN
DASHBOARD_PASSWORD = a senha definida para o dashboard

IMPORTANTE:
- Não coloque a senha diretamente no app.py.
- Não publique um arquivo .env no GitHub.
- Depois de salvar as variáveis no Render, faça um novo deploy/redeploy.

Comando de inicialização recomendado:
gunicorn app:app

Para testar localmente:
python app.py
Acesse: http://127.0.0.1:5000

ROTAS DE ACESSO
/login  -> tela de login
/logout -> encerra a sessão
/       -> dashboard protegido
/exportar/* -> exportações protegidas
