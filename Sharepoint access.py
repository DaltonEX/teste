from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

# Configuração da URL do SharePoint e credenciais de autenticação
sharepoint_url = "https://deere.sharepoint.com/sites/itnw"
username = input("Username: ")
password = input("Senha: ")

# Criação do contexto do cliente SharePoint
ctx_auth = AuthenticationContext(sharepoint_url)
if ctx_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(sharepoint_url, ctx_auth)
    print("Autenticação bem-sucedida!")
else:
    print("Falha na autenticação.")