import easy_vault

def vault():
    vault_file = './/config.yml' 
    password = easy_vault.get_password(vault_file)
    vault = easy_vault.EasyVault(vault_file, password)
    easy_vault.set_password(vault_file, password)
    return vault.get_yaml()

config = vault()