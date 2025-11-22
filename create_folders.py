import os

# Extrai o caminho onde as pastas serão criadas
folder_path = os.getcwd()

# Lista de pastas que serão criadas
folders = [folder_path+'\\Dados brutos', folder_path+'\\Relatórios', folder_path+'\\Dados tratados' ]

# Lista de pastas que serão criadas dentro das pastas principais
subfolders = [folders[1]+'\\Envio']

# Criação das pastas
for folder in folders:
    if not os.path.exists(folder):
        os.makedirs(folder)

# Criação das subpastas
for folder in subfolders:
    if not os.path.exists(folder):
        os.makedirs(folder)        
