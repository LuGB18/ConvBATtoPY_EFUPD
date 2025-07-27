#Importar Bibliotecas
import os
from shutil import move, rmtree
from glob import glob
from zipfile import ZipFile
# Definir Variáveis: extensões de arquivos apagadas no .bat
extsimp = [
    '*.ini', '*.doc', '*.exe', '*.chm', '*.txt', '*.log'
]

extdoc = [
    '*.pdf', '*.csv', '*.xls', '*.rpf', '*.rar', '*.zip'
]

extespc = [
    'Plano de Contas Básico.doc', 'Formatos dos Códigos.doc',
    'EFBackup1160010.exe', 'EFBackup1160011.exe', 'IBPT.exe',
    'Eficaz3*.exe', 'Eficaz - old*.exe', '*old*.exe', 'Eficaz4*.exe',
    'EFUpdate1*.exe', 'email.exe', 'Impressão*', 'Schemas.exe',
    'EFBackup1*.exe', 'dll.exe', 'Atualizacao0*.exe', 'Atualizacao0*.rar',
    'EficazHelp0*.chm', 'Arquivos/Plano*.doc', 'Arquivos/Forma*.doc',
    'EFUpdate/*exe', 'EFUpdate/*txt', 'EFUpdate/EficazHelp0*.chm',
    'EFUpdate/EficazHelp0*.pdf', 'Utilitarios/Firebird-1*', 'Utilitarios/Firebird-2*',
    'Utilitarios/Firebird-3*.', 'Utilitarios/Firebird-4*.', 'Utilitarios/InstaladorCadeiaV2*',
    'Utilitarios/InstaladorCadeiaV5*', 'Utilitarios/Team*.exe', 'Utilitarios/IBExpert_2013*zip',
    'Utilitarios/IBExpert_2013*rar', 'Utilitarios/IBExpert_2010*exe', 'Utilitarios/IBExpert_2010*rar',
    'Utilitarios/EFBackup1160001.exe', 'Utilitarios/EFBackup1160002.exe',
    'Utilitarios/EFBackup1160003.exe', 'Utilitarios/EFBackup1160004.exe',
    'Utilitarios/EFBackup1160005.exe', 'Utilitarios/EFBackup1160006.exe',
    'Utilitarios/EFBackup1160007.exe', 'Utilitarios/EFBackup1160008.exe',
    'Utilitarios/EFBackup1160009.exe', 'Utilitarios/EFBackup1160010.exe',
    'Utilitarios/Suporte*', 'Utilitarios/Show*', 'Utilitarios/Atualizacao011*.',
    'Utilitarios/Atualizacao012*.', 'Utilitarios/Atualizacao013*.', 'Utilitarios/Atualizacao014*.',
    'Utilitarios/Atualizacao015*', 'Utilitarios/Atualizacao016*', 'Utilitarios/Atualizacao017*',
    'Utilitarios/Atualizacao018*', 'Utilitarios/Atualizacao019*', 'Utilitarios/Atualizacao020*',
    'EFupdate2*', 'EFupdate1*', 'EFupdateB*', 'Utilitarios/Atualizacao.exe',
    'Utilitarios/AA*.exe', 'Utilitarios/AA*.log', 'Utilitarios/Ammy*.exe',
    'Utilitarios/Ammy*.log', 'Atualizacao/*.sql', 'EFUpdate/Servicos/*zip',
    'EFUpdate/"Versao Beta"/*zip'
]


print("Efetuando Limpeza de Base")
os.chdir('..')
os.mkdir('Lixo')

for ext in extdoc:
    for file in glob(ext):    
        move(file,f'Lixo/{file}')


with ZipFile('Lixo.zip', 'w') as fzip:
    os.chdir('Lixo')
    for file in glob('*.*'):
        fzip.write(file)
    os.chdir('..')
rmtree('Lixo')

os.chdir('Arquivos')
for fileglb in glob('*.*'):
    os.remove(fileglb)
os.chdir('..')
for ext in extespc:
    for file in glob(ext):
        if os.path.isfile(file):
            os.remove(file)
        elif os.path.isdir(file):
            rmtree(file)
os.chdir('Arquivos')
os.startfile('Capicom.exe')

# A partir da linha 86 do .bat (ECHO *** Atualizando Capicom ***)
print("*** Atualizando Capicom ***")
if os.path.exists('Capicom'):
    rmtree('Capicom')
os.system('Capicom.exe')

import platform

# Detecta arquitetura do Windows
is_64bits = platform.machine().endswith('64')

print("*** Copiando as DLLs ***")
if is_64bits:
    syswow64 = os.path.join(os.environ['windir'], 'SysWOW64')
    dlls = [
        ('capicom.dll', syswow64),
        ('msxml5.dll', syswow64),
        ('msxml5r.dll', syswow64)
    ]
else:
    system32 = os.path.join(os.environ['windir'], 'System32')
    dlls = [
        ('capicom.dll', system32),
        ('msxml5.dll', system32),
        ('msxml5r.dll', system32)
    ]

for dll, destino in dlls:
    origem = os.path.join('Capicom', dll)
    destino_arquivo = os.path.join(destino, dll)
    if not os.path.exists(destino_arquivo):
        try:
            os.makedirs(destino, exist_ok=True)
            move(origem, destino_arquivo)
        except Exception as e:
            print(f"Erro ao copiar {dll}: {e}")

print("*** Registrando as DLLs ***")
for dll, destino in dlls:
    dll_path = os.path.join(destino, dll)
    if os.path.exists(dll_path):
        os.system(f'regsvr32 "{dll_path}" /s')

print("*** Atualizando as DLLs ***")
if os.path.exists('dll'):
    rmtree('dll')
os.system('dll.exe')
if os.path.exists('dll'):
    for file in glob(os.path.join('dll', '*.*')):
        move(file, '../')
    rmtree('dll')

print("*** Atualizando as DLLs do EFServer ***")
def atualizar_dlls_efserver(destino_path):
    print("*** Atualizando as DLLs do EFServer ***")
    if os.path.exists('dll'):
        rmtree('dll')
    os.system('dll.exe')
    if os.path.exists('dll'):
        for file in glob(os.path.join('dll', '*.*')):
            move(file, destino_path)
        rmtree('dll')

efserver_path = os.path.abspath(os.path.join('..', '..', 'EFServer'))
atualizar_dlls_efserver(efserver_path)

efserver_path2 = os.path.abspath(os.path.join('..', 'EFServer'))
atualizar_dlls_efserver(efserver_path2)
print("*** Atualizando Schemas ***")
if os.path.exists('Schemas'):
    rmtree('Schemas')
schemas_parent = os.path.abspath(os.path.join('..', 'Schemas'))
if os.path.exists(schemas_parent):
    rmtree(schemas_parent)
os.system('Schemas.exe')
if os.path.exists('Schemas'):
    move('Schemas', os.path.abspath('..'))

print("*** Atualizando Schemas-NFSe ***")
if os.path.exists('Schemas-NFSe'):
    rmtree('Schemas-NFSe')
schemas_nfse_parent = os.path.abspath(os.path.join('..', 'Schemas-NFSe'))
if os.path.exists(schemas_nfse_parent):
    rmtree(schemas_nfse_parent)
os.system('Schemas-NFSe.exe')
if os.path.exists('Schemas-NFSe'):
    move('Schemas-NFSe', os.path.abspath('..'))

print("*** Atualizando Schemas-CTe ***")
if os.path.exists('Schemas-CTe'):
    rmtree('Schemas-CTe')
schemas_cte_parent = os.path.abspath(os.path.join('..', 'Schemas-CTe'))
if os.path.exists(schemas_cte_parent):
    rmtree(schemas_cte_parent)
os.system('Schemas-CTe.exe')
if os.path.exists('Schemas-CTe'):
    move('Schemas-CTe', os.path.abspath('..'))

print("*** Atualizando Schemas-MDF-e ***")
if os.path.exists('Schemas-MDFe'):
    rmtree('Schemas-MDFe')
schemas_mdfe_parent = os.path.abspath(os.path.join('..', 'Schemas-MDFe'))
if os.path.exists(schemas_mdfe_parent):
    rmtree(schemas_mdfe_parent)
os.system('Schemas-MDFe.exe')
if os.path.exists('Schemas-MDFe'):
    move('Schemas-MDFe', os.path.abspath('..'))

print("*** Atualizando Schemas-DFe ***")
if os.path.exists('Schemas-DFe'):
    rmtree('Schemas-DFe')
schemas_dfe_parent = os.path.abspath(os.path.join('..', 'Schemas-DFe'))
if os.path.exists(schemas_dfe_parent):
    rmtree(schemas_dfe_parent)
os.system('Schemas-DFe.exe')
if os.path.exists('Schemas-DFe'):
    move('Schemas-DFe', os.path.abspath('..'))

print("*** Limpando Logs ***")
for file in glob(os.path.join('..', 'Atualizacao', '*.sql')):
    os.remove(file)

print("*** Atualizando Openssl ***")
for openssl_file in ['openssl.cnf', 'openssl.exe']:
    origem = os.path.join('..', 'Utilitarios', openssl_file)
    destino = os.path.join('..', openssl_file)
    if os.path.exists(origem):
        move(origem, destino)

print("*** Atualizando Help ***")
help_src = os.path.join('..', 'EFUpdate', 'EficazHelp.chm')
help_dst = os.path.join('..', 'EficazHelp.chm')
if os.path.exists(help_src):
    move(help_src, help_dst)

print("Copiando TrocaExe.bat para EFUpdate")
trocaexe_src = 'TrocaExe.bat'
trocaexe_dst = os.path.join('..', 'EFUpdate', 'TrocaExe.bat')
if os.path.exists(trocaexe_src):
    move(trocaexe_src, trocaexe_dst)

for file in glob(os.path.join('..', 'EFUpdate', 'Servicos', '*zip')):
    os.remove(file)
for file in glob(os.path.join('..', 'EFUpdate', 'Versao Beta', '*zip')):
    os.remove(file)

print("*** Iniciando Backup em segundo plano ***")
os.system('start "" /b cmd')
print("*** Iniciando Backup em segundo plano ***")
os.system('start "" /b cmd /c "limpabackup.bat"')
