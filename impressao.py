import win32print
import win32api
import os
from tqdm import tqdm

limpar = lambda: os.system('cls')

# escolher qual impressora a gente vai querer usar
lista_impressoras = win32print.EnumPrinters(2)
tamanho = len(lista_impressoras)


escolha_int = 3000
while escolha_int > int(tamanho):
    limpar()
    print('###########################')
    print('# Impressoras disponíveis #')
    print('###########################\n')
    numero = 0
    for impressoras in lista_impressoras:
        print(f'{numero} - {impressoras}')
        numero = numero + 1
    escolha = input('\nEscolha uma impressora: ')
    escolha_int = int(escolha)

    while escolha_int > int(tamanho):
        limpar()
        print('#######################################')
        print(f"# Erro: Escolha um número entre 0 e {tamanho} #")
        print('#######################################\n')
        numero = 0
        for impressoras in lista_impressoras:
            print(f'{numero} - {impressoras}')
            numero = numero + 1
        escolha = input('\nEscolha uma impressora: ')
        escolha_int = int(escolha)

    if escolha_int in range(int(tamanho)):
        impressora_escolhida = (f"Você escolheu a impressora {lista_impressoras[escolha_int]}")
        impressora = lista_impressoras[escolha_int] 
        win32print.SetDefaultPrinter(impressora[2])
    else:
        limpar()
        print('#######################################')
        print(f"# Erro: Escolha um número entre 0 e {tamanho - 1} #")
        print('#######################################\n')
        numero = 0
        for impressoras in lista_impressoras:
            print(f'{numero} - {impressoras}')
            numero = numero + 1
        escolha = input('\nEscolha uma impressora: ')
        escolha_int = int(escolha)
 
# mandar imprimir todos os arquivos de uma pasta
limpar()
print('\n############################################')
print('# Caminho dos arquivos que serão impressos #')
print('############################################\n')
caminho = input('Caminho onde estão os arquivos que serão impressos: ')

limpar()
print('\n################################')
print('# Arquivos que serão impressos #')
print('################################\n')
print(f'{impressora_escolhida}\n')
print('Os arquivos abaixo serão impressos:\n')

lista_arquivos = os.listdir(caminho)
for arquivo in lista_arquivos:
    print(arquivo)

enter = input("\nAperte enter para continuar: ")


for arquivo in tqdm(lista_arquivos):
    limpar()
    print("#######################")
    print("# Imprimindo arquivos #")
    print("#######################\n")
    print(f'{impressora_escolhida}\n')
    print(f"Imprimindo: {arquivo}")
    win32api.ShellExecute(0, "print", arquivo, None, caminho, 0)

limpar()
print("##########################")
print("# Procedimento concluído #")
print("##########################")
enter = input("\nAperte enter para fechar: ")  