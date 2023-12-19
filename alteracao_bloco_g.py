# Abra o arquivo SPED EFD em modo de leitura
with open('c://temp/seuarquivo.txt', 'r', encoding='latin-1') as arquivo:
    linhas = arquivo.readlines()

# Lista para armazenar as linhas modificadas
novas_linhas = []

for linha in linhas:
    campos = linha.strip().split('|')
    print(campos[1])
    if campos[1] == 'G125':
        # Incrementa o número da parcela em 1
        numero_parcela = int(campos[9])+1
        print(numero_parcela)
        # Atualiza a linha com o novo número da parcela
        nova_linha = f'|G125|{campos[2]}|{campos[2]}|{numero_parcela}|{campos[9]}\n'
        novas_linhas.append(nova_linha)
    else:
        novas_linhas.append(linha)

# Abre o arquivo em modo de escrita e escreve as linhas modificadas
with open('c://temp/seuarquivo_modificado.txt', 'w',encoding='latin-1') as arquivo_modificado:
    arquivo_modificado.writelines(novas_linhas)
