import pandas as pd
import pyperclip

caminho_arquivo_excel = r"C:\Users\t2s-ailton\Desktop\Análise de dados - Inicial.xlsx"

def ler_excel(caminho_arquivo, nome_planilha):
    try:
        df = pd.read_excel(caminho_arquivo, sheet_name=nome_planilha)
        return df
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return None

def obter_dados_products():
    nome_planilha = 'CadastroProdutos'
    dados_excel = ler_excel(caminho_arquivo_excel, nome_planilha)
    result = "BEGIN\n"

    for produto in dados_excel['Produto'].values:
            # Usa placeholders (?) para evitar injeção de SQL
            sql_statement = "INSERT INTO PRODUCTS (NAME) VALUES ('"+produto.replace("'","")+"');"
            # Imprime a instrução SQL formatada com o valor atual
            result += sql_statement

    result += "END;\n"
    return result

def obter_dados_production():
    nome_planilha = 'DadosProducao'
    dados_excel = ler_excel(caminho_arquivo_excel, nome_planilha)

    result = "BEGIN\n"
    for index, row in dados_excel.iterrows():
        codigo_produto = row['Codigo Produto']
        qtd_produzida = row['Qtd Produzida']
        emissao = row['Emissão']

        result += f"INSERT INTO PRODUCTION (ID_PRODUCT, QTD, DT_CREATION) VALUES ('{codigo_produto}', {qtd_produzida}, TO_DATE('{emissao}', 'YYYY-MM-DD HH24:MI:SS'));\n"

    result += "END;\n"
    return result
def main():
    menu_selecionado = input("1 - OBTER DADOS COMPOSTOS\n2 - OBTER DADOS COMPOSTOS\n")
    if menu_selecionado == '1':
        pyperclip.copy(obter_dados_products())
    if menu_selecionado == '2':
        pyperclip.copy(obter_dados_production())

    print("Conteúdo copiado para a área de transferência.")

if __name__ == "__main__":
    main()
    input("Pressione Enter para sair...")