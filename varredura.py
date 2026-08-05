import pandas as pd
import re
import unicodedata
import os
import glob

class ExtratorRelatorioConvenio:
    def __init__(self, caminho_excel):
        self.caminho_excel = caminho_excel
        self.df = pd.read_excel(caminho_excel, header=None, engine='xlrd')
        self.dados_processados = []
        self.df_erros = pd.DataFrame()

    def converter_para_float(self, valor):
        if pd.isna(valor) or str(valor).strip() == "":
            return 0.0
        try:
            return float(str(valor).replace(',', '.'))
        except ValueError:
            return 0.0

    def remover_acentos(self, texto):
        """Remove acentos e padroniza o texto em maiúsculas para evitar falhas no merge"""
        if pd.isna(texto):
            return ""
        texto_str = str(texto).strip().upper()
        nfkd_form = unicodedata.normalize('NFKD', texto_str)
        return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

    def processar_planilha(self):
        titular_atual = None
        codigo_titular_atual = None
        beneficiario_atual = None
        grau_atual = None
        
        for index, linha in self.df.iterrows():
            celula_a = str(linha[0]).strip() if pd.notna(linha[0]) else ""
            
            # 1. Identifica o Bloco do Titular
            if celula_a.startswith("Cod. Titular:"):
                match = re.search(r"Cod\. Titular:\s*(\d+)\s*-\s*(.*)", celula_a)
                if match:
                    codigo_titular_atual = match.group(1)
                    titular_atual = match.group(2).strip()
                continue
                
            # 2. Identifica o Beneficiário (Titular ou Dependente)
            if celula_a.startswith("Beneficiário:"):
                match = re.search(r"Beneficiário:\s*(\d+)\s*-\s*([^\(]+)\s*\((\d+)\)", celula_a)
                if match:
                    beneficiario_atual = match.group(2).strip()
                    grau_atual = match.group(3).strip()
                continue
            
            # 3. Captura os procedimentos associados (Colunas I e J)
            valor_coluna_i = linha[8]  
            valor_coluna_j = linha[9]  
            
            if beneficiario_atual and (pd.notna(valor_coluna_i) or pd.notna(valor_coluna_j)):
                if "Total" in celula_a or "Competência" in celula_a or celula_a == "":
                    continue
                
                v_i = self.converter_para_float(valor_coluna_i)
                v_j = self.converter_para_float(valor_coluna_j)
                total_procedimento = v_i + v_j
                
                if total_procedimento == 0.0:
                    continue
                    
                self.dados_processados.append({
                    "Codigo Titular": codigo_titular_atual,
                    "Titular": titular_atual,
                    "Beneficiário/Dependente": beneficiario_atual,
                    "Grau Parentesco": grau_atual,
                    "Valor Coparticipação": total_procedimento
                })

    def consolidar_por_beneficiario(self):
        if not self.dados_processados:
            print("Nenhum dado processado para consolidar.")
            return pd.DataFrame()
            
        df_detalhado = pd.DataFrame(self.dados_processados)
        
        df_consolidado = df_detalhado.groupby(
            ["Codigo Titular", "Titular", "Beneficiário/Dependente", "Grau Parentesco"], 
            as_index=False
        )["Valor Coparticipação"].sum()
        
        return df_consolidado.round({"Valor Coparticipação": 2})

    def mesclar_dados_rh(self, df_consolidado, caminho_rh):
        print("Carregando dados cadastrais do RH...")
        df_rh = pd.read_excel(caminho_rh, dtype={'CHAPA': str})
        
        df_rh['Nome_Chave'] = df_rh['NOME DEPENDENTE'].astype(str).str.strip()
        condicao_vazio = df_rh['Nome_Chave'].isna() | (df_rh['Nome_Chave'] == '') | (df_rh['Nome_Chave'].str.upper().isin(['NAN', 'NONE']))
        df_rh.loc[condicao_vazio, 'Nome_Chave'] = df_rh['NOME TITULAR']
        
        # Padronização para o cruzamento
        df_rh['Chave_Match'] = df_rh['Nome_Chave'].apply(self.remover_acentos)
        df_consolidado['Chave_Match'] = df_consolidado['Beneficiário/Dependente'].apply(self.remover_acentos)
        
        df_rh_reduzido = df_rh[['Chave_Match', 'CHAPA', 'NÚMERO DEPENDENTE']].drop_duplicates(subset=['Chave_Match'])
        
        print("Cruzando informações com o RH...")
        df_final = pd.merge(df_consolidado, df_rh_reduzido, on="Chave_Match", how="left")
        
        if 'NÚMERO DEPENDENTE' in df_final.columns:
            df_final['NÚMERO DEPENDENTE'] = df_final['NÚMERO DEPENDENTE'].fillna('').apply(
                lambda x: str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) else str(x)
            )
            df_final.loc[df_final['NÚMERO DEPENDENTE'] == 'nan', 'NÚMERO DEPENDENTE'] = ''

        colunas_ordenadas = [
            "CHAPA", "Titular", "NÚMERO DEPENDENTE", "Beneficiário/Dependente", 
            "Grau Parentesco", "Codigo Titular", "Valor Coparticipação"
        ]
        df_final = df_final.reindex(columns=colunas_ordenadas)
        
        # Isola quem ficou sem CHAPA/Número para gerar relatório de inconsistência
        self.df_erros = df_final[df_final['CHAPA'].isna() | (df_final['CHAPA'] == '')].copy()
        
        sem_chapa = len(self.df_erros)
        if sem_chapa > 0:
            print(f"\n⚠️ Atenção: {sem_chapa} beneficiário(s) não foram localizados no arquivo do RH.")
            print("Eles serão salvos na aba separada: 'CONFERIR_NO_RH'.")
            
            df_rh['Primeiro_Nome_Titular'] = df_rh['NOME TITULAR'].apply(lambda x: str(x).split()[0].upper() if pd.notna(x) else "")
            dict_primeiro_nome = df_rh.drop_duplicates(subset=['Primeiro_Nome_Titular']).set_index('Primeiro_Nome_Titular')['CHAPA'].to_dict()
            
            def sugerir_chapa(titular):
                if pd.isna(titular): return ""
                p_nome = str(titular).split()[0].upper()
                return dict_primeiro_nome.get(p_nome, "Não Encontrado")
                
            self.df_erros['Possivel_Chapa_Empresa'] = self.df_erros['Titular'].apply(sugerir_chapa)
        else:
            self.df_erros = pd.DataFrame()
            
        return df_final

    def exportar_para_flat_table(self, dataframe, caminho_saida):
        """Salva o resultado gerando duas abas caso existam erros de cadastro"""
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            dataframe.to_excel(writer, sheet_name='Consolidado_Folha', index=False)
            
            if not self.df_erros.empty:
                self.df_erros[['Titular', 'Beneficiário/Dependente', 'Valor Coparticipação', 'Possivel_Chapa_Empresa']].to_excel(
                    writer, sheet_name='CONFERIR_NO_RH', index=False
                )
        print(f"Processamento concluído com sucesso! Arquivo salvo em: {caminho_saida}")

def buscar_arquivo_mais_recente(pasta=".dados", extensao="*.xls"):
    """Encontra o arquivo .xls modificado mais recentemente na pasta informada"""
    caminho_busca = os.path.join(pasta, extensao)
    arquivos = glob.glob(caminho_busca)
    
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo {extensao} foi encontrado na pasta '{pasta}'.")
        
    # Ordena os arquivos pela data de modificação (mais recente primeiro)
    arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
    return arquivo_mais_recente


if __name__ == "__main__":
    pasta_dados = ".dados"
    arquivo_rh = os.path.join(pasta_dados, "Chapa colab e Dados dependentes.xlsx")
    arquivo_saida = os.path.join(pasta_dados, "coparticipacao_consolidada_rh.xlsx")
    
    try:
        # Busca automática do arquivo da operadora mais recente baixado na pasta .dados
        arquivo_real = buscar_arquivo_mais_recente(pasta=pasta_dados, extensao="*.xls")
        print(f"🔍 Arquivo mais recente detectado automaticamente: '{arquivo_real}'")
        
        print("Iniciando a varredura do arquivo...")
        extrator = ExtratorRelatorioConvenio(arquivo_real)
        extrator.processar_planilha()
        
        print("Consolidando valores por beneficiário...")
        df_consolidado = extrator.consolidar_por_beneficiario()
        
        df_com_chapa = extrator.mesclar_dados_rh(df_consolidado, arquivo_rh)
        extrator.exportar_para_flat_table(df_com_chapa, arquivo_saida)
        
    except Exception as e:
        print(f"\n❌ Erro durante a execução: {e}")