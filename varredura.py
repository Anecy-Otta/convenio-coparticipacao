import pandas as pd
import re

class ExtratorRelatorioConvenio:
    def __init__(self, caminho_excel):
        self.caminho_excel = caminho_excel
        self.df = pd.read_excel(caminho_excel, header=None, engine='xlrd')
        self.dados_processados = []

    def converter_para_float(self, valor):
        """Auxiliar para tratar e converter os valores das células para float com segurança"""
        if pd.isna(valor) or str(valor).strip() == "":
            return 0.0
        try:
            return float(str(valor).replace(',', '.'))
        except ValueError:
            return 0.0

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
                    titular_atual = match.group(2)
                continue
                
            # 2. Identifica o Beneficiário (Titular ou Dependente)
            if celula_a.startswith("Beneficiário:"):
                match = re.search(r"Beneficiário:\s*(\d+)\s*-\s*([^\(]+)\s*\((\d+)\)", celula_a)
                if match:
                    beneficiario_atual = match.group(2).strip()
                    grau_atual = match.group(3).strip()
                continue
            
            # 3. Captura os procedimentos associados ao beneficiário atual
            # AJUSTE: Ignorando coluna H (índice 7). Capturando coluna I (índice 8) e J (índice 9)
            valor_coluna_i = linha[8]  # Coluna I
            valor_coluna_j = linha[9]  # Coluna J
            
            # O gatilho de execução verifica se há valores em alguma das duas colunas alvo
            if beneficiario_atual and (pd.notna(valor_coluna_i) or pd.notna(valor_coluna_j)):
                if "Total" in celula_a or "Competência" in celula_a or celula_a == "":
                    continue
                
                # Converte e soma os valores das duas colunas para compor o total do procedimento
                v_i = self.converter_para_float(valor_coluna_i)
                v_j = self.converter_para_float(valor_coluna_j)
                total_procedimento = v_i + v_j
                
                # Se o procedimento resultou em valor zerado em ambas, pula para otimizar a tabela
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
        """Agrupa os procedimentos somando os valores por beneficiário único"""
        if not self.dados_processados:
            print("Nenhum dado processado para consolidar.")
            return pd.DataFrame()
            
        df_detalhado = pd.DataFrame(self.dados_processados)
        
        # Agrupa pelas colunas de identificação e soma a coluna de valores combinados
        df_consolidado = df_detalhado.groupby(
            ["Codigo Titular", "Titular", "Beneficiário/Dependente", "Grau Parentesco"], 
            as_index=False
        )["Valor Coparticipação"].sum()
        
        # Arredonda o valor final para duas casas decimais
        df_consolidado = df_consolidado.round({"Valor Coparticipação": 2})
        
        return df_consolidado

    def exportar_para_flat_table(self, dataframe, caminho_saida):
        """Exporta o dataframe gerado para o Excel"""
        dataframe.to_excel(caminho_saida, index=False)
        print(f"Processamento e consolidação concluídos! Arquivo salvo em: {caminho_saida}")


if __name__ == "__main__":
    arquivo_real = ".dados/SLIP1-012735300-202605.xls" 
    arquivo_saida = ".dados/coparticipacao_consolidada.xlsx"
    
    print("Iniciando a varredura do arquivo real com colunas I e J...")
    extrator = ExtratorRelatorioConvenio(arquivo_real)
    extrator.processar_planilha()
    
    print("Consolidando valores por beneficiário...")
    df_consolidado = extrator.consolidar_por_beneficiario()
    
    extrator.exportar_para_flat_table(df_consolidado, arquivo_saida)