import pandas as pd
import re

class ExtratorRelatorioConvenio:
    def __init__(self, caminho_excel):
        self.caminho_excel = caminho_excel
        # Carrega o Excel sem cabeçalho definido para ler as linhas puras
        self.df = pd.read_excel(caminho_excel, header=None)
        self.dados_processados = []

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
            # Correção aplicada: Data na coluna G (índice 6) e Valor na coluna H (índice 7)
            data_procedimento = linha[6]   # Coluna G
            valor_procedimento = linha[7]  # Coluna H
            
            if beneficiario_atual and pd.notna(valor_procedimento):
                # Evita capturar linhas de totais ou cabeçalhos do meio do relatório
                if "Total" in celula_a or "Competência" in celula_a or celula_a == "":
                    continue
                    
                self.dados_processados.append({
                    "Codigo Titular": codigo_titular_atual,
                    "Titular": titular_atual,
                    "Beneficiário/Dependente": beneficiario_atual,
                    "Grau Parentesco": grau_atual,
                    "Tipo Servico": celula_a,
                    "Data": data_procedimento,
                    "Valor Coparticipação": valor_procedimento
                })

    def exportar_para_flat_table(self, caminho_saida):
        """Transforma a lista de dicionários em uma tabela limpa e exporta"""
        df_final = pd.DataFrame(self.dados_processados)
        df_final.to_excel(caminho_saida, index=False)
        print(f"Processamento concluído! Arquivo salvo em: {caminho_saida}")
        return df_final

# Para testar localmente no seu terminal, você pode descomentar as linhas abaixo:
# if __name__ == "__main__":
#     extrator = ExtratorRelatorioConvenio("seu_arquivo_origem.xlsx")
#     extrator.processar_planilha()
#     extrator.exportar_para_flat_table("coparticipacao_final.xlsx")