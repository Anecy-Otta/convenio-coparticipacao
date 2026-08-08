import pandas as pd
import re
import unicodedata
import os
import glob
from openpyxl import load_workbook

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
            
            # 3. Captura das Colunas I (índice 8) e J (índice 9)
            valor_coluna_i = linha[8] if len(linha) > 8 else None
            valor_coluna_j = linha[9] if len(linha) > 9 else None
            
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
                    "Tipo Servico": celula_a,
                    "Valor Coparticipação": total_procedimento
                })

    def consolidar_por_beneficiario(self, ignorar_internacoes=False):
        if not self.dados_processados:
            print("Nenhum dado processado para consolidar.")
            return pd.DataFrame()
            
        df_detalhado = pd.DataFrame(self.dados_processados)
        
        if ignorar_internacoes:
            df_detalhado = df_detalhado[~df_detalhado['Tipo Servico'].str.upper().str.contains('INTERNA', na=False)]
        
        df_consolidado = df_detalhado.groupby(
            ["Codigo Titular", "Titular", "Beneficiário/Dependente", "Grau Parentesco"], 
            as_index=False
        )["Valor Coparticipação"].sum()
        
        return df_consolidado.round({"Valor Coparticipação": 2})

    def mesclar_dados_rh(self, df_consolidado, caminho_rh):
        print("Carregando dados cadastrais do RH...")
        df_rh = pd.read_excel(caminho_rh, dtype={'CHAPA': str})
        
        coluna_tipo = None
        for col in df_rh.columns:
            if col.strip().upper() in ['TIPO', 'TIPO ']:
                coluna_tipo = col
                break
                
        if coluna_tipo is None:
            coluna_tipo = df_rh.columns[4]
            
        df_rh['TIPO_CADASTRO'] = df_rh[coluna_tipo].astype(str).str.strip().str.upper()
        
        df_rh['Nome_Chave'] = df_rh['NOME DEPENDENTE'].astype(str).str.strip()
        condicao_vazio = df_rh['Nome_Chave'].isna() | (df_rh['Nome_Chave'] == '') | (df_rh['Nome_Chave'].str.upper().isin(['NAN', 'NONE']))
        df_rh.loc[condicao_vazio, 'Nome_Chave'] = df_rh['NOME TITULAR']
        
        df_rh['Chave_Match'] = df_rh['Nome_Chave'].apply(self.remover_acentos)
        df_consolidado['Chave_Match'] = df_consolidado['Beneficiário/Dependente'].apply(self.remover_acentos)
        
        df_rh_reduzido = df_rh[['Chave_Match', 'CHAPA', 'NÚMERO DEPENDENTE', 'TIPO_CADASTRO']].drop_duplicates(subset=['Chave_Match'])
        
        df_final = pd.merge(df_consolidado, df_rh_reduzido, on="Chave_Match", how="left")
        
        if 'NÚMERO DEPENDENTE' in df_final.columns:
            df_final['NÚMERO DEPENDENTE'] = df_final['NÚMERO DEPENDENTE'].fillna('').apply(
                lambda x: str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) else str(x)
            )
            df_final.loc[df_final['NÚMERO DEPENDENTE'] == 'nan', 'NÚMERO DEPENDENTE'] = ''

        colunas_ordenadas = [
            "CHAPA", "Titular", "NÚMERO DEPENDENTE", "Beneficiário/Dependente", 
            "Grau Parentesco", "Codigo Titular", "Valor Coparticipação", "TIPO_CADASTRO"
        ]
        df_final = df_final.reindex(columns=colunas_ordenadas)
        
        self.df_erros = df_final[df_final['CHAPA'].isna() | (df_final['CHAPA'] == '')].copy()
        
        sem_chapa = len(self.df_erros)
        if sem_chapa > 0:
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

    def aplicar_teto_por_chapa(self, df_tratado, teto_maximo=150.0):
        """Aplica o teto limite de R$ 150,00 por CHAPA dividindo proporcionalmente ou travando o grupo"""
        print(f"🛡️ Aplicando trava de teto máximo de R$ {teto_maximo:.2f} por CHAPA...")
        df = df_tratado.copy()
        
        # Calcula o total acumulado por CHAPA
        totais_por_chapa = df.groupby('CHAPA')['Valor Coparticipação'].transform('sum')
        
        # Aplica o ajuste proporcional para quem estourar o teto
        condicao_estouro = totais_por_chapa > teto_maximo
        
        df.loc[condicao_estouro, 'Valor Coparticipação'] = (
            (df.loc[condicao_estouro, 'Valor Coparticipação'] / totais_por_chapa[condicao_estouro]) * teto_maximo
        ).round(2)
        
        chapas_afetadas = df_tratado.loc[condicao_estouro, 'CHAPA'].nunique()
        print(f"   ➜ {chapas_afetadas} CHAPA(s) tiveram o valor reajustado para respeitar o teto de R$ {teto_maximo:.2f}.")
        
        return df

    def exportar_para_flat_table(self, dataframe, caminho_saida):
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            dataframe.to_excel(writer, sheet_name='Consolidado_Folha', index=False)
            
            if not self.df_erros.empty:
                self.df_erros[['Titular', 'Beneficiário/Dependente', 'Valor Coparticipação', 'Possivel_Chapa_Empresa']].to_excel(
                    writer, sheet_name='CONFERIR_NO_RH', index=False
                )
        print(f"📊 Arquivo de auditoria salvo em: {caminho_saida}")

    def preencher_planilha_base(self, df_consolidado, caminho_base_modelo, caminho_saida_base, ano, mes, data_pagamento):
        wb = load_workbook(caminho_base_modelo)
        ws = wb.active  
        
        linha_inicial = 7
        max_linha = max(ws.max_row, linha_inicial + 100)
        colunas_para_limpar = ['C', 'E', 'G', 'I', 'K', 'M', 'O', 'Q', 'S', 'U', 'W', 'Y', 'AA', 'AC', 'AE']
        
        for l in range(linha_inicial, max_linha + 1):
            for col in colunas_para_limpar:
                ws[f'{col}{l}'] = None

        for idx, row in df_consolidado.reset_index(drop=True).iterrows():
            linha_atual = linha_inicial + idx
            
            chapa_val = str(row['CHAPA']) if pd.notna(row['CHAPA']) else ""
            dep_val = str(row['NÚMERO DEPENDENTE']) if pd.notna(row['NÚMERO DEPENDENTE']) else ""
            valor_val = float(row['Valor Coparticipação']) if pd.notna(row['Valor Coparticipação']) else 0.0
            
            ws[f'C{linha_atual}'] = chapa_val             
            ws[f'E{linha_atual}'] = ano                   
            ws[f'G{linha_atual}'] = mes                   
            ws[f'I{linha_atual}'] = 5                     
            ws[f'K{linha_atual}'] = dep_val               
            ws[f'M{linha_atual}'] = "0176"                
            ws[f'O{linha_atual}'] = 0                     
            ws[f'Q{linha_atual}'] = data_pagamento        
            ws[f'S{linha_atual}'] = ano                   
            ws[f'U{linha_atual}'] = mes                   
            ws[f'W{linha_atual}'] = 5                     
            ws[f'Y{linha_atual}'] = valor_val             
            ws[f'AA{linha_atual}'] = valor_val            
            ws[f'AC{linha_atual}'] = 0                    
            ws[f'AE{linha_atual}'] = 0                    
            
        wb.save(caminho_saida_base)
        print(f" Planilha salva em: {caminho_saida_base}")


def buscar_arquivo_mais_recente(pasta=".dados", extensao="*.xls"):
    caminho_busca = os.path.join(pasta, extensao)
    arquivos = glob.glob(caminho_busca)
    
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo {extensao} foi encontrado na pasta '{pasta}'.")
        
    return max(arquivos, key=os.path.getmtime)


if __name__ == "__main__":
    pasta_dados = ".dados"
    arquivo_rh = os.path.join(pasta_dados, "Chapa colab e Dados dependentes.xlsx")
    arquivo_saida_auditoria = os.path.join(pasta_dados, "coparticipacao_consolidada_rh.xlsx")
    
    planilha_base_modelo = os.path.join(pasta_dados, "Planilha base.xlsx") 
    planilha_base_preenchida = os.path.join(pasta_dados, "planilha_base_pronta.xlsx")
    planilha_base_tratada = os.path.join(pasta_dados, "planilha_base_tratada.xlsx")
    
    try:
        arquivo_real = buscar_arquivo_mais_recente(pasta=pasta_dados, extensao="*.xls")
        print(f"🔍 Arquivo mais recente detectado automaticamente: '{arquivo_real}'\n")
        
        print("--- PARÂMETROS PARA PREENCHIMENTO DA PLANILHA BASE ---")
        ano_input = input("Informe o ANO da competência (ex: 2026): ").strip()
        mes_input = input("Informe o MÊS da competência (ex: 05 ou 5): ").strip().zfill(2)
        data_pagto_input = input("Informe a DATA DE PAGAMENTO no formato DDMMAAAA (ex: 30052026): ").strip()
        print("-----------------------------------------------------\n")
        
        print("Iniciando a varredura do arquivo...")
        extrator = ExtratorRelatorioConvenio(arquivo_real)
        extrator.processar_planilha()
        
        # 1. PLANILHA AUDITORIA E PLANILHA BASE PRONTA (Sem teto, com todas as despesas)
        print("Consolidando valores totais...")
        df_consolidado_total = extrator.consolidar_por_beneficiario(ignorar_internacoes=False)
        df_com_chapa_total = extrator.mesclar_dados_rh(df_consolidado_total, arquivo_rh)
        
        extrator.exportar_para_flat_table(df_com_chapa_total, arquivo_saida_auditoria)
        
        if os.path.exists(planilha_base_modelo):
            print("\n✍️ Gerando 'planilha_base_pronta.xlsx' (Completa, sem teto)...")
            extrator.preencher_planilha_base(
                df_com_chapa_total, 
                planilha_base_modelo, 
                planilha_base_preenchida,
                ano=ano_input,
                mes=mes_input,
                data_pagamento=data_pagto_input
            )
            
            # 2. PLANILHA BASE TRATADA (Sem Internações, Sem TIPO 'X' e Com Teto de R$ 150,00 por CHAPA)
            print("\nConsolidando valores para a base tratada...")
            df_consolidado_sem_internacao = extrator.consolidar_por_beneficiario(ignorar_internacoes=True)
            df_com_chapa_sem_internacao = extrator.mesclar_dados_rh(df_consolidado_sem_internacao, arquivo_rh)
            
            # Filtra cadastros marcados com 'X'
            df_tratado_temp = df_com_chapa_sem_internacao[df_com_chapa_sem_internacao['TIPO_CADASTRO'] != 'X'].copy()
            
            # Aplica a trava limite de R$ 150,00 por CHAPA
            df_tratado_final = extrator.aplicar_teto_por_chapa(df_tratado_temp, teto_maximo=150.0)
            
            print("\n✂️ Gerando 'planilha_base_tratada.xlsx' (Limitada a R$ 150,00 por CHAPA)...")
            extrator.preencher_planilha_base(
                df_tratado_final, 
                planilha_base_modelo, 
                planilha_base_tratada,
                ano=ano_input,
                mes=mes_input,
                data_pagamento=data_pagto_input
            )
        else:
            print(f"⚠️ AVISO: A planilha modelo '{planilha_base_modelo}' não foi encontrada na pasta .dados.")
        
    except Exception as e:
        print(f"\n❌ Erro durante a execução: {e}")