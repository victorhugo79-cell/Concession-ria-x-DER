"""
Gera um Excel com a consolidação por trecho (rodovia, km_int, sentido) a partir da planilha original.
Produz múltiplas planilhas com análises consolidadas e paretos.
"""

import pandas as pd
import numpy as np
import re
import shutil
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# -------- UTILITÁRIOS ----------

def format_excel_sheet(ws, is_header=True):
    """
    Formata uma planilha do Excel com:
    - Cabeçalho em negrito com fundo cinza
    - Ajuste automático de largura de colunas
    - Alinhamento centralizado no cabeçalho
    - Bordas nas células
    """
    # Estilos
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    # Formata cabeçalho
    if is_header and ws.max_row > 0:
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_alignment
            cell.border = border
    
    # Ajusta largura das colunas
    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        
        for cell in column:
            try:
                if cell.value:
                    # Calcula largura baseada no conteúdo
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            except:
                pass
        
        # Define largura mínima e máxima
        adjusted_width = min(max(max_length + 2, 10), 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Aplica bordas e alinhamento nas células de dados
    for row in ws.iter_rows(min_row=2 if is_header else 1, max_row=ws.max_row):
        for cell in row:
            cell.border = border
            # Alinha números à direita, texto à esquerda
            if isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(horizontal='right', vertical='center')
            else:
                cell.alignment = left_alignment
    
    # Congela primeira linha (cabeçalho)
    if is_header:
        ws.freeze_panes = 'A2'

def normalize_colname(s: str) -> str:
    """Normaliza nome de coluna para ascii lower_snake_case"""
    if pd.isna(s):
        return ""
    s = str(s).strip().lower()
    accents = {'á': 'a', 'ã': 'a', 'â': 'a', 'à': 'a', 'é': 'e', 'ê': 'e', 
               'í': 'i', 'ó': 'o', 'ô': 'o', 'õ': 'o', 'ú': 'u', 'ç': 'c', 
               'ü': 'u', 'ª': ''}
    for a, b in accents.items():
        s = s.replace(a, b)
    s = re.sub(r'[^0-9a-z]+', '_', s)
    s = re.sub(r'__+', '_', s).strip('_')
    return s

def find_col(df, keywords):
    """Procura a primeira coluna cujo nome contenha qualquer uma das keywords"""
    for k in keywords:
        for c in df.columns:
            if k in c.lower():
                return c
    return None

def detect_vehicle_flags(text):
    """Heurística simples para detectar presença de veículos em uma string"""
    t = "" if pd.isna(text) else str(text).lower()
    return {
        'moto': 1 if re.search(r'\bmoto\b|\bmotocicleta\b|\bmotoc\b', t) else 0,
        'carro': 1 if re.search(r'\bcarro\b|\bauto\b|\bautomovel\b|\bautomóvel\b|\bpassageiro\b', t) else 0,
        'caminhao': 1 if re.search(r'\bcaminh[aã]o\b|\bcarreta\b|\btruck\b|\btracto\b', t) else 0,
        'onibus': 1 if re.search(r'\bonibus\b|\bônibus\b|\bbus\b', t) else 0,
        'outro': 0
    }

def detect_acc_type(text):
    """Extrai tipo de acidente (heurística básica)"""
    t = "" if pd.isna(text) else str(text).lower()
    patterns = {
        'colisao_lateral': r'colis[aã]o lateral|colisao lateral|colis[aã]o.*lateral',
        'colisao_traseira': r'colis[aã]o traseira|traseira',
        'colisao_frontal': r'colis[aã]o frontal|frontal',
        'tombamento': r'tombamento|tombou',
        'capotamento': r'capotamento|capotou',
        'atropelamento': r'atropelamento|atropelou',
        'saida_de_pista': r'saida de pista|sa[íi]da de pista|saida_de_pista',
        'incendio': r'incendio|incêndio',
    }
    for name, pat in patterns.items():
        if re.search(pat, t):
            return name
    return 'outro'

def normalize_acc_type_group(acc_type):
    """
    Normaliza e agrupa tipos de acidentes similares.
    Agrupa todos os choques em "Choque", todas as quedas em "Queda", etc.
    """
    if pd.isna(acc_type):
        return 'Desconhecido'
    
    t = str(acc_type).strip().lower()
    
    # Primeiro verifica casos específicos (ordem importa - mais específicos primeiro)
    # Atropelamento de Animal (deve ser separado)
    if re.search(r'atropelamento.*animal|atropelou.*animal|atropelo.*animal|animal.*atropel|atrop\.?.*animal|animal.*atrop\.?', t):
        return 'Atropelamento de Animal'
    
    # Padrões de agrupamento (ordem importa - mais específicos primeiro)
    groups = {
        'Colisão': [
            r'colis[aã]o',
            r'colisao',
            r'colis\.',
        ],
        'Choque': [
            r'choque',
            r'batida',
            r'embate',
            r'impacto',
        ],
        'Queda': [
            r'queda',
            r'caiu',
            r'caida',
            r'caída',
        ],
        'Tombamento': [
            r'tombamento',
            r'tombou',
            r'virou',
            r'tomb\.',
        ],
        'Capotamento': [
            r'capotamento',
            r'capotou',
            r'capot\.',
        ],
        'Atropelamento': [
            r'atropelamento',
            r'atropelou',
            r'atropelo',
            r'atrop\.',  # Abreviação "Atrop."
            r'^atrop$',  # Apenas "atrop"
        ],
        'Saída de Pista': [
            r'sa[íi]da de pista',
            r'saida de pista',
            r'saída de pista',
            r'saiu da pista',
            r'fora da pista',
        ],
        'Incêndio': [
            r'inc[êe]ndio',
            r'incendio',
            r'queimou',
        ],
        'Abalroamento': [
            r'abalroamento',
            r'abalroou',
        ],
    }
    
    # Verifica cada grupo
    for group_name, patterns in groups.items():
        for pattern in patterns:
            if re.search(pattern, t):
                return group_name
    
    # Se não encontrou nenhum padrão conhecido, retorna o original capitalizado
    if t and t != 'nan' and t != 'none':
        # Capitaliza primeira letra
        return t.capitalize()
    
    return 'Outro'

# -------- SCRIPT PRINCIPAL ----------

def main():
    input_file = "L23_LESTE PAULISTA.xlsx"
    output_file = "Trechos_Consolidados.xlsx"
    sheet_name = "ACIDENTES__MITS"
    
    inp = Path(input_file)
    outp = Path(output_file)

    if not inp.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {inp}")

    # Lê o sheet especificado
    df_raw = pd.read_excel(inp, sheet_name=sheet_name, engine='openpyxl')

    # Normaliza colunas
    orig_cols = list(df_raw.columns)
    col_map = {c: normalize_colname(c) for c in orig_cols}
    df = df_raw.rename(columns=col_map).copy()

    # Detecta colunas importantes por heurística
    col_rodovia = find_col(df, ['rodovia', 'rodo'])
    col_km = find_col(df, ['km'])
    col_sentido = find_col(df, ['sentido'])
    # Prioriza "data_hora" sobre "data_hora_extracao"
    if 'data_hora' in df.columns:
        col_data = 'data_hora'
    else:
        col_data = find_col(df, ['data_hora', 'datahora', 'data'])
    
    # Para coluna de hora, procura especificamente por "hora" mas exclui "data_hora_extracao"
    col_hora = None
    if 'hora' in df.columns and 'hora' != col_data:
        col_hora = 'hora'
    else:
        # Procura por coluna que contenha "hora" mas não seja data_hora ou data_hora_extracao
        for col in df.columns:
            if 'hora' in col.lower() and col != col_data and 'extracao' not in col.lower():
                col_hora = col
                break
    col_veiculos = find_col(df, ['veiculos', 'veiculo', 'veiculos_envolvidos'])
    col_cond_climatica = find_col(df, ['cond', 'climaticas', 'climatica'])
    
    # Colunas de vítimas (tentando múltiplas variações)
    col_vit_grave = find_col(df, ['vit_grave', 'grave'])
    col_vit_leve = find_col(df, ['vit_leve', 'leve'])
    col_vit_moderada = find_col(df, ['vit_moderada', 'moderada'])
    col_vit_fatal = find_col(df, ['vit_fatal', 'fatal', 'obito', 'óbito', 'morte'])
    col_nro_oc = find_col(df, ['nro_oc', 'oc', 'id', 'numero'])
    col_tipo_acidente = find_col(df, ['tipo', 'classif', 'classificacao', 'natureza', 'tipo_do_acidente'])

    print("Colunas detectadas:")
    print(f"  rodovia: {col_rodovia}")
    print(f"  km: {col_km}")
    print(f"  sentido: {col_sentido}")
    print(f"  data: {col_data}, hora: {col_hora}")
    print(f"  veiculos: {col_veiculos}")
    print(f"  cond_climatica: {col_cond_climatica}")

    # Cria cópia de trabalho
    dfw = df.copy()

    # Parse datetime/horas
    if col_data and col_data in dfw.columns:
        try:
            dfw['__dt'] = pd.to_datetime(dfw[col_data], errors='coerce')
        except:
            dfw['__dt'] = pd.to_datetime(dfw[col_data].astype(str).str.replace('T', ' '), errors='coerce')
        
    else:
        dfw['__dt'] = pd.NaT
        print("  AVISO: Nenhuma coluna de data encontrada!")

    # Extrai hora da coluna de data/hora (sempre usa __dt se disponível)
    # Só usa coluna hora separada se for realmente uma coluna de hora (não data_hora_extracao)
    if col_hora and col_hora in dfw.columns and col_hora != col_data and col_hora != 'data_hora_extracao':
        try:
            # Tenta converter hora no formato HH:MM
            dfw['__hour'] = pd.to_datetime(dfw[col_hora], format="%H:%M", errors='coerce').dt.hour
            # Se a maioria dos valores for NaN, usa __dt em vez disso
            if dfw['__hour'].isna().sum() > len(dfw) * 0.5:
                print(f"  AVISO: Coluna '{col_hora}' não contém horas válidas. Usando hora de '{col_data}'.")
                dfw['__hour'] = dfw['__dt'].dt.hour
        except:
            try:
                dfw['__hour'] = pd.to_datetime(dfw[col_hora], errors='coerce').dt.hour
                if dfw['__hour'].isna().sum() > len(dfw) * 0.5:
                    dfw['__hour'] = dfw['__dt'].dt.hour
            except:
                dfw['__hour'] = dfw['__dt'].dt.hour
    else:
        # Extrai hora da coluna de data/hora
        dfw['__hour'] = dfw['__dt'].dt.hour
    

    # Período dia/noite (dia 06:00-17:59)
    def daynight(h):
        try:
            h = int(h)
        except:
            return 'unknown'
        return 'dia' if 6 <= h < 18 else 'noite'

    dfw['hora_int'] = dfw['__hour'].fillna(-1).astype(int)
    dfw['periodo'] = dfw['hora_int'].apply(lambda x: daynight(x) if x >= 0 else 'unknown')
    
    # Dia da semana (0=Segunda, 6=Domingo)
    # Nota: pandas dayofweek retorna 0=Segunda, 1=Terça, ..., 6=Domingo
    dfw['dia_semana_num'] = dfw['__dt'].dt.dayofweek
    dias_semana = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']
    dfw['dia_semana'] = dfw['dia_semana_num'].apply(
        lambda x: dias_semana[int(x)] if pd.notna(x) and 0 <= int(x) < 7 else 'Desconhecido'
    )
    
    
    # Mês e Ano (formato: MM/AAAA)
    dfw['mes_ano'] = dfw['__dt'].apply(
        lambda x: x.strftime('%m/%Y') if pd.notna(x) else 'Desconhecido'
    )

    # KM numérico e km_int (km arredondado para baixo)
    if col_km and col_km in dfw.columns:
        dfw['km_num'] = pd.to_numeric(dfw[col_km], errors='coerce')
    else:
        dfw['km_num'] = np.nan
    dfw['km_int'] = dfw['km_num'].apply(lambda x: int(np.floor(x)) if not pd.isna(x) else np.nan)

    # Vítimas numéricas
    dfw['num_vitimas'] = 0
    if col_vit_grave and col_vit_grave in dfw.columns:
        dfw['num_vitimas'] += pd.to_numeric(dfw[col_vit_grave], errors='coerce').fillna(0).astype(int)
    if col_vit_leve and col_vit_leve in dfw.columns:
        dfw['num_vitimas'] += pd.to_numeric(dfw[col_vit_leve], errors='coerce').fillna(0).astype(int)
    if col_vit_moderada and col_vit_moderada in dfw.columns:
        dfw['num_vitimas'] += pd.to_numeric(dfw[col_vit_moderada], errors='coerce').fillna(0).astype(int)

    dfw['num_fatais'] = 0
    if col_vit_fatal and col_vit_fatal in dfw.columns:
        dfw['num_fatais'] = pd.to_numeric(dfw[col_vit_fatal], errors='coerce').fillna(0).astype(int)

    # Flags de veículo
    for v in ['moto', 'carro', 'caminhao', 'onibus']:
        dfw[f'veh_{v}'] = 0
    
    if col_veiculos and col_veiculos in dfw.columns:
        for idx, val in dfw[col_veiculos].fillna('').items():
            flags = detect_vehicle_flags(val)
            dfw.at[idx, 'veh_moto'] = flags['moto']
            dfw.at[idx, 'veh_carro'] = flags['carro']
            dfw.at[idx, 'veh_caminhao'] = flags['caminhao']
            dfw.at[idx, 'veh_onibus'] = flags['onibus']

    # Tipo de acidente (normaliza e agrupa tipos similares)
    if col_tipo_acidente and col_tipo_acidente in dfw.columns:
        dfw['acc_type_norm'] = dfw[col_tipo_acidente].apply(normalize_acc_type_group)
    else:
        dfw['acc_type_norm'] = 'Desconhecido'

    # Identificar chave do trecho: rodovia + km_int + sentido
    if not col_rodovia or col_rodovia not in dfw.columns:
        raise RuntimeError("Não foi possível localizar coluna 'rodovia' automaticamente.")
    
    if not col_sentido or col_sentido not in dfw.columns:
        dfw['sentido'] = np.nan
        col_sentido = 'sentido'
    else:
        dfw['sentido'] = dfw[col_sentido]

    dfw['rodovia'] = dfw[col_rodovia].astype(str)

    # Substitui NaN de km_int por -1 para agrupar separados
    dfw['km_int_fill'] = dfw['km_int'].fillna(-1).astype(int)

    # --- AGREGAÇÕES POR TRECHO (rodovia, km_int_fill, sentido) ---
    group_cols = ['rodovia', 'km_int_fill', 'sentido']
    
    # Agregações básicas
    df_group = dfw.groupby(group_cols).agg(
        sinistros_total=pd.NamedAgg(column='rodovia', aggfunc='size'),
        total_vitimas=pd.NamedAgg(column='num_vitimas', aggfunc='sum'),
        total_fatais=pd.NamedAgg(column='num_fatais', aggfunc='sum'),
        motos=pd.NamedAgg(column='veh_moto', aggfunc='sum'),
        carros=pd.NamedAgg(column='veh_carro', aggfunc='sum'),
        caminhoes=pd.NamedAgg(column='veh_caminhao', aggfunc='sum'),
        onibus=pd.NamedAgg(column='veh_onibus', aggfunc='sum'),
    ).reset_index()

    # Dia/noite counts
    # Cria coluna auxiliar para contagem
    dfw['__count'] = 1
    periodo_counts = dfw.pivot_table(
        index=group_cols, 
        columns='periodo', 
        values='__count', 
        aggfunc='sum', 
        fill_value=0
    ).reset_index()
    
    # Garante que todas as colunas de período existam
    periodo_cols = ['dia', 'noite', 'unknown']
    for c in periodo_cols:
        if c not in periodo_counts.columns:
            periodo_counts[c] = 0

    df_group = df_group.merge(
        periodo_counts[group_cols + periodo_cols], 
        on=group_cols, how='left'
    )
    # Preenche NaN com 0
    for c in periodo_cols:
        if c in df_group.columns:
            df_group[c] = df_group[c].fillna(0).astype(int)

    # Dia da semana counts (exclui "Desconhecido")
    dfw_dia_semana = dfw[dfw['dia_semana'] != 'Desconhecido'].copy()
    dias_semana_cols = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']
    
    if len(dfw_dia_semana) > 0:
        # Usa pivot_table para garantir que todas as colunas sejam criadas
        dia_semana_counts = dfw_dia_semana.pivot_table(
            index=group_cols, 
            columns='dia_semana', 
            values='__count', 
            aggfunc='sum', 
            fill_value=0
        ).reset_index()
        
        # Garante que todas as colunas de dias da semana existam
        for dia in dias_semana_cols:
            if dia not in dia_semana_counts.columns:
                dia_semana_counts[dia] = 0
        
        df_group = df_group.merge(
            dia_semana_counts[group_cols + dias_semana_cols], 
            on=group_cols, how='left'
        )
        # Preenche NaN com 0
        for dia in dias_semana_cols:
            if dia in df_group.columns:
                df_group[dia] = df_group[dia].fillna(0).astype(int)
    else:
        # Se não houver dados válidos, adiciona colunas zeradas
        for dia in dias_semana_cols:
            df_group[dia] = 0

    # Mês/Ano: adiciona uma única coluna com o mês/ano mais frequente para cada trecho
    dfw_mes_ano = dfw[dfw['mes_ano'] != 'Desconhecido'].copy()
    if len(dfw_mes_ano) > 0:
        # Pega o mês/ano mais frequente para cada trecho
        mes_ano_mais_frequente = dfw_mes_ano.groupby(group_cols)['mes_ano'].apply(
            lambda x: x.mode()[0] if len(x.mode()) > 0 else x.iloc[0]
        ).reset_index(name='mes_ano')
        
        df_group = df_group.merge(
            mes_ano_mais_frequente[group_cols + ['mes_ano']], 
            on=group_cols, how='left'
        )
        # Preenche NaN com "Desconhecido"
        df_group['mes_ano'] = df_group['mes_ano'].fillna('Desconhecido')
    else:
        df_group['mes_ano'] = 'Desconhecido'

    # Condição climática (se disponível)
    if col_cond_climatica and col_cond_climatica in dfw.columns:
        cond_counts = dfw.groupby(group_cols + [col_cond_climatica]).size().unstack(fill_value=0).reset_index()
        
        # Normaliza nomes das colunas de condição climática para facilitar busca
        cond_cols_map = {}
        for col in cond_counts.columns:
            if col not in group_cols:
                col_upper = str(col).upper()
                for cond in ['BOA', 'NUBLADO', 'CHUVA', 'NEBLINA']:
                    if cond in col_upper:
                        cond_cols_map[col] = f'cond_{cond.lower()}'
                        break
        
        # Renomeia colunas encontradas
        cond_counts_renamed = cond_counts.rename(columns=cond_cols_map)
        
        # Adiciona colunas que não foram encontradas com valor 0
        for cond in ['BOA', 'NUBLADO', 'CHUVA', 'NEBLINA']:
            col_name = f'cond_{cond.lower()}'
            if col_name not in cond_counts_renamed.columns:
                cond_counts_renamed[col_name] = 0
        
        # Merge único com todas as colunas de condição climática
        cond_cols_to_merge = [c for c in cond_counts_renamed.columns if c.startswith('cond_')]
        if cond_cols_to_merge:
            df_group = df_group.merge(
                cond_counts_renamed[group_cols + cond_cols_to_merge], 
                on=group_cols, how='left'
            )
            # Preenche NaN com 0
            for col in cond_cols_to_merge:
                df_group[col] = df_group[col].fillna(0).astype(int)

    # Tipo de acidente: pivot para contagem por tipo
    acc_pivot = dfw.groupby(group_cols + ['acc_type_norm']).size().unstack(fill_value=0).reset_index()

    df_group = df_group.merge(acc_pivot, on=group_cols, how='left')

    # Cálculo ICR (Índice de Criticidade de Risco)
    # ICR = 10*fatais + 3*(vitimas - fatais) + 1*(ocorrencias - vitimas)
    df_group['ICR'] = (
        df_group['total_fatais'] * 10 + 
        (df_group['total_vitimas'] - df_group['total_fatais']) * 3 + 
        (df_group['sinistros_total'] - df_group['total_vitimas']) * 1
    )

    # Ordena por ICR decrescente
    df_group = df_group.sort_values('ICR', ascending=False)

    # Renomeia km_int_fill -> km_int
    df_group = df_group.rename(columns={'km_int_fill': 'km_int'})

    # Formata km_int para exibição (mantém como inteiro)
    df_group['km_int'] = df_group['km_int'].astype(int)
    # Se km_int for -1, substitui por NaN para melhor visualização
    df_group.loc[df_group['km_int'] == -1, 'km_int'] = np.nan

    # Folhas de paretos: veículos e acidentes (globais)
    pareto_veiculos = pd.DataFrame({
        'veiculo': ['moto', 'carro', 'caminhao', 'onibus'],
        'count': [
            int(dfw['veh_moto'].sum()),
            int(dfw['veh_carro'].sum()),
            int(dfw['veh_caminhao'].sum()),
            int(dfw['veh_onibus'].sum())
        ]
    }).sort_values('count', ascending=False)

    pareto_acidentes = dfw['acc_type_norm'].value_counts().reset_index()
    pareto_acidentes.columns = ['acc_type', 'count']

    resumo_periodo = dfw['periodo'].value_counts().reset_index()
    resumo_periodo.columns = ['periodo', 'count']

    # Verifica espaço em disco disponível
    free_gb = 0
    estimated_gb = 0
    try:
        total, used, free = shutil.disk_usage(outp.parent)
        free_gb = free / (1024**3)
        print(f"Espaço disponível no disco: {free_gb:.2f} GB")
        
        # Estima tamanho necessário (aproximadamente 2x o tamanho do arquivo original)
        if inp.exists():
            original_size = inp.stat().st_size
            estimated_size = original_size * 2
            estimated_gb = estimated_size / (1024**3)
            
            # Se não há espaço suficiente (menos de 10% de margem), aborta
            if free < estimated_size * 1.1:
                print(f"AVISO: Espaço disponível ({free_gb:.2f} GB) pode ser insuficiente.")
                print(f"Tamanho estimado necessário: {estimated_gb:.2f} GB")
                
                # Se realmente não há espaço (menos de 50MB), aborta imediatamente
                if free < 50 * 1024 * 1024:  # 50MB
                    raise RuntimeError(
                        f"ERRO CRÍTICO: Sem espaço em disco suficiente!\n"
                        f"Espaço disponível: {free_gb:.2f} GB\n"
                        f"Espaço necessário estimado: {estimated_gb:.2f} GB\n\n"
                        f"Soluções:\n"
                        f"  1. Libere espaço no disco (pelo menos {estimated_gb:.2f} GB)\n"
                        f"  2. Exclua arquivos temporários da pasta Temp\n"
                        f"  3. Execute o script em outro local com mais espaço\n"
                        f"  4. Limpe a pasta: {outp.parent}"
                    )
                
                print("Considerando remover planilhas grandes para economizar espaço...")
    except RuntimeError:
        raise  # Re-lança RuntimeError
    except Exception:
        pass  # Continua mesmo se não conseguir verificar espaço
    
    # Salva em Excel com várias sheets
    try:
        # Verifica se o arquivo está aberto tentando abrir em modo exclusivo
        if outp.exists():
            try:
                # Tenta abrir o arquivo para verificar se está em uso
                with open(outp, 'r+b') as f:
                    pass
            except PermissionError:
                raise PermissionError(
                    f"ERRO: O arquivo '{outp}' está aberto em outro programa (provavelmente Excel).\n"
                    "Por favor, feche o arquivo e execute o script novamente."
                )
        
        print("Salvando arquivo Excel...")
        with pd.ExcelWriter(outp, engine='openpyxl') as ew:
            # Trechos consolidados (planilha principal)
            print("  - Salvando trechos_consolidados...")
            df_group.to_excel(ew, sheet_name='trechos_consolidados', index=False)
            
            # Paretos e resumos (planilhas pequenas)
            print("  - Salvando paretos e resumos...")
            pareto_veiculos.to_excel(ew, sheet_name='pareto_veiculos', index=False)
            pareto_acidentes.to_excel(ew, sheet_name='pareto_acidentes', index=False)
            resumo_periodo.to_excel(ew, sheet_name='resumo_periodo', index=False)
            
            # Instruções
            print("  - Salvando instruções...")
            instr = [
                ("Objetivo", "Consolidação por (rodovia, km_int, sentido). Uma linha por trecho com indicadores."),
                ("ICR (fórmula)", "ICR = 10 * total_fatais + 3*(total_vitimas - total_fatais) + 1*(sinistros_total - total_vitimas)"),
                ("Periodização", "Dia = 06:00-17:59 ; Noite = 18:00-05:59"),
                ("KM Int", "KM arredondado para baixo (ex: 26.1 -> 26, 26.9 -> 26)"),
                ("Observações", "Detecção de veículos e tipos de acidente por heurística textual.")
            ]
            pd.DataFrame(instr, columns=['Item', 'Descrição']).to_excel(ew, sheet_name='instrucoes', index=False)
            
            # Planilhas grandes apenas se houver espaço suficiente
            try:
                print("  - Salvando raw_original...")
                df_raw.to_excel(ew, sheet_name='raw_original', index=False)
            except OSError as e:
                if "No space left" in str(e) or "28" in str(e):
                    print("  AVISO: Não foi possível salvar raw_original por falta de espaço. Continuando...")
                else:
                    raise
            
            # Raw normalizada removida para economizar espaço (pode ser recriada se necessário)

        # Aplica formatação ao arquivo Excel
        print("  - Aplicando formatação...")
        wb = load_workbook(outp)
        
        # Formata cada planilha
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            # Instruções não precisa de formatação de cabeçalho
            is_header = sheet_name != 'instrucoes'
            format_excel_sheet(ws, is_header=is_header)
        
        wb.save(outp)
        wb.close()
        
        print(f"\nProcessamento concluído. Arquivo salvo em: {outp.resolve()}")
        if outp.exists():
            file_size_mb = outp.stat().st_size / (1024**2)
            print(f"Tamanho do arquivo: {file_size_mb:.2f} MB")
    
    except PermissionError as e:
        print(f"\nERRO: {str(e)}")
        raise
    except (OSError, IOError) as e:
        error_str = str(e).lower()
        if any(keyword in error_str for keyword in ["no space", "28", "io_write", "disk full", "insufficient"]):
            print("\n" + "="*60)
            print("ERRO CRÍTICO: Sem espaço em disco!")
            print("="*60)
            print(f"Espaço disponível: {free_gb:.2f} GB")
            print(f"Espaço necessário estimado: {estimated_gb:.2f} GB")
            print("\nSoluções:")
            print("  1. Libere espaço no disco (pelo menos 100-200 MB)")
            print("  2. Exclua arquivos temporários da pasta Temp:")
            print(f"     {Path.home() / 'AppData' / 'Local' / 'Temp'}")
            print("  3. Execute o script em outro local com mais espaço")
            print(f"  4. Limpe a pasta de destino: {outp.parent}")
            if outp.exists():
                print(f"\n  5. O arquivo parcial '{outp}' pode estar corrompido - considere excluí-lo")
            print("="*60)
            raise RuntimeError("Falha ao salvar: sem espaço em disco") from e
        else:
            print(f"\nErro ao salvar arquivo: {e}")
            raise
    except Exception as e:
        error_str = str(e).lower()
        # Captura erros de serialização do lxml (que ocorrem quando não há espaço)
        if "serialisationerror" in error_str or "io_write" in error_str or "lxml" in error_str:
            print("\n" + "="*60)
            print("ERRO CRÍTICO: Falha ao escrever arquivo (provavelmente falta de espaço)")
            print("="*60)
            print(f"Espaço disponível: {free_gb:.2f} GB")
            print(f"Espaço necessário estimado: {estimated_gb:.2f} GB")
            print("\nSoluções:")
            print("  1. Libere espaço no disco (pelo menos 100-200 MB)")
            print("  2. Exclua arquivos temporários da pasta Temp:")
            print(f"     {Path.home() / 'AppData' / 'Local' / 'Temp'}")
            print("  3. Execute o script em outro local com mais espaço")
            if outp.exists():
                print(f"\n  4. O arquivo parcial '{outp}' pode estar corrompido - considere excluí-lo")
            print("="*60)
            raise RuntimeError("Falha ao salvar: erro de escrita (verifique espaço em disco)") from e
        else:
            print(f"\nErro ao salvar arquivo: {e}")
            print(f"Tipo de erro: {type(e).__name__}")
            raise

if __name__ == "__main__":
    main()
