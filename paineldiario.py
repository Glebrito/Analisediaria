import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
import calendar
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors as rl_colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import io

# Função para carregar credenciais (Streamlit Cloud ou local)
def get_google_credentials():
    """Carrega credenciais do Google - Streamlit Secrets ou arquivo local"""
    try:
        # Tentar carregar do Streamlit Secrets (produção)
        if hasattr(st, 'secrets') and 'gcp_service_account' in st.secrets:
            return Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=['https://spreadsheets.google.com/feeds',
                       'https://www.googleapis.com/auth/drive']
            )
    except:
        pass
    
    # Fallback: arquivo local (desenvolvimento)
    try:
        return Credentials.from_service_account_file(
            'bustling-day-459711-q8-e889589cda14.json',
            scopes=['https://spreadsheets.google.com/feeds',
                   'https://www.googleapis.com/auth/drive']
        )
    except Exception as e:
        st.error(f"❌ Erro ao carregar credenciais: {e}")
        st.info("💡 Configure os secrets no Streamlit Cloud ou adicione o arquivo JSON localmente")
        return None

# Dicionário de meses
meses = {
    1: "01 - Janeiro",
    2: "02 - Fevereiro",
    3: "03 - Março",
    4: "04 - Abril",
    5: "05 - Maio",
    6: "06 - Junho",
    7: "07 - Julho",
    8: "08 - Agosto",
    9: "09 - Setembro",
    10: "10 - Outubro",
    11: "11 - Novembro",
    12: "12 - Dezembro"
}

# Função para obter o número de dias no mês
def obter_dias_no_mes(mes, ano):
    """Retorna o número de dias no mês especificado"""
    return calendar.monthrange(ano, mes)[1]

# Função para ordenar tipos de vendedor na sequência desejada
def ordenar_tipos_vendedor(tipos):
    """Ordena os tipos de vendedor na sequência: Desks, Online, Transferistas, Guias"""
    ordem_desejada = ['Desks', 'Online', 'Transferistas', 'Guias']
    tipos_ordenados = []
    
    # Adicionar tipos na ordem específica se existirem nos dados
    for tipo_ordem in ordem_desejada:
        if tipo_ordem in tipos:
            tipos_ordenados.append(tipo_ordem)
    
    # Adicionar outros tipos que não estão na ordem específica (alfabeticamente)
    outros_tipos = sorted([tipo for tipo in tipos if tipo not in ordem_desejada])
    tipos_ordenados.extend(outros_tipos)
    
    return tipos_ordenados

# Função para conectar ao Google Sheets
@st.cache_data(ttl=300)
def carregar_dados_google_sheets():
    try:
        # Configurar credenciais
        creds = get_google_credentials()
        if not creds:
            return pd.DataFrame()
        
        client = gspread.authorize(creds)
        
        # Abrir a planilha
        sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1tkjltVrS_8SI4assF0Wlp2CwHhCZbmMt2Lfwi9BVuxY/edit?gid=753037688#gid=753037688')
        
        # Ler a aba Vendedores
        worksheet = sheet.worksheet('Vendedores')
        dados = worksheet.get_all_records()
        df = pd.DataFrame(dados)
        
        return df
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()

# Função para carregar dados de vendas da aba "Dados Finais Vendas"
@st.cache_data(ttl=300)
def carregar_dados_vendas():
    try:
        # Configurar credenciais
        creds = get_google_credentials()
        if not creds:
            return pd.DataFrame()
        
        client = gspread.authorize(creds)
        
        # Abrir a planilha
        sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1tkjltVrS_8SI4assF0Wlp2CwHhCZbmMt2Lfwi9BVuxY/edit?gid=1889628361#gid=1889628361')
        
        # Ler a aba Dados Finais Vendas
        worksheet = sheet.worksheet('Dados Finais Vendas')
        dados = worksheet.get_all_records()
        df_vendas = pd.DataFrame(dados)
        
        return df_vendas
    except Exception as e:
        st.error(f"Erro ao carregar dados de vendas: {e}")
        return pd.DataFrame()

# Função para carregar dados de Paxs In da aba "Dados In de Escala"
@st.cache_data(ttl=300)
def carregar_dados_paxs_in():
    try:
        # Configurar credenciais
        creds = get_google_credentials()
        if not creds:
            return pd.DataFrame()
        
        client = gspread.authorize(creds)
        
        # Abrir a planilha
        sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1tkjltVrS_8SI4assF0Wlp2CwHhCZbmMt2Lfwi9BVuxY/edit?gid=453466581#gid=453466581')
        
        # Ler a aba Dados In de Escala
        worksheet = sheet.worksheet('Dados In de Escala')
        dados = worksheet.get_all_records()
        df_paxs = pd.DataFrame(dados)
        
        return df_paxs
    except Exception as e:
        st.error(f"Erro ao carregar dados de Paxs In: {e}")
        return pd.DataFrame()

# Função para carregar serviços terceirizados
@st.cache_data(ttl=300)
def carregar_servicos_terceiros():
    """Carrega a lista de serviços terceirizados do Google Sheets"""
    try:
        # Configurar credenciais
        creds = get_google_credentials()
        if not creds:
            return []
        
        client = gspread.authorize(creds)
        
        # ID da planilha de valores finais onde estão os serviços terceirizados
        ALL_INCLUSIVE_SPREADSHEET_ID = '1tkjltVrS_8SI4assF0Wlp2CwHhCZbmMt2Lfwi9BVuxY'
        SERVICOS_TERCEIROS_GID = 1111997089  # ID da aba de serviços terceirizados
        
        # Abrir planilha e aba específica de serviços
        spreadsheet = client.open_by_key(ALL_INCLUSIVE_SPREADSHEET_ID)
        worksheet = spreadsheet.get_worksheet_by_id(SERVICOS_TERCEIROS_GID)
        
        if not worksheet:
            st.error("Não foi possível encontrar a aba de serviços terceirizados")
            return []
            
        # Pegar todos os valores da coluna Nome do Serviço
        try:
            valores = worksheet.get_all_values()
            if len(valores) > 1:  # Se tiver pelo menos cabeçalho e uma linha
                # Encontrar índice da coluna Nome do Serviço
                header = valores[0]
                servico_idx = header.index("Nome do Serviço") if "Nome do Serviço" in header else 0
                return [row[servico_idx] for row in valores[1:] if row and row[servico_idx]]
        except Exception as e:
            st.error(f"Erro ao ler serviços terceirizados: {e}")
            return []
            
        return []
            
    except Exception as e:
        st.error(f"Erro ao carregar serviços terceirizados: {e}")
        return []

# Função para carregar dados de vendedores (para comissão Luck)
@st.cache_data(ttl=300)
def carregar_dados_vendedores():
    """Carrega dados de vendedores da aba Dados Vendedores para buscar comissões"""
    try:
        # Configurar credenciais
        creds = get_google_credentials()
        if not creds:
            return pd.DataFrame()
        
        client = gspread.authorize(creds)
        
        # ID da planilha principal onde estão os dados de vendedores
        VENDEDORES_SPREADSHEET_ID = '1--dYU8SplKM8wdYtag2MzdggWujtsgAx2lRCmFBPlJs'
        
        # Abrir planilha e aba Dados Vendedores
        spreadsheet = client.open_by_key(VENDEDORES_SPREADSHEET_ID)
        worksheet = spreadsheet.worksheet('Dados Vendedores')
        
        # Pegar todos os dados
        dados = worksheet.get_all_records()
        df_vendedores = pd.DataFrame(dados)
        
        return df_vendedores
            
    except Exception as e:
        st.error(f"Erro ao carregar dados de vendedores: {e}")
        return pd.DataFrame()

# Função para carregar dados de Meta Diaria
@st.cache_data(ttl=300)
def carregar_dados_meta_diaria():
    """Carrega dados da aba Meta Diaria"""
    try:
        # Configurar credenciais
        creds = get_google_credentials()
        if not creds:
            return pd.DataFrame()
        
        client = gspread.authorize(creds)
        
        # Abrir a planilha
        sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1tkjltVrS_8SI4assF0Wlp2CwHhCZbmMt2Lfwi9BVuxY/edit?gid=2134646456#gid=2134646456')
        
        # Ler a aba Meta Diaria
        worksheet = sheet.worksheet('Meta Diaria')
        dados = worksheet.get_all_records()
        df_meta = pd.DataFrame(dados)
        
        return df_meta
    except Exception as e:
        st.error(f"Erro ao carregar dados de Meta Diaria: {e}")
        return pd.DataFrame()

# Função para normalizar nomes
def normalizar_nome(nome):
    """Normaliza nomes para comparação"""
    if pd.isna(nome):
        return ''
    
    import unicodedata
    import re
    
    # Converter para string e remover espaços extras
    nome = str(nome).strip()
    
    # Remover acentos
    nome = unicodedata.normalize('NFD', nome)
    nome = ''.join(char for char in nome if unicodedata.category(char) != 'Mn')
    
    # Converter para maiúsculas e remover caracteres especiais
    nome = re.sub(r'[^A-Za-z0-9\s]', '', nome.upper())
    
    # Remover espaços duplos
    nome = re.sub(r'\s+', ' ', nome).strip()
    
    return nome

# Função para buscar comissão Luck
def buscar_comissao_luck(vendedor, mes, ano, df_vendedores):
    """Busca comissão Luck baseada em vendedor, mês e ano"""
    try:
        if df_vendedores.empty:
            return ''
        
        # Normalizar dados de entrada
        vendedor_norm = normalizar_nome(vendedor)
        
        # DEBUG: Para vendedores específicos
        vendedores_debug = ['XAVIER', 'SENA', 'FLAVIA']
        debug_ativo = vendedor_norm.upper() in vendedores_debug
        
        # Converter mês para formato numérico com 2 dígitos
        meses_ordem = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                      'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
        
        def mes_para_str_num(mes):
            mes_str = str(mes).strip().lower()
            # Tentar converter nome do mês para número
            for idx, nome in enumerate(meses_ordem, 1):
                if mes_str == nome.lower():
                    return f'{idx:02d}'
            # Se já é número, formatar com 2 dígitos
            try:
                num = int(mes_str)
                return f'{num:02d}'
            except:
                pass
            if mes_str.isdigit():
                return mes_str.zfill(2)
            return mes_str
        
        mes_str = mes_para_str_num(mes)
        
        if debug_ativo:
            st.warning(f"🔍 DEBUG {vendedor_norm.upper()} - Buscando comissão Luck:")
            st.write(f"- Vendedor original: '{vendedor}' → normalizado: '{vendedor_norm}'")
            st.write(f"- Mês solicitado: {mes} (tipo: {type(mes)}) → convertido para: '{mes_str}'")
            st.write(f"- Ano: {ano} (tipo: {type(ano)})")
            st.write(f"- Colunas disponíveis: {list(df_vendedores.columns)}")
            
            # Mostrar exemplos de como está o mês na planilha
            if 'mês' in df_vendedores.columns:
                exemplos_mes = df_vendedores['mês'].head(10).tolist()
                tipos_mes = [type(x).__name__ for x in exemplos_mes]
                st.write(f"- Exemplos de valores na coluna 'mês': {exemplos_mes}")
                st.write(f"- Tipos dos valores: {tipos_mes}")
        
        # Verificar se as colunas necessárias existem
        colunas_necessarias = ['Vendedor', 'mês', 'Ano', 'Comissão Luck']
        colunas_existentes = [col for col in colunas_necessarias if col in df_vendedores.columns]
        
        if len(colunas_existentes) < 4:
            if debug_ativo:
                st.error(f"❌ Colunas faltando: {set(colunas_necessarias) - set(colunas_existentes)}")
            return ''
        
        # Filtrar dados - converter mês para int para comparação
        mes_int = int(mes_str)
        
        filtro = (
            (df_vendedores['Vendedor'].apply(normalizar_nome) == vendedor_norm) &
            (df_vendedores['mês'].apply(lambda x: int(x) if pd.notna(x) else 0) == mes_int) &
            (df_vendedores['Ano'] == ano)
        )
        
        if debug_ativo:
            vendedores_unicos = df_vendedores['Vendedor'].apply(normalizar_nome).unique()
            st.write(f"- Total de vendedores únicos: {len(vendedores_unicos)}")
            st.write(f"- Primeiros 30 vendedores (normalizados): {sorted(vendedores_unicos)[:30]}")
            st.write(f"- '{vendedor_norm}' está na lista? {vendedor_norm in vendedores_unicos}")
            st.write(f"- Total de linhas que atendem o filtro: {filtro.sum()}")
            
            # Verificar registros do vendedor (independente de mês/ano)
            filtro_vendedor_apenas = df_vendedores['Vendedor'].apply(normalizar_nome) == vendedor_norm
            st.write(f"- Total de registros deste vendedor (qualquer período): {filtro_vendedor_apenas.sum()}")
            
            if filtro_vendedor_apenas.sum() > 0:
                registros_vendedor = df_vendedores[filtro_vendedor_apenas][['Vendedor', 'mês', 'Ano', 'Comissão Luck']].head(10)
                st.write("- Alguns registros deste vendedor:")
                st.dataframe(registros_vendedor)
            
            if filtro.sum() > 0:
                registros_encontrados = df_vendedores[filtro][['Vendedor', 'mês', 'Ano', 'Comissão Luck']]
                st.write("- Registros encontrados para o período específico:")
                st.dataframe(registros_encontrados)
        
        resultado = df_vendedores.loc[filtro, 'Comissão Luck']
        
        if not resultado.empty:
            valor = resultado.iloc[0]
            # Retornar valor se não for nulo/vazio
            if pd.notna(valor) and str(valor).strip() != '':
                if debug_ativo:
                    st.success(f"✅ Comissão Luck encontrada: '{valor}'")
                return str(valor).strip()
            else:
                if debug_ativo:
                    st.warning(f"⚠️ Valor encontrado mas vazio/nulo: '{valor}'")
        else:
            if debug_ativo:
                st.error("❌ Nenhum resultado encontrado para o filtro")
        
        return ''
        
    except Exception as e:
        if debug_ativo:
            import traceback
            st.error(f"💥 Erro na busca: {str(e)}")
            st.code(traceback.format_exc())
        return f'Erro: {str(e)}'

# Função para buscar comissão Terceiros
def buscar_comissao_terceiros(vendedor, mes, ano, df_vendedores):
    """Busca comissão Terceiros baseada em vendedor, mês e ano"""
    try:
        if df_vendedores.empty:
            return ''
        
        # Normalizar dados de entrada
        vendedor_norm = normalizar_nome(vendedor)
        
        # DEBUG: Para vendedores específicos
        vendedores_debug = ['XAVIER', 'SENA', 'FLAVIA']
        debug_ativo = vendedor_norm.upper() in vendedores_debug
        
        # Converter mês para formato numérico com 2 dígitos
        meses_ordem = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                      'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
        
        def mes_para_str_num(mes):
            mes_str = str(mes).strip().lower()
            # Tentar converter nome do mês para número
            for idx, nome in enumerate(meses_ordem, 1):
                if mes_str == nome.lower():
                    return f'{idx:02d}'
            # Se já é número, formatar com 2 dígitos
            try:
                num = int(mes_str)
                return f'{num:02d}'
            except:
                pass
            if mes_str.isdigit():
                return mes_str.zfill(2)
            return mes_str
        
        mes_str = mes_para_str_num(mes)
        
        if debug_ativo:
            st.warning(f"🔍 DEBUG {vendedor_norm.upper()} - Buscando comissão Terceiros:")
            st.write(f"- Vendedor original: '{vendedor}' → normalizado: '{vendedor_norm}'")
            st.write(f"- Mês solicitado: {mes} (tipo: {type(mes)}) → convertido para: '{mes_str}'")
            st.write(f"- Ano: {ano} (tipo: {type(ano)})")
        
        # Verificar se as colunas necessárias existem
        colunas_necessarias = ['Vendedor', 'mês', 'Ano', 'Comissão Terceiros']
        colunas_existentes = [col for col in colunas_necessarias if col in df_vendedores.columns]
        
        if len(colunas_existentes) < 4:
            if debug_ativo:
                st.error(f"❌ Colunas faltando: {set(colunas_necessarias) - set(colunas_existentes)}")
            return ''
        
        # Filtrar dados - converter mês para int para comparação
        mes_int = int(mes_str)
        
        filtro = (
            (df_vendedores['Vendedor'].apply(normalizar_nome) == vendedor_norm) &
            (df_vendedores['mês'].apply(lambda x: int(x) if pd.notna(x) else 0) == mes_int) &
            (df_vendedores['Ano'] == ano)
        )
        
        if debug_ativo:
            st.write(f"- Total de linhas que atendem o filtro: {filtro.sum()}")
            
            # Verificar registros do vendedor (independente de mês/ano)
            filtro_vendedor_apenas = df_vendedores['Vendedor'].apply(normalizar_nome) == vendedor_norm
            st.write(f"- Total de registros deste vendedor (qualquer período): {filtro_vendedor_apenas.sum()}")
            
            if filtro_vendedor_apenas.sum() > 0:
                registros_vendedor = df_vendedores[filtro_vendedor_apenas][['Vendedor', 'mês', 'Ano', 'Comissão Terceiros']].head(10)
                st.write("- Alguns registros deste vendedor:")
                st.dataframe(registros_vendedor)
            
            if filtro.sum() > 0:
                registros_encontrados = df_vendedores[filtro][['Vendedor', 'mês', 'Ano', 'Comissão Terceiros']]
                st.write("- Registros encontrados:")
                st.dataframe(registros_encontrados)
        
        resultado = df_vendedores.loc[filtro, 'Comissão Terceiros']
        
        if not resultado.empty:
            valor = resultado.iloc[0]
            # Retornar valor se não for nulo/vazio
            if pd.notna(valor) and str(valor).strip() != '':
                if debug_ativo:
                    st.success(f"✅ Comissão Terceiros encontrada: '{valor}'")
                return str(valor).strip()
            else:
                if debug_ativo:
                    st.warning(f"⚠️ Valor encontrado mas vazio/nulo: '{valor}'")
        else:
            if debug_ativo:
                st.error("❌ Nenhum resultado encontrado para o filtro")
        
        return ''
        
    except Exception as e:
        if debug_ativo:
            import traceback
            st.error(f"💥 Erro na busca Terceiros: {str(e)}")
            st.code(traceback.format_exc())
        return f'Erro: {str(e)}'
        
        # Verificar se as colunas necessárias existem
        colunas_necessarias = ['Vendedor', 'mês', 'Ano', 'Comissão Terceiros']
        colunas_existentes = [col for col in colunas_necessarias if col in df_vendedores.columns]
        
        if len(colunas_existentes) < 4:
            return ''
        
        # Filtrar dados
        filtro = (
            (df_vendedores['Vendedor'].apply(normalizar_nome) == vendedor_norm) &
            (df_vendedores['mês'].apply(lambda x: str(x).zfill(2)) == mes_str) &
            (df_vendedores['Ano'] == ano)
        )
        
        resultado = df_vendedores.loc[filtro, 'Comissão Terceiros']
        
        if not resultado.empty:
            valor = resultado.iloc[0]
            # Retornar valor se não for nulo/vazio
            if pd.notna(valor) and str(valor).strip() != '':
                return str(valor).strip()
        
        return ''
        
    except Exception as e:
        return f'Erro: {str(e)}'

# Função auxiliar para buscar ticket médio real do vendedor (aproximação)
def calcular_ticket_medio_aproximado(vendedor, mes, ano, df_vendas_global=None):
    """Calcula uma aproximação do ticket médio do vendedor para o período"""
    try:
        # Se tiver dados de vendas disponíveis, usar para cálculo mais preciso
        if df_vendas_global is not None and not df_vendas_global.empty:
            vendedor_norm = normalizar_nome(vendedor)
            
            # Filtrar vendas do vendedor no período
            # Aqui seria necessário ter dados de vendas com data detalhada
            # Por ora, retorna valor padrão
            return 1000.0  # Valor padrão para teste
        else:
            return 1000.0  # Valor padrão quando não há dados
    except:
        return 1000.0

# Função para calcular premiação baseada no alcance da meta
def calcular_premiacao_transferista(alcance_meta_str):
    """Calcula a premiação baseada no alcance da meta para Transferistas"""
    try:
        if not alcance_meta_str or alcance_meta_str == '':
            return '0%'
        
        # Converter string de percentual para número
        # Formato esperado: "120,50%" ou "95,25%"
        alcance_limpo = str(alcance_meta_str).replace('%', '').replace(',', '.').strip()
        
        if not alcance_limpo:
            return '0%'
        
        alcance_num = float(alcance_limpo)
        
        # Aplicar lógica de premiação escalonada
        if alcance_num >= 150.0:
            return '5%'
        elif alcance_num >= 120.0:
            return '4%'
        elif alcance_num >= 100.0:
            return '3%'
        elif alcance_num >= 90.0:
            return '2%'
        elif alcance_num >= 80.0:
            return '1%'
        else:
            return '0%'
            
    except Exception as e:
        return '0%'

# ================== FUNÇÕES DE GERAÇÃO DE PDF ==================

def gerar_pdf_estatistico(vendedor, periodo_texto, dados_grid1, dados_grid2, dados_resumo):
    """Gera PDF com relatório estatístico do vendedor"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=15*mm, leftMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    elementos = []
    styles = getSampleStyleSheet()
    
    # Estilo customizado
    titulo_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=14,
        textColor=rl_colors.HexColor('#1f77b4'),
        spaceAfter=6,
        alignment=TA_CENTER
    )
    
    subtitulo_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Normal'],
        fontSize=9,
        textColor=rl_colors.grey,
        spaceAfter=8,
        alignment=TA_CENTER
    )
    
    secao_style = ParagraphStyle(
        'CustomSection',
        parent=styles['Heading2'],
        fontSize=11,
        textColor=rl_colors.HexColor('#2ca02c'),
        spaceAfter=4,
        spaceBefore=6
    )
    
    # Título
    elementos.append(Paragraph("RELATÓRIO ESTATÍSTICO", titulo_style))
    elementos.append(Paragraph(f"Vendedor: <b>{vendedor}</b> | Período: {periodo_texto}", subtitulo_style))
    elementos.append(Spacer(1, 4*mm))
    
    # Seção 1: Vendas Luck Sem Adicionais
    elementos.append(Paragraph("📊 VENDAS LUCK SEM ADICIONAIS", secao_style))
    
    if dados_grid1:
        data_grid1 = [
            ['Métrica', 'Valor'],
            ['Vendas Luck Sem Adicionais', dados_grid1.get('Vendas Luck Sem Adicionais', 'R$ 0,00')],
            ['Paxs In', dados_grid1.get('Paxs In', '0')],
            ['Ticket Médio', dados_grid1.get('Ticket Médio', 'R$ 0,00')],
            ['Meta', dados_grid1.get('Meta', 'R$ 0,00')],
            ['Alcance de Meta', dados_grid1.get('Alcance de Meta', '0,00%')],
        ]
        
        if 'Premiação' in dados_grid1:
            data_grid1.append(['Premiação', dados_grid1.get('Premiação', '0%')])
        
        table1 = Table(data_grid1, colWidths=[90*mm, 60*mm])
        table1.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), rl_colors.HexColor('#1f77b4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), rl_colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('BACKGROUND', (0, 1), (-1, -1), rl_colors.beige),
            ('GRID', (0, 0), (-1, -1), 0.5, rl_colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [rl_colors.white, rl_colors.lightgrey]),
        ]))
        elementos.append(table1)
    else:
        elementos.append(Paragraph("Sem dados disponíveis", styles['Normal']))
    
    elementos.append(Spacer(1, 4*mm))
    
    # Seção 2: All Inclusive
    elementos.append(Paragraph("📊 ALL INCLUSIVE", secao_style))
    
    if dados_grid2:
        data_grid2 = [
            ['Métrica', 'Valor'],
            ['Vendas Luck Sem Adicionais AI', dados_grid2.get('Vendas Luck Sem Adicionais All Inclusive', 'R$ 0,00')],
            ['Paxs In All Inclusive', dados_grid2.get('Paxs In All Inclusive', '0')],
            ['Ticket Médio All Inclusive', dados_grid2.get('Ticket Médio All Inclusive', 'R$ 0,00')],
            ['Meta All Inclusive', dados_grid2.get('Meta All Inclusive', 'R$ 0,00')],
            ['Alcance de Meta All Inclusive', dados_grid2.get('Alcance de Meta All Inclusive', '0,00%')],
        ]
        
        if 'Premiação All Inclusive' in dados_grid2:
            data_grid2.append(['Premiação All Inclusive', dados_grid2.get('Premiação All Inclusive', '0%')])
        
        table2 = Table(data_grid2, colWidths=[90*mm, 60*mm])
        table2.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), rl_colors.HexColor('#1f77b4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), rl_colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('BACKGROUND', (0, 1), (-1, -1), rl_colors.beige),
            ('GRID', (0, 0), (-1, -1), 0.5, rl_colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [rl_colors.white, rl_colors.lightgrey]),
        ]))
        elementos.append(table2)
    else:
        elementos.append(Paragraph("Sem dados disponíveis", styles['Normal']))
    
    elementos.append(Spacer(1, 4*mm))
    
    # Seção 3: Resumo de Comissão
    elementos.append(Paragraph("💰 RESUMO DE COMISSÃO", secao_style))
    
    if dados_resumo:
        data_resumo = [
            ['Descrição', 'Valor'],
            ['Valor Total de Venda', dados_resumo.get('Valor Total de Venda', 'R$ 0,00')],
            ['Valor Total Comissão Luck', dados_resumo.get('Valor Total Comissão Luck', 'R$ 0,00')],
            ['Valor Total Comissão Terceiros', dados_resumo.get('Valor Total Comissão Terceiros', 'R$ 0,00')],
        ]
        
        if 'Valor Total Comissão Premiação' in dados_resumo:
            data_resumo.append(['Valor Total Comissão Premiação', dados_resumo.get('Valor Total Comissão Premiação', 'R$ 0,00')])
        
        if 'Valor Total Comissão Premiação All Inclusive' in dados_resumo:
            data_resumo.append(['Valor Total Comissão Premiação AI', dados_resumo.get('Valor Total Comissão Premiação All Inclusive', 'R$ 0,00')])
        
        data_resumo.append(['VALOR TOTAL DE COMISSÃO', dados_resumo.get('Valor Total de Comissão', 'R$ 0,00')])
        
        table3 = Table(data_resumo, colWidths=[90*mm, 60*mm])
        table3.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), rl_colors.HexColor('#2ca02c')),
            ('TEXTCOLOR', (0, 0), (-1, 0), rl_colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('BACKGROUND', (0, 1), (-1, -1), rl_colors.beige),
            ('GRID', (0, 0), (-1, -1), 0.5, rl_colors.black),
            ('FONTSIZE', (0, 1), (-1, -2), 8),
            ('ROWBACKGROUNDS', (0, 1), (-1, -2), [rl_colors.white, rl_colors.lightgrey]),
            ('BACKGROUND', (0, -1), (-1, -1), rl_colors.HexColor('#2ca02c')),
            ('TEXTCOLOR', (0, -1), (-1, -1), rl_colors.whitesmoke),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, -1), (-1, -1), 10),
        ]))
        elementos.append(table3)
    else:
        elementos.append(Paragraph("Sem dados disponíveis", styles['Normal']))
    
    # Construir PDF
    doc.build(elementos)
    buffer.seek(0)
    return buffer

def gerar_pdf_comissao(vendedor, periodo_texto, dados_detalhes, dados_resumo, tipo_vendedor):
    """Gera PDF com relatório de comissão detalhado do vendedor"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=10*mm, leftMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    elementos = []
    styles = getSampleStyleSheet()
    
    # Estilos
    titulo_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=14,
        textColor=rl_colors.HexColor('#1f77b4'),
        spaceAfter=4,
        alignment=TA_CENTER
    )
    
    subtitulo_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Normal'],
        fontSize=9,
        textColor=rl_colors.grey,
        spaceAfter=6,
        alignment=TA_CENTER
    )
    
    secao_style = ParagraphStyle(
        'CustomSection',
        parent=styles['Heading2'],
        fontSize=11,
        textColor=rl_colors.HexColor('#d62728'),
        spaceAfter=4,
        spaceBefore=4
    )
    
    # Título
    elementos.append(Paragraph("RELATÓRIO DE COMISSÃO", titulo_style))
    elementos.append(Paragraph(f"Vendedor: <b>{vendedor}</b> | Período: {periodo_texto}", subtitulo_style))
    elementos.append(Spacer(1, 2*mm))
    
    # Detalhes das Vendas
    elementos.append(Paragraph("📋 DETALHES DAS VENDAS", secao_style))
    
    if dados_detalhes is not None and not dados_detalhes.empty:
        # Calcular largura disponível
        largura_disponivel = landscape(A4)[0] - 20*mm  # 277mm disponível
        
        # Preparar cabeçalhos com larguras específicas otimizadas para caber na página
        colunas_config = [
            ('Data da Venda', 18*mm),
            ('Código da Reserva', 22*mm),
            ('Serviço', 38*mm),
            ('Valor da Venda', 18*mm),
            ('Venda All Inclusive', 15*mm),
            ('Tipo de Serviço', 18*mm),
            ('Comissão Luck', 15*mm),
            ('Comissão Terceiros', 18*mm),
        ]
        
        # Adicionar colunas de premiação apenas para Transferistas
        if tipo_vendedor == 'Transferistas':
            colunas_config.extend([
                ('Premiação', 15*mm),
                ('Premiação All Inclusive', 18*mm),
            ])
        
        colunas_config.extend([
            ('Valor Comissão Luck', 20*mm),
            ('Valor Comissão Terceiros', 22*mm),
        ])
        
        if tipo_vendedor == 'Transferistas':
            colunas_config.extend([
                ('Valor Comissão Premiação', 22*mm),
                ('Valor Comissão Premiação All Inclusive', 25*mm),
            ])
        
        colunas_config.append(('Valor Total de Comissão', 22*mm))
        
        # Filtrar apenas colunas que existem
        colunas_existentes = []
        col_widths = []
        for col_nome, col_largura in colunas_config:
            if col_nome in dados_detalhes.columns:
                colunas_existentes.append(col_nome)
                col_widths.append(col_largura)
        
        # Ajustar larguras proporcionalmente se exceder a largura disponível
        largura_total = sum(col_widths)
        if largura_total > largura_disponivel:
            fator_ajuste = largura_disponivel / largura_total
            col_widths = [w * fator_ajuste for w in col_widths]
        
        # Criar estilo para células com quebra de texto
        cell_style = ParagraphStyle(
            'CellStyle',
            fontSize=6,
            leading=7,
            wordWrap='CJK',
            alignment=TA_CENTER
        )
        
        # Processar dados em lotes de 25 linhas por página
        total_registros = len(dados_detalhes)
        linhas_por_pagina = 25
        
        for inicio in range(0, total_registros, linhas_por_pagina):
            fim = min(inicio + linhas_por_pagina, total_registros)
            
            # Se não for a primeira página, adicionar quebra
            if inicio > 0:
                elementos.append(PageBreak())
                elementos.append(Paragraph("📋 DETALHES DAS VENDAS (continuação)", secao_style))
            
            # Criar cabeçalho da tabela com quebra de linha
            header_row = []
            for col in colunas_existentes:
                header_row.append(Paragraph(f"<b>{col}</b>", 
                    ParagraphStyle('HeaderStyle', fontSize=6, leading=7, wordWrap='CJK', 
                                 alignment=TA_CENTER, textColor=rl_colors.whitesmoke)))
            data_table = [header_row]
            
            # Adicionar linhas deste lote
            for idx, row in dados_detalhes.iloc[inicio:fim].iterrows():
                linha = []
                for col in colunas_existentes:
                    valor = str(row.get(col, ''))
                    # Sempre usar Paragraph para garantir quebra de texto
                    linha.append(Paragraph(valor, cell_style))
                data_table.append(linha)
            
            # Criar tabela
            table_detalhes = Table(data_table, colWidths=col_widths, repeatRows=1)
            table_detalhes.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), rl_colors.HexColor('#d62728')),
                ('TEXTCOLOR', (0, 0), (-1, 0), rl_colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 6),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
                ('TOPPADDING', (0, 0), (-1, 0), 4),
                ('BACKGROUND', (0, 1), (-1, -1), rl_colors.beige),
                ('GRID', (0, 0), (-1, -1), 0.5, rl_colors.grey),
                ('FONTSIZE', (0, 1), (-1, -1), 6),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [rl_colors.white, rl_colors.lightgrey]),
                ('LEFTPADDING', (0, 0), (-1, -1), 2),
                ('RIGHTPADDING', (0, 0), (-1, -1), 2),
                ('TOPPADDING', (0, 1), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 3),
            ]))
            elementos.append(table_detalhes)
            elementos.append(Spacer(1, 3*mm))
        
        # Informação sobre total de registros
        if total_registros > linhas_por_pagina:
            elementos.append(Paragraph(f"<i>Total de {total_registros} registros exibidos</i>", 
                                      ParagraphStyle('Italic', fontSize=8, textColor=rl_colors.grey)))
    else:
        elementos.append(Paragraph("Sem dados de comissão disponíveis", styles['Normal']))
    
    # Criar elementos do resumo para KeepTogether
    elementos_resumo = []
    elementos_resumo.append(Spacer(1, 8*mm))
    elementos_resumo.append(Paragraph("💰 RESUMO FINAL DE COMISSÃO", secao_style))
    elementos_resumo.append(Spacer(1, 3*mm))
    
    if dados_resumo:
        data_resumo = [
            ['Descrição', 'Valor'],
            ['Valor Total de Venda', dados_resumo.get('Valor Total de Venda', 'R$ 0,00')],
            ['Valor Total Comissão Luck', dados_resumo.get('Valor Total Comissão Luck', 'R$ 0,00')],
            ['Valor Total Comissão Terceiros', dados_resumo.get('Valor Total Comissão Terceiros', 'R$ 0,00')],
        ]
        
        if tipo_vendedor == 'Transferistas':
            if 'Valor Total Comissão Premiação' in dados_resumo:
                data_resumo.append(['Valor Total Comissão Premiação', dados_resumo.get('Valor Total Comissão Premiação', 'R$ 0,00')])
            
            if 'Valor Total Comissão Premiação All Inclusive' in dados_resumo:
                data_resumo.append(['Valor Total Comissão Premiação AI', dados_resumo.get('Valor Total Comissão Premiação All Inclusive', 'R$ 0,00')])
        
        data_resumo.append(['VALOR TOTAL DE COMISSÃO', dados_resumo.get('Valor Total de Comissão', 'R$ 0,00')])
        
        table_resumo = Table(data_resumo, colWidths=[120*mm, 80*mm])
        table_resumo.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), rl_colors.HexColor('#d62728')),
            ('TEXTCOLOR', (0, 0), (-1, 0), rl_colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 14),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), rl_colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, rl_colors.black),
            ('FONTSIZE', (0, 1), (-1, -2), 12),
            ('ROWBACKGROUNDS', (0, 1), (-1, -2), [rl_colors.white, rl_colors.lightgrey]),
            ('BACKGROUND', (0, -1), (-1, -1), rl_colors.HexColor('#d62728')),
            ('TEXTCOLOR', (0, -1), (-1, -1), rl_colors.whitesmoke),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, -1), (-1, -1), 16),
        ]))
        elementos_resumo.append(table_resumo)
    else:
        elementos_resumo.append(Paragraph("Sem dados de resumo disponíveis", styles['Normal']))
    
    # Usar KeepTogether para manter o resumo na mesma página
    from reportlab.platypus import KeepTogether
    elementos.append(KeepTogether(elementos_resumo))
    
    # Construir PDF
    doc.build(elementos)
    buffer.seek(0)
    return buffer

# ================== FIM DAS FUNÇÕES DE PDF ==================

# Função para carregar dados da aba "Comissão"
@st.cache_data(ttl=300)
def carregar_dados_comissao():
    try:
        # Configurar credenciais
        creds = get_google_credentials()
        if not creds:
            return pd.DataFrame()
        
        client = gspread.authorize(creds)
        
        # Abrir a planilha da aba Comissão
        sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1--dYU8SplKM8wdYtag2MzdggWujtsgAx2lRCmFBPlJs/edit?gid=1274112565#gid=1274112565')
        
        # Ler a aba Comissão
        worksheet = sheet.worksheet('Comissão')
        dados = worksheet.get_all_records()
        df_comissao = pd.DataFrame(dados)
        
        return df_comissao
    except Exception as e:
        st.error(f"Erro ao carregar dados da Comissão: {e}")
        return pd.DataFrame()

# Função para calcular Vendas Luck Sem Adicionais por vendedor
def calcular_vendas_luck_sem_adicionais(df_vendas, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final):
    """
    Calcula as vendas Luck Sem Adicionais para uma lista de vendedores no período especificado
    NOTA: Filtro por dia não é usado pois a coluna 'dia' contém códigos, não dias reais
    """
    try:
        if df_vendas.empty:
            return {}
        
        # Verificar se existem as colunas necessárias
        colunas_alternativas = {
            'dia': ['dia', 'Dia', 'DIA'],
            'mês': ['mês', 'Mês', 'MES', 'Mes'],
            'ano': ['ano', 'Ano', 'ANO'],
            'vendedor': ['Vendedor', 'vendedor', 'VENDEDOR'],
            'valor': ['Valor Real', 'valor real', 'VALOR REAL'],
            'tipo_servico': ['Tipo de Serviço', 'tipo de serviço', 'TIPO DE SERVIÇO', 'Tipo de Servico', 'Serviço Buggy', 'Servico Buggy'],
            'all_inclusive': ['All Inclusive', 'all inclusive', 'ALL INCLUSIVE']
        }
        
        # Mapear colunas existentes
        colunas_mapeadas = {}
        for campo, alternativas in colunas_alternativas.items():
            for alt in alternativas:
                if alt in df_vendas.columns:
                    colunas_mapeadas[campo] = alt
                    break
        
        # Verificar se todas as colunas foram encontradas
        if len(colunas_mapeadas) < 7:
            return {}
        
        # Converter colunas para tipos adequados
        
        # DIA e ANO são números
        df_vendas[colunas_mapeadas['dia']] = pd.to_numeric(df_vendas[colunas_mapeadas['dia']], errors='coerce')
        df_vendas[colunas_mapeadas['ano']] = pd.to_numeric(df_vendas[colunas_mapeadas['ano']], errors='coerce')
        
        # VALOR REAL - limpar formatação antes de converter
        # Limpar formatação de moeda brasileira
        df_vendas['valor_limpo'] = df_vendas[colunas_mapeadas['valor']].astype(str)
        df_vendas['valor_limpo'] = df_vendas['valor_limpo'].str.replace('R$', '', regex=False)
        df_vendas['valor_limpo'] = df_vendas['valor_limpo'].str.replace('.', '', regex=False)  # Remove pontos de milhares
        df_vendas['valor_limpo'] = df_vendas['valor_limpo'].str.replace(',', '.', regex=False)  # Converte vírgula decimal para ponto
        df_vendas['valor_limpo'] = df_vendas['valor_limpo'].str.strip()
        df_vendas['valor_limpo'] = pd.to_numeric(df_vendas['valor_limpo'], errors='coerce').fillna(0)
        
        # Atualizar referência da coluna valor
        colunas_mapeadas['valor'] = 'valor_limpo'
        
        # MÊS é STRING - converter nomes de meses para números
        meses_para_numeros = {
            'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4, 
            'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8,
            'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12,
            # Versões sem acento também
            'Marco': 3, 'Decembro': 12
        }
        
        # Converter meses de string para número
        df_vendas['mes_numero'] = df_vendas[colunas_mapeadas['mês']].map(meses_para_numeros)
        
        # Filtrar por período usando dia, mês e ano separadamente
        # Filtrar por período usando dia (número), mês (convertido para número) e ano (número)
        mask_periodo = (
            (df_vendas[colunas_mapeadas['ano']] >= ano_inicial) & 
            (df_vendas[colunas_mapeadas['ano']] <= ano_final) &
            (df_vendas['mes_numero'] >= mes_inicial) & 
            (df_vendas['mes_numero'] <= mes_final) &
            (df_vendas[colunas_mapeadas['dia']] >= dia_inicial) & 
            (df_vendas[colunas_mapeadas['dia']] <= dia_final)
        )
        
        df_periodo = df_vendas[mask_periodo]
        
        # Ajustar filtros baseado nos dados reais
        # A coluna "Tipo de Serviço" tem valores "Luck" e "Terceiro"
        mask_luck = (
            (df_periodo[colunas_mapeadas['tipo_servico']] == 'Luck') & 
            (df_periodo[colunas_mapeadas['all_inclusive']] == 'Não')
        )
        
        df_luck = df_periodo[mask_luck]
        
        # Calcular soma por vendedor
        vendas_por_vendedor = {}
        
        for vendedor in vendedores_list:
            vendedor_dados = df_luck[df_luck[colunas_mapeadas['vendedor']] == vendedor]
            soma = vendedor_dados[colunas_mapeadas['valor']].sum()
            vendas_por_vendedor[vendedor] = soma
        
        return vendas_por_vendedor
        
    except Exception as e:
        st.error(f"Erro ao calcular vendas Luck sem adicionais: {e}")
        return {}

# Função para calcular Vendas Luck Sem Adicionais All Inclusive por vendedor
def calcular_vendas_luck_all_inclusive(df_vendas, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final):
    """
    Calcula as vendas Luck COM All Inclusive para uma lista de vendedores no período especificado
    Mesma lógica da função anterior, mas com All Inclusive = "Sim"
    """
    try:
        if df_vendas.empty:
            return {}
        
        # Verificar se existem as colunas necessárias
        colunas_alternativas = {
            'dia': ['dia', 'Dia', 'DIA'],
            'mês': ['mês', 'Mês', 'MES', 'Mes'],
            'ano': ['ano', 'Ano', 'ANO'],
            'vendedor': ['Vendedor', 'vendedor', 'VENDEDOR'],
            'valor': ['Valor Real', 'valor real', 'VALOR REAL'],
            'tipo_servico': ['Tipo de Serviço', 'tipo de serviço', 'TIPO DE SERVIÇO', 'Tipo de Servico', 'Serviço Buggy', 'Servico Buggy'],
            'all_inclusive': ['All Inclusive', 'all inclusive', 'ALL INCLUSIVE']
        }
        
        # Mapear colunas existentes
        colunas_mapeadas = {}
        for campo, alternativas in colunas_alternativas.items():
            for alt in alternativas:
                if alt in df_vendas.columns:
                    colunas_mapeadas[campo] = alt
                    break
        
        # Verificar se todas as colunas foram encontradas
        if len(colunas_mapeadas) < 7:
            return {}
        
        # Converter colunas para tipos adequados
        df_vendas[colunas_mapeadas['dia']] = pd.to_numeric(df_vendas[colunas_mapeadas['dia']], errors='coerce')
        df_vendas[colunas_mapeadas['ano']] = pd.to_numeric(df_vendas[colunas_mapeadas['ano']], errors='coerce')
        
        # VALOR REAL - limpar formatação antes de converter
        df_vendas['valor_limpo_ai'] = df_vendas[colunas_mapeadas['valor']].astype(str)
        df_vendas['valor_limpo_ai'] = df_vendas['valor_limpo_ai'].str.replace('R$', '', regex=False)
        df_vendas['valor_limpo_ai'] = df_vendas['valor_limpo_ai'].str.replace('.', '', regex=False)
        df_vendas['valor_limpo_ai'] = df_vendas['valor_limpo_ai'].str.replace(',', '.', regex=False)
        df_vendas['valor_limpo_ai'] = df_vendas['valor_limpo_ai'].str.strip()
        df_vendas['valor_limpo_ai'] = pd.to_numeric(df_vendas['valor_limpo_ai'], errors='coerce').fillna(0)
        
        # MÊS - converter nomes de meses para números
        meses_para_numeros = {
            'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4, 
            'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8,
            'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12,
            'Marco': 3, 'Decembro': 12
        }
        
        df_vendas['mes_numero_ai'] = df_vendas[colunas_mapeadas['mês']].map(meses_para_numeros)
        
        # Filtrar por período
        mask_periodo = (
            (df_vendas[colunas_mapeadas['ano']] >= ano_inicial) & 
            (df_vendas[colunas_mapeadas['ano']] <= ano_final) &
            (df_vendas['mes_numero_ai'] >= mes_inicial) & 
            (df_vendas['mes_numero_ai'] <= mes_final) &
            (df_vendas[colunas_mapeadas['dia']] >= dia_inicial) & 
            (df_vendas[colunas_mapeadas['dia']] <= dia_final)
        )
        
        df_periodo = df_vendas[mask_periodo]
        
        # DIFERENÇA: Filtrar por Luck + All Inclusive = "Sim"
        mask_luck_ai = (
            (df_periodo[colunas_mapeadas['tipo_servico']] == 'Luck') & 
            (df_periodo[colunas_mapeadas['all_inclusive']] == 'Sim')
        )
        
        df_luck_ai = df_periodo[mask_luck_ai]
        
        # Calcular soma por vendedor
        vendas_por_vendedor = {}
        
        for vendedor in vendedores_list:
            vendedor_dados = df_luck_ai[df_luck_ai[colunas_mapeadas['vendedor']] == vendedor]
            soma = vendedor_dados['valor_limpo_ai'].sum()
            vendas_por_vendedor[vendedor] = soma
        
        return vendas_por_vendedor
        
    except Exception as e:
        st.error(f"Erro ao calcular vendas Luck All Inclusive: {e}")
        return {}

# Função para calcular Vendas Luck para Online e Desks
def calcular_vendas_luck_online_desks(df_vendas, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final):
    """
    Calcula soma de Valor Final onde Tipo de Serviço = "Luck" para vendedores Online e Desks
    """
    try:
        if df_vendas.empty:
            return {}
        
        # Mapear colunas da aba "Dados Finais Vendas"
        colunas_alternativas = {
            'dia': ['dia', 'Dia', 'DIA'],
            'mês': ['mês', 'Mês', 'MES', 'Mes'],
            'ano': ['ano', 'Ano', 'ANO'],
            'vendedor': ['Vendedor', 'vendedor', 'VENDEDOR'],
            'valor_final': ['Valor Final', 'valor final', 'VALOR FINAL'],
            'tipo_servico': ['Tipo de Serviço', 'tipo de serviço', 'TIPO DE SERVIÇO', 'Tipo de Servico']
        }
        
        # Mapear colunas existentes
        colunas_mapeadas = {}
        for campo, alternativas in colunas_alternativas.items():
            for alt in alternativas:
                if alt in df_vendas.columns:
                    colunas_mapeadas[campo] = alt
                    break
        
        # Verificar se todas as colunas necessárias existem
        if len(colunas_mapeadas) < 6:
            return {}
        
        # Converter colunas para tipos adequados
        df_vendas[colunas_mapeadas['dia']] = pd.to_numeric(df_vendas[colunas_mapeadas['dia']], errors='coerce')
        df_vendas[colunas_mapeadas['ano']] = pd.to_numeric(df_vendas[colunas_mapeadas['ano']], errors='coerce')
        
        # Limpar e converter Valor Final
        df_vendas['valor_final_limpo'] = df_vendas[colunas_mapeadas['valor_final']].astype(str)
        df_vendas['valor_final_limpo'] = df_vendas['valor_final_limpo'].str.replace('R$', '', regex=False)
        df_vendas['valor_final_limpo'] = df_vendas['valor_final_limpo'].str.replace('.', '', regex=False)
        df_vendas['valor_final_limpo'] = df_vendas['valor_final_limpo'].str.replace(',', '.', regex=False)
        df_vendas['valor_final_limpo'] = df_vendas['valor_final_limpo'].str.strip()
        df_vendas['valor_final_limpo'] = pd.to_numeric(df_vendas['valor_final_limpo'], errors='coerce').fillna(0)
        
        # Converter meses de string para número
        meses_para_numeros = {
            'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4, 
            'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8,
            'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12,
            'Marco': 3, 'Decembro': 12
        }
        
        df_vendas['mes_numero'] = df_vendas[colunas_mapeadas['mês']].map(meses_para_numeros)
        
        # Filtrar por período
        mask_periodo = (
            (df_vendas[colunas_mapeadas['ano']] >= ano_inicial) & 
            (df_vendas[colunas_mapeadas['ano']] <= ano_final) &
            (df_vendas['mes_numero'] >= mes_inicial) & 
            (df_vendas['mes_numero'] <= mes_final) &
            (df_vendas[colunas_mapeadas['dia']] >= dia_inicial) & 
            (df_vendas[colunas_mapeadas['dia']] <= dia_final)
        )
        
        df_periodo = df_vendas[mask_periodo]
        
        # Filtrar apenas Tipo de Serviço = "Luck"
        mask_luck = (df_periodo[colunas_mapeadas['tipo_servico']] == 'Luck')
        df_luck = df_periodo[mask_luck]
        
        # Normalizar vendedor para comparação case-insensitive
        df_luck['vendedor_normalizado'] = df_luck[colunas_mapeadas['vendedor']].str.strip().str.upper()
        
        # Calcular soma por vendedor
        vendas_por_vendedor = {}
        
        for vendedor in vendedores_list:
            vendedor_normalizado = vendedor.strip().upper()
            vendedor_dados = df_luck[df_luck['vendedor_normalizado'] == vendedor_normalizado]
            soma = vendedor_dados['valor_final_limpo'].sum()
            vendas_por_vendedor[vendedor] = soma
        
        return vendas_por_vendedor
        
    except Exception as e:
        st.error(f"Erro ao calcular Vendas Luck para Online/Desks: {e}")
        return {}

# Função para calcular Vendas Terceiros para Online e Desks
def calcular_vendas_terceiros_online_desks(df_vendas, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final):
    """
    Calcula soma de Valor Final onde Tipo de Serviço = "Terceiro" para vendedores Online e Desks
    """
    try:
        if df_vendas.empty:
            return {}
        
        # Mapear colunas da aba "Dados Finais Vendas"
        colunas_alternativas = {
            'dia': ['dia', 'Dia', 'DIA'],
            'mês': ['mês', 'Mês', 'MES', 'Mes'],
            'ano': ['ano', 'Ano', 'ANO'],
            'vendedor': ['Vendedor', 'vendedor', 'VENDEDOR'],
            'valor_final': ['Valor Final', 'valor final', 'VALOR FINAL'],
            'tipo_servico': ['Tipo de Serviço', 'tipo de serviço', 'TIPO DE SERVIÇO', 'Tipo de Servico']
        }
        
        # Mapear colunas existentes
        colunas_mapeadas = {}
        for campo, alternativas in colunas_alternativas.items():
            for alt in alternativas:
                if alt in df_vendas.columns:
                    colunas_mapeadas[campo] = alt
                    break
        
        # Verificar se todas as colunas necessárias existem
        if len(colunas_mapeadas) < 6:
            return {}
        
        # Converter colunas para tipos adequados
        df_vendas[colunas_mapeadas['dia']] = pd.to_numeric(df_vendas[colunas_mapeadas['dia']], errors='coerce')
        df_vendas[colunas_mapeadas['ano']] = pd.to_numeric(df_vendas[colunas_mapeadas['ano']], errors='coerce')
        
        # Limpar e converter Valor Final
        df_vendas['valor_final_terceiro_limpo'] = df_vendas[colunas_mapeadas['valor_final']].astype(str)
        df_vendas['valor_final_terceiro_limpo'] = df_vendas['valor_final_terceiro_limpo'].str.replace('R$', '', regex=False)
        df_vendas['valor_final_terceiro_limpo'] = df_vendas['valor_final_terceiro_limpo'].str.replace('.', '', regex=False)
        df_vendas['valor_final_terceiro_limpo'] = df_vendas['valor_final_terceiro_limpo'].str.replace(',', '.', regex=False)
        df_vendas['valor_final_terceiro_limpo'] = df_vendas['valor_final_terceiro_limpo'].str.strip()
        df_vendas['valor_final_terceiro_limpo'] = pd.to_numeric(df_vendas['valor_final_terceiro_limpo'], errors='coerce').fillna(0)
        
        # Converter meses de string para número
        meses_para_numeros = {
            'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4, 
            'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8,
            'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12,
            'Marco': 3, 'Decembro': 12
        }
        
        df_vendas['mes_numero_terceiro'] = df_vendas[colunas_mapeadas['mês']].map(meses_para_numeros)
        
        # Filtrar por período
        mask_periodo = (
            (df_vendas[colunas_mapeadas['ano']] >= ano_inicial) & 
            (df_vendas[colunas_mapeadas['ano']] <= ano_final) &
            (df_vendas['mes_numero_terceiro'] >= mes_inicial) & 
            (df_vendas['mes_numero_terceiro'] <= mes_final) &
            (df_vendas[colunas_mapeadas['dia']] >= dia_inicial) & 
            (df_vendas[colunas_mapeadas['dia']] <= dia_final)
        )
        
        df_periodo = df_vendas[mask_periodo]
        
        # Filtrar apenas Tipo de Serviço = "Terceiro"
        mask_terceiro = (df_periodo[colunas_mapeadas['tipo_servico']] == 'Terceiro')
        df_terceiro = df_periodo[mask_terceiro]
        
        # Normalizar vendedor para comparação case-insensitive
        df_terceiro['vendedor_normalizado'] = df_terceiro[colunas_mapeadas['vendedor']].str.strip().str.upper()
        
        # Calcular soma por vendedor
        vendas_por_vendedor = {}
        
        for vendedor in vendedores_list:
            vendedor_normalizado = vendedor.strip().upper()
            vendedor_dados = df_terceiro[df_terceiro['vendedor_normalizado'] == vendedor_normalizado]
            soma = vendedor_dados['valor_final_terceiro_limpo'].sum()
            vendas_por_vendedor[vendedor] = soma
        
        return vendas_por_vendedor
        
    except Exception as e:
        st.error(f"Erro ao calcular Vendas Terceiros para Online/Desks: {e}")
        return {}

# Função para calcular Meta Diaria para Online e Desks
def calcular_meta_diaria_online_desks(df_meta_diaria, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final):
    """
    Calcula a Meta Diaria multiplicada pelo número de dias do período
    """
    try:
        if df_meta_diaria.empty:
            return {}
        
        # Calcular total de dias no período
        from datetime import date
        data_inicial = date(ano_inicial, mes_inicial, dia_inicial)
        data_final = date(ano_final, mes_final, dia_final)
        total_dias = (data_final - data_inicial).days + 1
        
        # Mapear colunas da aba "Meta Diaria"
        colunas_alternativas = {
            'vendedor': ['Vendedor', 'vendedor', 'VENDEDOR', 'Nome do Vendedor', 'Nome Do Vendedor'],
            'data': ['Data', 'data', 'DATA'],
            'meta_diaria': ['Meta Diaria', 'Meta Diária', 'meta diaria', 'META DIARIA', 'Meta']
        }
        
        # Mapear colunas existentes
        colunas_mapeadas = {}
        for campo, alternativas in colunas_alternativas.items():
            for alt in alternativas:
                if alt in df_meta_diaria.columns:
                    colunas_mapeadas[campo] = alt
                    break
        
        # Verificar se todas as colunas necessárias existem
        if len(colunas_mapeadas) < 3:
            st.warning(f"Colunas disponíveis em Meta Diaria: {list(df_meta_diaria.columns)}")
            return {}
        
        # Converter coluna Data para datetime (normalizar para date)
        df_meta_diaria['data_convertida'] = pd.to_datetime(df_meta_diaria[colunas_mapeadas['data']], format='%d/%m/%Y', errors='coerce')
        
        # Limpar e converter Meta Diaria primeiro
        df_meta_diaria['meta_diaria_limpa'] = df_meta_diaria[colunas_mapeadas['meta_diaria']].astype(str)
        df_meta_diaria['meta_diaria_limpa'] = df_meta_diaria['meta_diaria_limpa'].str.replace('R$', '', regex=False)
        df_meta_diaria['meta_diaria_limpa'] = df_meta_diaria['meta_diaria_limpa'].str.replace('.', '', regex=False)
        df_meta_diaria['meta_diaria_limpa'] = df_meta_diaria['meta_diaria_limpa'].str.replace(',', '.', regex=False)
        df_meta_diaria['meta_diaria_limpa'] = df_meta_diaria['meta_diaria_limpa'].str.strip()
        df_meta_diaria['meta_diaria_limpa'] = pd.to_numeric(df_meta_diaria['meta_diaria_limpa'], errors='coerce').fillna(0)
        
        # Normalizar nomes de vendedores para comparação case-insensitive
        df_meta_diaria['vendedor_normalizado'] = df_meta_diaria[colunas_mapeadas['vendedor']].str.strip().str.upper()
        
        # Calcular meta por vendedor
        meta_por_vendedor = {}
        
        for vendedor in vendedores_list:
            vendedor_normalizado = vendedor.strip().upper()
            vendedor_dados = df_meta_diaria[df_meta_diaria['vendedor_normalizado'] == vendedor_normalizado]
            
            if not vendedor_dados.empty:
                # Filtrar por período se possível, senão pegar qualquer meta do vendedor
                data_inicial_ts = pd.Timestamp(data_inicial)
                data_final_ts = pd.Timestamp(data_final)
                
                vendedor_periodo = vendedor_dados[
                    (vendedor_dados['data_convertida'] >= data_inicial_ts) &
                    (vendedor_dados['data_convertida'] <= data_final_ts)
                ]
                
                # Se não encontrar no período, pegar a meta mais recente do vendedor
                if vendedor_periodo.empty:
                    vendedor_periodo = vendedor_dados.sort_values('data_convertida', ascending=False)
                
                if not vendedor_periodo.empty:
                    # Pegar a primeira meta diaria encontrada
                    meta_diaria_valor = vendedor_periodo['meta_diaria_limpa'].iloc[0]
                    # Multiplicar pelo total de dias
                    meta_total = meta_diaria_valor * total_dias
                    meta_por_vendedor[vendedor] = meta_total
                else:
                    meta_por_vendedor[vendedor] = 0
            else:
                meta_por_vendedor[vendedor] = 0
        
        return meta_por_vendedor
        
    except Exception as e:
        st.error(f"Erro ao calcular Meta Diaria para Online/Desks: {e}")
        import traceback
        st.error(traceback.format_exc())
        return {}

# Função para calcular Meta para Online e Desks
def calcular_meta_online_desks(df_vendedores, vendedores_list, mes_inicial, ano_inicial, mes_final, ano_final):
    """
    Calcula a Meta da aba Vendedores para vendedores Online e Desks
    Match: Vendedor + Mês (inicial e final) + Ano (inicial e final)
    Retorna o valor da coluna Meta
    """
    try:
        if df_vendedores.empty:
            return {}
        
        # Mapear colunas da aba "Vendedores"
        colunas_alternativas = {
            'vendedor': ['Nome Do Vendedor', 'Nome do Vendedor', 'Vendedor', 'vendedor', 'VENDEDOR'],
            'mes': ['Mês', 'mês', 'MES', 'Mes'],
            'ano': ['Ano', 'ano', 'ANO'],
            'meta': ['Meta', 'meta', 'META']
        }
        
        # Mapear colunas existentes
        colunas_mapeadas = {}
        for campo, alternativas in colunas_alternativas.items():
            for alt in alternativas:
                if alt in df_vendedores.columns:
                    colunas_mapeadas[campo] = alt
                    break
        
        # Verificar se todas as colunas necessárias existem
        if len(colunas_mapeadas) < 4:
            st.warning(f"Colunas disponíveis em Vendedores para Meta: {list(df_vendedores.columns)}")
            return {}
        
        # Converter coluna Ano para numérico
        df_vendedores['ano_numerico'] = pd.to_numeric(df_vendedores[colunas_mapeadas['ano']], errors='coerce')
        
        # Limpar e converter Meta
        df_vendedores['meta_limpa'] = df_vendedores[colunas_mapeadas['meta']].astype(str)
        df_vendedores['meta_limpa'] = df_vendedores['meta_limpa'].str.replace('R$', '', regex=False)
        df_vendedores['meta_limpa'] = df_vendedores['meta_limpa'].str.replace('.', '', regex=False)
        df_vendedores['meta_limpa'] = df_vendedores['meta_limpa'].str.replace(',', '.', regex=False)
        df_vendedores['meta_limpa'] = df_vendedores['meta_limpa'].str.strip()
        df_vendedores['meta_limpa'] = pd.to_numeric(df_vendedores['meta_limpa'], errors='coerce').fillna(0)
        
        # Normalizar vendedor para comparação case-insensitive
        df_vendedores['vendedor_normalizado'] = df_vendedores[colunas_mapeadas['vendedor']].str.strip().str.upper()
        
        # Converter meses de string para número (se necessário)
        meses_para_numeros = {
            'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4, 
            'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8,
            'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12,
            'Marco': 3, 'Decembro': 12
        }
        
        # Tentar converter mês (pode ser string ou número)
        if df_vendedores[colunas_mapeadas['mes']].dtype == 'object':
            df_vendedores['mes_numerico'] = df_vendedores[colunas_mapeadas['mes']].map(meses_para_numeros)
        else:
            df_vendedores['mes_numerico'] = pd.to_numeric(df_vendedores[colunas_mapeadas['mes']], errors='coerce')
        
        # Calcular meta por vendedor
        meta_por_vendedor = {}
        
        for vendedor in vendedores_list:
            vendedor_normalizado = vendedor.strip().upper()
            
            # Filtrar por vendedor e período (mês e ano)
            vendedor_dados = df_vendedores[
                (df_vendedores['vendedor_normalizado'] == vendedor_normalizado) &
                (df_vendedores['ano_numerico'] >= ano_inicial) &
                (df_vendedores['ano_numerico'] <= ano_final) &
                (df_vendedores['mes_numerico'] >= mes_inicial) &
                (df_vendedores['mes_numerico'] <= mes_final)
            ]
            
            if not vendedor_dados.empty:
                # Somar todas as metas do período
                meta_total = vendedor_dados['meta_limpa'].sum()
                meta_por_vendedor[vendedor] = meta_total
            else:
                meta_por_vendedor[vendedor] = 0
        
        return meta_por_vendedor
        
    except Exception as e:
        st.error(f"Erro ao calcular Meta para Online/Desks: {e}")
        import traceback
        st.error(traceback.format_exc())
        return {}

# Função para calcular Paxs In para Transferistas e Guias
def calcular_paxs_in(df_paxs, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final):
    """
    Calcula os Paxs In para uma lista de vendedores no período especificado
    Filtra por Guia (match com Vendedor) + All Inclusive = Não + período selecionado
    Soma a coluna Total_Paxs
    """
    try:
        if df_paxs.empty:
            return {}
        
        # Verificar se existem as colunas necessárias
        colunas_alternativas = {
            'dia': ['dia', 'Dia', 'DIA'],
            'mês': ['mês', 'Mês', 'MES', 'Mes'],
            'ano': ['ano', 'Ano', 'ANO'],
            'guia': ['Guia', 'guia', 'GUIA'],
            'total_paxs': ['Total_Paxs', 'total_paxs', 'TOTAL_PAXS', 'Total Paxs'],
            'all_inclusive': ['All Inclusive', 'all inclusive', 'ALL INCLUSIVE']
        }
        
        # Mapear colunas existentes
        colunas_mapeadas = {}
        for campo, alternativas in colunas_alternativas.items():
            for alt in alternativas:
                if alt in df_paxs.columns:
                    colunas_mapeadas[campo] = alt
                    break
        
        # Verificar se todas as colunas foram encontradas
        if len(colunas_mapeadas) < 6:
            return {}
        
        # Converter colunas para tipos adequados
        
        # DIA e ANO são números
        df_paxs[colunas_mapeadas['dia']] = pd.to_numeric(df_paxs[colunas_mapeadas['dia']], errors='coerce')
        df_paxs[colunas_mapeadas['ano']] = pd.to_numeric(df_paxs[colunas_mapeadas['ano']], errors='coerce')
        
        # TOTAL_PAXS - converter para numérico
        df_paxs[colunas_mapeadas['total_paxs']] = pd.to_numeric(df_paxs[colunas_mapeadas['total_paxs']], errors='coerce').fillna(0)
        
        # MÊS é STRING - converter nomes de meses para números
        meses_para_numeros = {
            'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4, 
            'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8,
            'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12,
            # Versões sem acento também
            'Marco': 3, 'Decembro': 12
        }
        
        # Converter meses de string para número
        df_paxs['mes_numero'] = df_paxs[colunas_mapeadas['mês']].map(meses_para_numeros)
        
        # Filtrar por período usando dia, mês e ano separadamente
        mask_periodo = (
            (df_paxs[colunas_mapeadas['ano']] >= ano_inicial) & 
            (df_paxs[colunas_mapeadas['ano']] <= ano_final) &
            (df_paxs['mes_numero'] >= mes_inicial) & 
            (df_paxs['mes_numero'] <= mes_final) &
            (df_paxs[colunas_mapeadas['dia']] >= dia_inicial) & 
            (df_paxs[colunas_mapeadas['dia']] <= dia_final)
        )
        
        df_periodo = df_paxs[mask_periodo]
        
        # Filtrar por All Inclusive = Não
        mask_all_inclusive = (df_periodo[colunas_mapeadas['all_inclusive']] == 'Não')
        
        df_filtrado = df_periodo[mask_all_inclusive]
        
        # Calcular soma por vendedor (matching Guia com Vendedor)
        paxs_por_vendedor = {}
        
        for vendedor in vendedores_list:
            vendedor_dados = df_filtrado[df_filtrado[colunas_mapeadas['guia']] == vendedor]
            soma = vendedor_dados[colunas_mapeadas['total_paxs']].sum()
            paxs_por_vendedor[vendedor] = soma
        
        return paxs_por_vendedor
        
    except Exception as e:
        st.error(f"Erro ao calcular Paxs In: {e}")
        return {}

# Função para calcular Paxs In All Inclusive para Transferistas e Guias
def calcular_paxs_in_all_inclusive(df_paxs, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final):
    """
    Calcula os Paxs In COM All Inclusive para uma lista de vendedores no período especificado
    Mesma lógica da função anterior, mas com All Inclusive = "Sim"
    """
    try:
        if df_paxs.empty:
            return {}
        
        # Verificar se existem as colunas necessárias
        colunas_alternativas = {
            'dia': ['dia', 'Dia', 'DIA'],
            'mês': ['mês', 'Mês', 'MES', 'Mes'],
            'ano': ['ano', 'Ano', 'ANO'],
            'guia': ['Guia', 'guia', 'GUIA'],
            'total_paxs': ['Total_Paxs', 'total_paxs', 'TOTAL_PAXS', 'Total Paxs'],
            'all_inclusive': ['All Inclusive', 'all inclusive', 'ALL INCLUSIVE']
        }
        
        # Mapear colunas existentes
        colunas_mapeadas = {}
        for campo, alternativas in colunas_alternativas.items():
            for alt in alternativas:
                if alt in df_paxs.columns:
                    colunas_mapeadas[campo] = alt
                    break
        
        # Verificar se todas as colunas foram encontradas
        if len(colunas_mapeadas) < 6:
            return {}
        
        # Converter colunas para tipos adequados
        df_paxs[colunas_mapeadas['dia']] = pd.to_numeric(df_paxs[colunas_mapeadas['dia']], errors='coerce')
        df_paxs[colunas_mapeadas['ano']] = pd.to_numeric(df_paxs[colunas_mapeadas['ano']], errors='coerce')
        
        # TOTAL_PAXS - converter para numérico
        df_paxs[colunas_mapeadas['total_paxs']] = pd.to_numeric(df_paxs[colunas_mapeadas['total_paxs']], errors='coerce').fillna(0)
        
        # MÊS é STRING - converter nomes de meses para números
        meses_para_numeros = {
            'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4, 
            'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8,
            'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12,
            'Marco': 3, 'Decembro': 12
        }
        
        # Converter meses de string para número
        df_paxs['mes_numero_ai'] = df_paxs[colunas_mapeadas['mês']].map(meses_para_numeros)
        
        # Filtrar por período usando dia, mês e ano separadamente
        mask_periodo = (
            (df_paxs[colunas_mapeadas['ano']] >= ano_inicial) & 
            (df_paxs[colunas_mapeadas['ano']] <= ano_final) &
            (df_paxs['mes_numero_ai'] >= mes_inicial) & 
            (df_paxs['mes_numero_ai'] <= mes_final) &
            (df_paxs[colunas_mapeadas['dia']] >= dia_inicial) & 
            (df_paxs[colunas_mapeadas['dia']] <= dia_final)
        )
        
        df_periodo = df_paxs[mask_periodo]
        
        # DIFERENÇA: Filtrar por All Inclusive = "Sim"
        mask_all_inclusive = (df_periodo[colunas_mapeadas['all_inclusive']] == 'Sim')
        
        df_filtrado = df_periodo[mask_all_inclusive]
        
        # Calcular soma por vendedor (matching Guia com Vendedor)
        paxs_por_vendedor = {}
        
        for vendedor in vendedores_list:
            vendedor_dados = df_filtrado[df_filtrado[colunas_mapeadas['guia']] == vendedor]
            soma = vendedor_dados[colunas_mapeadas['total_paxs']].sum()
            paxs_por_vendedor[vendedor] = soma
        
        return paxs_por_vendedor
        
    except Exception as e:
        st.error(f"Erro ao calcular Paxs In All Inclusive: {e}")
        return {}

# Função para buscar Meta por vendedor
def formatar_meta(valor):
    # Remove 'R$', pontos e converte vírgula para ponto
    if isinstance(valor, str):
        valor = valor.replace('R$', '').replace('.', '').replace(',', '.').strip()
    try:
        return float(valor)
    except:
        return 0.0

def buscar_meta_vendedor(df_vendedores, vendedor, mes_inicial, mes_final, ano_inicial, ano_final):
    """
    Busca a meta de um vendedor específico baseado no período selecionado
    """
    try:
        if df_vendedores.empty:
            return 0.0
        # Filtrar por vendedor
        df_vendedor = df_vendedores[df_vendedores['Nome Do Vendedor'] == vendedor]
        if df_vendedor.empty:
            return 0.0
        # Converter colunas para tipos adequados
        df_vendedor = df_vendedor.copy()
        df_vendedor['Ano'] = pd.to_numeric(df_vendedor['Ano'], errors='coerce')
        df_vendedor['mês'] = pd.to_numeric(df_vendedor['mês'], errors='coerce')
        # Remover linhas com valores inválidos
        df_vendedor = df_vendedor.dropna(subset=['Ano', 'mês'])
        if df_vendedor.empty:
            return 0.0
        # Filtrar por período (mesmo que no filtro principal)
        df_periodo = df_vendedor[
            (df_vendedor['Ano'] >= ano_inicial) & 
            (df_vendedor['Ano'] <= ano_final) &
            (df_vendedor['mês'] >= mes_inicial) & 
            (df_vendedor['mês'] <= mes_final)
        ]
        if df_periodo.empty:
            return 0.0
        # Se existe coluna Meta, somar os valores (tratando formato)
        if 'Meta' in df_periodo.columns:
            metas = df_periodo['Meta'].apply(formatar_meta)
            meta_total = metas.sum()
            return float(meta_total) if pd.notna(meta_total) else 0.0
        else:
            return 0.0
    except Exception as e:
        return 0.0

# Função para buscar se a venda é All Inclusive (mesma lógica do painel_vendedores.py)
def buscar_venda_all_inclusive(data_venda, vendedor, codigo_reserva, servico, vendas_finais_df):
    """Busca se a venda é All Inclusive na planilha externa"""
    try:
        from datetime import datetime
        
        # Validações iniciais
        if pd.isna(data_venda) or pd.isna(vendedor) or pd.isna(codigo_reserva):
            return 'Não'
        
        # Converter data de dd/mm/aaaa para aaaa-mm-dd
        def converter_data_para_iso(data_str):
            try:
                if pd.isna(data_str) or str(data_str).strip() == '':
                    return None
                # Formato dd/mm/aaaa → aaaa-mm-dd
                data_obj = datetime.strptime(str(data_str).strip(), '%d/%m/%Y')
                return data_obj.strftime('%Y-%m-%d')
            except:
                return None
        
        data_iso = converter_data_para_iso(data_venda)
        if not data_iso:
            return 'Não'
        
        # Normalizar strings para comparação
        def normalizar_string(s):
            if pd.isna(s):
                return ''
            return str(s).strip()
        
        # Tentar diferentes variações de match
        vendedor_norm = normalizar_string(vendedor)
        codigo_norm = normalizar_string(codigo_reserva)
        
        # Verificar se as colunas existem na planilha
        colunas_disponíveis = vendas_finais_df.columns.tolist()
        
        # Tentar diferentes nomes de coluna para Data_Venda
        coluna_data = None
        for col in ['Data_Venda', 'Data da Venda', 'Data Venda', 'Data']:
            if col in colunas_disponíveis:
                coluna_data = col
                break
        
        # Tentar diferentes nomes de coluna para Reserva
        coluna_reserva = None
        for col in ['Reserva', 'Código da Reserva', 'Codigo da Reserva', 'Reservation']:
            if col in colunas_disponíveis:
                coluna_reserva = col
                break
        
        # Tentar diferentes nomes de coluna para ALL Inclusive
        coluna_all_inclusive = None
        for col in ['ALL Inclusive', 'All Inclusive', 'ALL_Inclusive', 'All_Inclusive']:
            if col in colunas_disponíveis:
                coluna_all_inclusive = col
                break
        
        if not all([coluna_data, coluna_reserva, coluna_all_inclusive]):
            return 'Não'
        
        # Filtrar por data e vendedor primeiro (match mais provável)
        filtro_inicial = (
            (vendas_finais_df[coluna_data].astype(str).str.strip() == data_iso) &
            (vendas_finais_df['Vendedor'].apply(normalizar_string) == vendedor_norm)
        )
        
        # Se encontrou registros, tentar match com reserva
        registros_data_vendedor = vendas_finais_df[filtro_inicial]
        
        if not registros_data_vendedor.empty:
            # Tentar match exato com código da reserva
            filtro_reserva = registros_data_vendedor[coluna_reserva].apply(normalizar_string) == codigo_norm
            registros_com_reserva = registros_data_vendedor[filtro_reserva]
            
            if not registros_com_reserva.empty:
                # Encontrou match completo
                valor_all_inclusive = registros_com_reserva[coluna_all_inclusive].iloc[0]
                valor_str = str(valor_all_inclusive).strip().lower()
                return 'Sim' if valor_str in ['sim', 'yes', '1', 'true', 's'] else 'Não'
            else:
                # Não encontrou match com reserva, tentar apenas data + vendedor
                valor_all_inclusive = registros_data_vendedor[coluna_all_inclusive].iloc[0]
                valor_str = str(valor_all_inclusive).strip().lower()
                return 'Sim' if valor_str in ['sim', 'yes', '1', 'true', 's'] else 'Não'
        
        return 'Não'
            
    except Exception as e:
        return 'Não'

# Função para filtrar dados de comissão por período e vendedor
def filtrar_comissao_por_periodo_vendedor(df_comissao, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final):
    """
    Filtra os dados de comissão por período de data e lista de vendedores
    """
    try:
        if df_comissao.empty:
            return pd.DataFrame()
        
        # Verificar se existem as colunas necessárias
        colunas_necessarias = ['Data da Venda', 'Vendedor', 'Código da Reserva', 'Serviço', 'Valor da Venda']
        colunas_existentes = [col for col in colunas_necessarias if col in df_comissao.columns]
        
        if len(colunas_existentes) < 4:  # Pelo menos 4 colunas principais
            return pd.DataFrame()
        
        # Converter coluna Data da Venda para datetime
        def converter_data_comissao(data_str):
            try:
                if pd.isna(data_str) or str(data_str).strip() == '':
                    return None
                # Tentar diferentes formatos de data
                data_str = str(data_str).strip()
                
                # Formato dd/mm/yyyy
                if '/' in data_str:
                    return pd.to_datetime(data_str, format='%d/%m/%Y', errors='coerce')
                # Formato yyyy-mm-dd
                elif '-' in data_str:
                    return pd.to_datetime(data_str, format='%Y-%m-%d', errors='coerce')
                else:
                    return pd.to_datetime(data_str, errors='coerce')
            except:
                return None
        
        # Aplicar conversão de data
        df_comissao['data_convertida'] = df_comissao['Data da Venda'].apply(converter_data_comissao)
        
        # Remover linhas com datas inválidas
        df_comissao = df_comissao.dropna(subset=['data_convertida'])
        
        if df_comissao.empty:
            return pd.DataFrame()
        
        # Criar data inicial e final para filtro
        from datetime import datetime
        data_inicial_filtro = datetime(ano_inicial, mes_inicial, dia_inicial)
        data_final_filtro = datetime(ano_final, mes_final, dia_final)
        
        # Filtrar por período
        mask_periodo = (
            (df_comissao['data_convertida'] >= data_inicial_filtro) &
            (df_comissao['data_convertida'] <= data_final_filtro)
        )
        
        df_periodo = df_comissao[mask_periodo]
        
        if df_periodo.empty:
            return pd.DataFrame()
        
        # Filtrar por vendedores
        if vendedores_list:
            mask_vendedores = df_periodo['Vendedor'].isin(vendedores_list)
            df_filtrado = df_periodo[mask_vendedores]
        else:
            df_filtrado = df_periodo
        
        # Selecionar apenas as colunas solicitadas
        colunas_resultado = ['Data da Venda', 'Código da Reserva', 'Serviço', 'Valor da Venda']
        colunas_resultado.insert(1, 'Vendedor')  # Adicionar Vendedor na segunda posição
        
        # Filtrar colunas que existem
        colunas_finais = [col for col in colunas_resultado if col in df_filtrado.columns]
        
        resultado = df_filtrado[colunas_finais].copy()
        
        # NOVA COLUNA: Venda All Inclusive
        # Buscar se a venda é All Inclusive usando dados de vendas finais
        if 'df_vendas' in globals() and not df_vendas.empty:
            resultado['Venda All Inclusive'] = resultado.apply(
                lambda row: buscar_venda_all_inclusive(
                    row.get('Data da Venda', ''),
                    row.get('Vendedor', ''),
                    row.get('Código da Reserva', ''),
                    row.get('Serviço', ''),
                    df_vendas
                ), axis=1
            )
        else:
            resultado['Venda All Inclusive'] = 'Não'
        
        # NOVA COLUNA: Tipo de Serviço
        # Carregar lista de serviços terceirizados e classificar cada serviço
        servicos_terceiros = carregar_servicos_terceiros()
        
        def classificar_tipo_servico(row):
            servico = str(row.get('Serviço', '')).strip()
            if servico in servicos_terceiros:
                return 'Terceiro'
            else:
                return 'Luck'
        
        resultado['Tipo de Serviço'] = resultado.apply(classificar_tipo_servico, axis=1)
        
        # NOVA COLUNA: Comissão Luck
        # Carregar dados de vendedores e buscar comissão por vendedor/período
        df_vendedores = carregar_dados_vendedores()
        
        def buscar_comissao_luck_row(row):
            try:
                # Extrair mês e ano da data de venda
                data_venda = str(row.get('Data da Venda', '')).strip()
                if not data_venda:
                    return ''
                
                # Converter data de dd/mm/yyyy para extrair mês e ano
                try:
                    if '/' in data_venda:
                        partes = data_venda.split('/')
                        if len(partes) >= 3:
                            mes = int(partes[1])  # mês
                            ano = int(partes[2])  # ano
                    else:
                        return ''
                except:
                    return ''
                
                vendedor = str(row.get('Vendedor', '')).strip()
                
                return buscar_comissao_luck(vendedor, mes, ano, df_vendedores)
            except:
                return ''
        
        resultado['Comissão Luck'] = resultado.apply(buscar_comissao_luck_row, axis=1)
        
        # NOVA COLUNA: Comissão Terceiros
        # Buscar comissão terceiros por vendedor/período
        def buscar_comissao_terceiros_row(row):
            try:
                # Extrair mês e ano da data de venda
                data_venda = str(row.get('Data da Venda', '')).strip()
                if not data_venda:
                    return ''
                
                # Converter data de dd/mm/yyyy para extrair mês e ano
                try:
                    if '/' in data_venda:
                        partes = data_venda.split('/')
                        if len(partes) >= 3:
                            mes = int(partes[1])  # mês
                            ano = int(partes[2])  # ano
                    else:
                        return ''
                except:
                    return ''
                
                vendedor = str(row.get('Vendedor', '')).strip()
                
                return buscar_comissao_terceiros(vendedor, mes, ano, df_vendedores)
            except:
                return ''
        
        resultado['Comissão Terceiros'] = resultado.apply(buscar_comissao_terceiros_row, axis=1)
        
        # NOVA COLUNA: Premiação (versão simplificada que funciona)
        # Aplica premiação baseada no vendedor ser Transferista
        def buscar_premiacao_row(row):
            try:
                vendedor = str(row.get('Vendedor', '')).strip()
                
                if not vendedor:
                    return '0%'
                
                # Buscar o valor da Premiação diretamente do dicionário global
                # que foi preenchido pelo Grid Por Tipo de Vendedor
                if 'premiacao_por_vendedor' in globals():
                    premiacao_dict = globals()['premiacao_por_vendedor']
                    
                    # Tentar buscar por nome exato
                    if vendedor in premiacao_dict:
                        return premiacao_dict[vendedor]
                    
                    # Tentar buscar normalizando os nomes
                    vendedor_norm = normalizar_nome(vendedor)
                    for vend_key, premiacao_valor in premiacao_dict.items():
                        if normalizar_nome(vend_key) == vendedor_norm:
                            return premiacao_valor
                
                return '0%'  # Default se não encontrar
                
            except Exception:
                return '0%'
        
        resultado['Premiação'] = resultado.apply(buscar_premiacao_row, axis=1)
        
        # NOVA COLUNA: Premiação All Inclusive (baseada no Alcance de Meta All Inclusive)
        # Aplica premiação baseada no Alcance de Meta All Inclusive para Transferistas
        def buscar_premiacao_all_inclusive_row(row):
            try:
                vendedor = str(row.get('Vendedor', '')).strip()
                
                if not vendedor:
                    return '0%'
                
                # Buscar o valor da Premiação All Inclusive diretamente do dicionário global
                # que foi preenchido pelo Grid All Inclusive - Transferistas
                if 'premiacao_ai_por_vendedor' in globals():
                    premiacao_ai_dict = globals()['premiacao_ai_por_vendedor']
                    
                    # Tentar buscar por nome exato
                    if vendedor in premiacao_ai_dict:
                        return premiacao_ai_dict[vendedor]
                    
                    # Tentar buscar normalizando os nomes
                    vendedor_norm = normalizar_nome(vendedor)
                    for vend_key, premiacao_ai_valor in premiacao_ai_dict.items():
                        if normalizar_nome(vend_key) == vendedor_norm:
                            return premiacao_ai_valor
                
                return '0%'  # Default se não encontrar
                
            except Exception:
                return '0%'
        
        resultado['Premiação All Inclusive'] = resultado.apply(buscar_premiacao_all_inclusive_row, axis=1)
        
        # NOVA COLUNA: Valor Comissão Luck (POSICIÓN 12 - Última coluna)
        # Multiplica Valor da Venda pela Comissão Luck se Tipo de Serviço for "Luck"
        def calcular_valor_comissao_luck(row):
            try:
                # Verificar se o Tipo de Serviço é "Luck"
                tipo_servico = str(row.get('Tipo de Serviço', '')).strip()
                
                if tipo_servico != 'Luck':
                    return 'R$ 0,00'
                
                # Obter Valor da Venda
                valor_venda_str = str(row.get('Valor da Venda', 'R$ 0,00'))
                valor_venda_limpo = valor_venda_str.replace('R$', '').replace('.', '').replace(',', '.').strip()
                valor_venda = float(valor_venda_limpo) if valor_venda_limpo not in ['', '0', '0,00'] else 0
                
                # Obter Comissão Luck (converter percentual para decimal)
                comissao_luck_str = str(row.get('Comissão Luck', '0%')).replace('%', '').strip()
                comissao_luck = float(comissao_luck_str) / 100 if comissao_luck_str not in ['', '0'] else 0
                
                # Calcular valor da comissão
                valor_comissao = valor_venda * comissao_luck
                
                # Formatar como moeda
                return f"R$ {valor_comissao:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                
            except Exception:
                return 'R$ 0,00'
        
        resultado['Valor Comissão Luck'] = resultado.apply(calcular_valor_comissao_luck, axis=1)
        
        # NOVA COLUNA: Valor Comissão Terceiros (POSIÇÃO 13 - após Valor Comissão Luck)
        # Multiplica Valor da Venda pela Comissão Terceiros se Tipo de Serviço for "Terceiro"
        def calcular_valor_comissao_terceiros(row):
            try:
                # Verificar se o Tipo de Serviço é "Terceiro"
                tipo_servico = str(row.get('Tipo de Serviço', '')).strip()
                
                if tipo_servico != 'Terceiro':
                    return 'R$ 0,00'
                
                # Obter Valor da Venda
                valor_venda_str = str(row.get('Valor da Venda', 'R$ 0,00'))
                valor_venda_limpo = valor_venda_str.replace('R$', '').replace('.', '').replace(',', '.').strip()
                valor_venda = float(valor_venda_limpo) if valor_venda_limpo not in ['', '0', '0,00'] else 0
                
                # Obter Comissão Terceiros (converter percentual para decimal)
                comissao_terceiros_str = str(row.get('Comissão Terceiros', '0%')).replace('%', '').strip()
                comissao_terceiros = float(comissao_terceiros_str) / 100 if comissao_terceiros_str not in ['', '0'] else 0
                
                # Calcular valor da comissão
                valor_comissao = valor_venda * comissao_terceiros
                
                # Formatar como moeda
                return f"R$ {valor_comissao:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                
            except Exception:
                return 'R$ 0,00'
        
        resultado['Valor Comissão Terceiros'] = resultado.apply(calcular_valor_comissao_terceiros, axis=1)
        
        # NOVA COLUNA: Valor Comissão Premiação (apenas para Transferistas)
        # Multiplica Valor da Venda pela Premiação se Venda All Inclusive = "Não" e Tipo de Serviço = "Luck"
        def calcular_valor_comissao_premiacao(row):
            try:
                # Verificar condições: Venda All Inclusive = "Não" E Tipo de Serviço = "Luck"
                venda_all_inclusive = str(row.get('Venda All Inclusive', '')).strip()
                tipo_servico = str(row.get('Tipo de Serviço', '')).strip()
                
                if venda_all_inclusive != 'Não' or tipo_servico != 'Luck':
                    return 'R$ 0,00'
                
                # Obter Valor da Venda
                valor_venda_str = str(row.get('Valor da Venda', 'R$ 0,00'))
                valor_venda_limpo = valor_venda_str.replace('R$', '').replace('.', '').replace(',', '.').strip()
                valor_venda = float(valor_venda_limpo) if valor_venda_limpo not in ['', '0', '0,00'] else 0
                
                # Obter Premiação (converter percentual para decimal)
                premiacao_str = str(row.get('Premiação', '0%')).replace('%', '').strip()
                premiacao = float(premiacao_str) / 100 if premiacao_str not in ['', '0'] else 0
                
                # Calcular valor da comissão
                valor_comissao = valor_venda * premiacao
                
                # Formatar como moeda
                return f"R$ {valor_comissao:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                
            except Exception:
                return 'R$ 0,00'
        
        resultado['Valor Comissão Premiação'] = resultado.apply(calcular_valor_comissao_premiacao, axis=1)
        
        # NOVA COLUNA: Valor Comissão Premiação All Inclusive (apenas para Transferistas)
        # Multiplica Valor da Venda pela Premiação All Inclusive se Venda All Inclusive = "Sim" e Tipo de Serviço = "Luck"
        def calcular_valor_comissao_premiacao_ai(row):
            try:
                # Verificar condições: Venda All Inclusive = "Sim" E Tipo de Serviço = "Luck"
                venda_all_inclusive = str(row.get('Venda All Inclusive', '')).strip()
                tipo_servico = str(row.get('Tipo de Serviço', '')).strip()
                
                if venda_all_inclusive != 'Sim' or tipo_servico != 'Luck':
                    return 'R$ 0,00'
                
                # Obter Valor da Venda
                valor_venda_str = str(row.get('Valor da Venda', 'R$ 0,00'))
                valor_venda_limpo = valor_venda_str.replace('R$', '').replace('.', '').replace(',', '.').strip()
                valor_venda = float(valor_venda_limpo) if valor_venda_limpo not in ['', '0', '0,00'] else 0
                
                # Obter Premiação All Inclusive (converter percentual para decimal)
                premiacao_ai_str = str(row.get('Premiação All Inclusive', '0%')).replace('%', '').strip()
                premiacao_ai = float(premiacao_ai_str) / 100 if premiacao_ai_str not in ['', '0'] else 0
                
                # Calcular valor da comissão
                valor_comissao = valor_venda * premiacao_ai
                
                # Formatar como moeda
                return f"R$ {valor_comissao:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                
            except Exception:
                return 'R$ 0,00'
        
        resultado['Valor Comissão Premiação All Inclusive'] = resultado.apply(calcular_valor_comissao_premiacao_ai, axis=1)
        
        # NOVA COLUNA: Valor Total de Comissão (POSIÇÃO 16 - última coluna)
        # Soma de todas as comissões: Luck + Terceiros + Premiação + Premiação All Inclusive
        def calcular_valor_total_comissao(row):
            try:
                # Obter todos os valores de comissão
                valores_str = [
                    str(row.get('Valor Comissão Luck', 'R$ 0,00')),
                    str(row.get('Valor Comissão Terceiros', 'R$ 0,00')),
                    str(row.get('Valor Comissão Premiação', 'R$ 0,00')),
                    str(row.get('Valor Comissão Premiação All Inclusive', 'R$ 0,00'))
                ]
                
                # Converter para float e somar
                total = 0
                for valor_str in valores_str:
                    valor_limpo = valor_str.replace('R$', '').replace('.', '').replace(',', '.').strip()
                    valor_float = float(valor_limpo) if valor_limpo not in ['', '0', '0,00'] else 0
                    total += valor_float
                
                # Formatar como moeda
                return f"R$ {total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                
            except Exception:
                return 'R$ 0,00'
        
        resultado['Valor Total de Comissão'] = resultado.apply(calcular_valor_total_comissao, axis=1)
        
        # Formatar Valor da Venda como moeda se necessário
        if 'Valor da Venda' in resultado.columns:
            def formatar_moeda_comissao(valor):
                try:
                    if pd.isna(valor) or str(valor).strip() == '':
                        return 'R$ 0,00'
                    
                    # Se já está formatado, retornar como está
                    if 'R$' in str(valor):
                        return str(valor)
                    
                    # Converter para float e formatar
                    valor_float = float(str(valor).replace(',', '.'))
                    return f"R$ {valor_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                except:
                    return 'R$ 0,00'
            
            resultado['Valor da Venda'] = resultado['Valor da Venda'].apply(formatar_moeda_comissao)
        
        # Ordenar por data (mais recente primeiro)
        resultado = resultado.sort_values('Data da Venda', ascending=False)
        
        return resultado.reset_index(drop=True)
        
    except Exception as e:
        st.error(f"Erro ao filtrar dados de comissão: {e}")
        return pd.DataFrame()

# Configuração da página
st.set_page_config(
    page_title="Painel Diário",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Painel Diário de Análise")

# Criar colunas para os filtros de data
st.subheader("🗓️ Filtros de Data")

# Data Inicial
st.write("**Data Inicial:**")
col1, col2, col3 = st.columns(3)

with col2:
    mes_inicial = st.selectbox(
        "Mês Inicial",
        options=list(meses.keys()),
        format_func=lambda x: meses[x],
        index=0,
        key="mes_inicial"
    )

with col3:
    ano_inicial = st.selectbox(
        "Ano Inicial",
        options=list(range(2020, 2031)),
        index=5,  # 2025
        key="ano_inicial"
    )

# Calcular dias disponíveis para o mês inicial selecionado
dias_disponiveis_inicial = obter_dias_no_mes(mes_inicial, ano_inicial)

with col1:
    dia_inicial = st.selectbox(
        "Dia Inicial",
        options=list(range(1, dias_disponiveis_inicial + 1)),
        index=0,
        key="dia_inicial",
        help=f"Mês {meses[mes_inicial]} tem {dias_disponiveis_inicial} dias"
    )

st.markdown("---")

# Data Final
st.write("**Data Final:**")
col4, col5, col6 = st.columns(3)

with col5:
    mes_final = st.selectbox(
        "Mês Final",
        options=list(meses.keys()),
        format_func=lambda x: meses[x],
        index=datetime.now().month - 1,
        key="mes_final"
    )

with col6:
    ano_final = st.selectbox(
        "Ano Final",
        options=list(range(2020, 2031)),
        index=datetime.now().year - 2020,
        key="ano_final"
    )

# Calcular dias disponíveis para o mês final selecionado
dias_disponiveis_final = obter_dias_no_mes(mes_final, ano_final)

# Ajustar o dia final se for maior que o disponível no mês
dia_atual = datetime.now().day
dia_default_final = min(dia_atual, dias_disponiveis_final)

with col4:
    dia_final = st.selectbox(
        "Dia Final",
        options=list(range(1, dias_disponiveis_final + 1)),
        index=dia_default_final - 1,
        key="dia_final",
        help=f"Mês {meses[mes_final]} tem {dias_disponiveis_final} dias"
    )

st.markdown("---")

# Botão para carregar dados
if st.button("🔍 Carregar Dados", type="primary"):
    # Validação e criação das datas
    try:
        data_inicial = datetime(ano_inicial, mes_inicial, dia_inicial)
        data_final = datetime(ano_final, mes_final, dia_final)
        
        # Validar se a data inicial é menor ou igual à data final
        if data_inicial > data_final:
            st.error("⚠️ A data inicial não pode ser maior que a data final!")
        else:
            # Marcar que dados foram carregados (usar chaves diferentes para evitar conflito)
            st.session_state['dados_carregados'] = True
            st.session_state['periodo_data_inicial'] = data_inicial
            st.session_state['periodo_data_final'] = data_final
            st.session_state['periodo_dia_inicial'] = dia_inicial
            st.session_state['periodo_mes_inicial'] = mes_inicial
            st.session_state['periodo_ano_inicial'] = ano_inicial
            st.session_state['periodo_dia_final'] = dia_final
            st.session_state['periodo_mes_final'] = mes_final
            st.session_state['periodo_ano_final'] = ano_final
    except ValueError as e:
        st.error(f"⚠️ Data inválida! Verifique os valores inseridos. Erro: {e}")

# Exibir dados se já foram carregados
if st.session_state.get('dados_carregados', False):
    # Recuperar variáveis do session_state
    data_inicial = st.session_state['periodo_data_inicial']
    data_final = st.session_state['periodo_data_final']
    dia_inicial = st.session_state['periodo_dia_inicial']
    mes_inicial = st.session_state['periodo_mes_inicial']
    ano_inicial = st.session_state['periodo_ano_inicial']
    dia_final = st.session_state['periodo_dia_final']
    mes_final = st.session_state['periodo_mes_final']
    ano_final = st.session_state['periodo_ano_final']
    
    # Exibir as datas selecionadas
    st.success(f"✅ Período selecionado: {data_inicial.strftime('%d/%m/%Y')} até {data_final.strftime('%d/%m/%Y')}")
    
    # Calcular diferença de dias
    diferenca_dias = (data_final - data_inicial).days + 1
    st.info(f"📅 Total de dias no período: {diferenca_dias} dias")
    
    # Carregar dados do Google Sheets (com cache)
    with st.spinner("Carregando dados..."):
        df_vendedores = carregar_dados_google_sheets()
        df_vendas = carregar_dados_vendas()
        df_paxs_in = carregar_dados_paxs_in()
        df_comissao = carregar_dados_comissao()
        df_meta_diaria = carregar_dados_meta_diaria()
    
    if not df_vendedores.empty:
        # Filtrar dados por período (mês e ano)
        if 'mês' in df_vendedores.columns and 'Ano' in df_vendedores.columns:
            # Converter colunas para numérico para evitar erro de comparação
            df_vendedores['mês'] = pd.to_numeric(df_vendedores['mês'], errors='coerce')
            df_vendedores['Ano'] = pd.to_numeric(df_vendedores['Ano'], errors='coerce')
            
            # Filtrar pelo período selecionado
            df_filtrado = df_vendedores[
                (df_vendedores['Ano'] == ano_inicial) & 
                (df_vendedores['mês'] >= mes_inicial) & 
                (df_vendedores['mês'] <= mes_final)
            ]
            
            # Se ano inicial e final forem diferentes, ajustar o filtro
            if ano_inicial != ano_final:
                df_filtrado = df_vendedores[
                    ((df_vendedores['Ano'] == ano_inicial) & (df_vendedores['mês'] >= mes_inicial)) |
                    ((df_vendedores['Ano'] == ano_final) & (df_vendedores['mês'] <= mes_final)) |
                    ((df_vendedores['Ano'] > ano_inicial) & (df_vendedores['Ano'] < ano_final))
                ]
            
            if not df_filtrado.empty:
                # Separar por tipo de vendedor
                if 'Tipo de Vendedor' in df_filtrado.columns and 'Nome Do Vendedor' in df_filtrado.columns:
                    tipos_vendedor_raw = [tipo for tipo in df_filtrado['Tipo de Vendedor'].unique() if pd.notna(tipo) and tipo != '']
                    tipos_vendedor = ordenar_tipos_vendedor(tipos_vendedor_raw)
                    
                    st.markdown("---")
                    st.subheader("📊 Grids por Tipo de Vendedor")
                    
                    # Criar tabs para cada tipo de vendedor
                    if tipos_vendedor:
                        tabs = st.tabs(tipos_vendedor)
                        
                        for i, tipo in enumerate(tipos_vendedor):
                                    with tabs[i]:
                                        # Filtrar vendedores deste tipo no período
                                        df_tipo = df_filtrado[df_filtrado['Tipo de Vendedor'] == tipo]
                                        
                                        # Mostrar colunas Nome Do Vendedor e Tipo de Vendedor
                                        df_display = df_tipo[['Nome Do Vendedor', 'Tipo de Vendedor']].drop_duplicates().reset_index(drop=True)
                                        df_display = df_display.rename(columns={'Nome Do Vendedor': 'Vendedor'})
                                        
                                        # Adicionar coluna "Vendas Luck" para Online e Desks
                                        if tipo in ['Online', 'Desks'] and not df_vendas.empty:
                                            try:
                                                vendedores_list = df_display['Vendedor'].tolist()
                                                vendas_luck_online_desks = calcular_vendas_luck_online_desks(
                                                    df_vendas, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final
                                                )
                                                
                                                # Adicionar coluna ao dataframe
                                                df_display['Vendas Luck'] = df_display['Vendedor'].map(vendas_luck_online_desks).fillna(0)
                                                
                                                # Formatar valores como moeda
                                                df_display['Vendas Luck'] = df_display['Vendas Luck'].apply(
                                                    lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                )
                                                
                                            except Exception as e:
                                                st.error(f"Erro ao calcular Vendas Luck: {e}")
                                        
                                        # Adicionar coluna "Vendas Terceiros" para Online e Desks
                                        if tipo in ['Online', 'Desks'] and not df_vendas.empty:
                                            try:
                                                vendedores_list = df_display['Vendedor'].tolist()
                                                vendas_terceiros_online_desks = calcular_vendas_terceiros_online_desks(
                                                    df_vendas, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final
                                                )
                                                
                                                # Adicionar coluna ao dataframe
                                                df_display['Vendas Terceiros'] = df_display['Vendedor'].map(vendas_terceiros_online_desks).fillna(0)
                                                
                                                # Formatar valores como moeda
                                                df_display['Vendas Terceiros'] = df_display['Vendas Terceiros'].apply(
                                                    lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                )
                                                
                                            except Exception as e:
                                                st.error(f"Erro ao calcular Vendas Terceiros: {e}")
                                        
                                        # Adicionar coluna "Meta Diaria" para Online e Desks
                                        if tipo in ['Online', 'Desks'] and not df_meta_diaria.empty:
                                            try:
                                                vendedores_list = df_display['Vendedor'].tolist()
                                                meta_diaria_online_desks = calcular_meta_diaria_online_desks(
                                                    df_meta_diaria, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final
                                                )
                                                
                                                # Adicionar coluna ao dataframe
                                                df_display['Meta Diaria'] = df_display['Vendedor'].map(meta_diaria_online_desks).fillna(0)
                                                
                                                # Formatar valores como moeda
                                                df_display['Meta Diaria'] = df_display['Meta Diaria'].apply(
                                                    lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                )
                                                
                                            except Exception as e:
                                                st.error(f"Erro ao calcular Meta Diaria: {e}")
                                        
                                        # Adicionar coluna "Meta" para Online e Desks
                                        if tipo in ['Online', 'Desks'] and not df_vendedores.empty:
                                            try:
                                                vendedores_list = df_display['Vendedor'].tolist()
                                                meta_online_desks = calcular_meta_online_desks(
                                                    df_vendedores, vendedores_list, mes_inicial, ano_inicial, mes_final, ano_final
                                                )
                                                
                                                # Adicionar coluna ao dataframe
                                                df_display['Meta'] = df_display['Vendedor'].map(meta_online_desks).fillna(0)
                                                
                                                # Formatar valores como moeda
                                                df_display['Meta'] = df_display['Meta'].apply(
                                                    lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                )
                                                
                                            except Exception as e:
                                                st.error(f"Erro ao calcular Meta: {e}")
                                        
                                        # Adicionar coluna "Vendas Luck Sem Adicionais" apenas para Transferistas e Guias
                                        if tipo in ['Transferistas', 'Guias'] and not df_vendas.empty:
                                            try:
                                                vendedores_list = df_display['Vendedor'].tolist()
                                                vendas_luck = calcular_vendas_luck_sem_adicionais(
                                                    df_vendas, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final
                                                )
                                                
                                                # Adicionar coluna ao dataframe
                                                df_display['Vendas Luck Sem Adicionais'] = df_display['Vendedor'].map(vendas_luck).fillna(0)
                                                
                                                # Formatar valores como moeda
                                                df_display['Vendas Luck Sem Adicionais'] = df_display['Vendas Luck Sem Adicionais'].apply(
                                                    lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                )
                                                
                                            except Exception as e:
                                                st.error(f"Erro ao calcular Vendas Luck Sem Adicionais: {e}")
                                        
                                        # Adicionar coluna "Paxs In" apenas para Transferistas e Guias
                                        if tipo in ['Transferistas', 'Guias'] and not df_paxs_in.empty:
                                            try:
                                                vendedores_list = df_display['Vendedor'].tolist()
                                                paxs_in = calcular_paxs_in(
                                                    df_paxs_in, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final
                                                )
                                                
                                                # Adicionar coluna ao dataframe
                                                df_display['Paxs In'] = df_display['Vendedor'].map(paxs_in).fillna(0)
                                                
                                                # Formatar valores como decimal com 1 casa decimal (corrigindo divisão por 100)
                                                df_display['Paxs In'] = df_display['Paxs In'].apply(
                                                    lambda x: f"{float(x)/100:.1f}".replace('.', ',') if x > 0 else "0"
                                                )
                                                
                                            except Exception as e:
                                                st.error(f"Erro ao calcular Paxs In: {e}")
                                        
                                        # Adicionar coluna "Ticket Médio" apenas para Transferistas e Guias
                                        if tipo in ['Transferistas', 'Guias']:
                                            if 'Vendas Luck Sem Adicionais' in df_display.columns and 'Paxs In' in df_display.columns:
                                                try:
                                                    def calcular_ticket_medio(row):
                                                        try:
                                                            # Converter Vendas Luck Sem Adicionais para float
                                                            vendas_str = str(row['Vendas Luck Sem Adicionais']).replace('R$', '').replace('.', '').replace(',', '.').strip()
                                                            vendas_float = float(vendas_str) if vendas_str not in ['', '0', '0,00'] else 0
                                                            
                                                            # Converter Paxs In para float
                                                            paxs_str = str(row['Paxs In']).replace(',', '.').strip()
                                                            paxs_float = float(paxs_str) if paxs_str not in ['', '0', '0,0'] else 0
                                                            
                                                            # Calcular ticket médio
                                                            if paxs_float > 0:
                                                                ticket_medio = vendas_float / paxs_float
                                                                return f"R$ {ticket_medio:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                            else:
                                                                return "R$ 0,00"
                                                        except:
                                                            return "R$ 0,00"
                                                    
                                                    # Aplicar cálculo de ticket médio
                                                    df_display['Ticket Médio'] = df_display.apply(calcular_ticket_medio, axis=1)
                                                    
                                                except Exception as e:
                                                    st.error(f"Erro ao calcular Ticket Médio: {e}")
                                            
                                            # Adicionar coluna "Meta" apenas para Transferistas e Guias
                                            try:
                                                def buscar_meta_por_vendedor(row):
                                                    vendedor = row['Vendedor']
                                                    meta = buscar_meta_vendedor(df_vendedores, vendedor, mes_inicial, mes_final, ano_inicial, ano_final)
                                                    # Corrigir formatação da meta
                                                    if meta > 0:
                                                        return f"R$ {float(meta):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                    else:
                                                        return "R$ 0,00"
                                                # Aplicar busca de meta
                                                df_display['Meta'] = df_display.apply(buscar_meta_por_vendedor, axis=1)
                                            except Exception as e:
                                                st.error(f"Erro ao buscar Meta: {e}")

                                            # Adicionar coluna "Alcance de Meta" apenas para Transferistas e Guias
                                            try:
                                                def calcular_alcance_meta(row):
                                                    # Extrai valores das colunas
                                                    ticket_str = str(row.get('Ticket Médio', 'R$ 0,00')).replace('R$', '').replace('.', '').replace(',', '.').strip()
                                                    meta_str = str(row.get('Meta', 'R$ 0,00')).replace('R$', '').replace('.', '').replace(',', '.').strip()
                                                    try:
                                                        ticket = float(ticket_str) if ticket_str not in ['', '0', '0,00'] else 0
                                                        meta = float(meta_str) if meta_str not in ['', '0', '0,00'] else 0
                                                        if meta > 0:
                                                            alcance = ticket / meta
                                                            return f"{alcance:.2%}".replace('.', ',')
                                                        else:
                                                            return "0,00%"
                                                    except:
                                                        return "0,00%"
                                                # Aplicar cálculo de alcance de meta
                                                df_display['Alcance de Meta'] = df_display.apply(calcular_alcance_meta, axis=1)
                                                
                                                # Adicionar coluna Premiação apenas para Transferistas
                                                if tipo == 'Transferistas':
                                                    def calcular_premiacao_row(row):
                                                        alcance_meta = row.get('Alcance de Meta', '0,00%')
                                                        return calcular_premiacao_transferista(alcance_meta)
                                                    
                                                    df_display['Premiação'] = df_display.apply(calcular_premiacao_row, axis=1)
                                                    
                                                    # Armazenar valores de Premiação em variável global para uso no Grid Detalhes
                                                    globals()['premiacao_por_vendedor'] = dict(zip(df_display['Vendedor'], df_display['Premiação']))
                                            except Exception as e:
                                                st.error(f"Erro ao calcular Alcance de Meta: {e}")
                                        
                                        st.write(f"**Total de vendedores:** {len(df_display)}")
                                        
                                        # Mostrar grid com os vendedores
                                        def highlight_alcance_meta(row):
                                            valor = str(row.get('Alcance de Meta', '0,00%')).replace('%', '').replace(',', '.')
                                            try:
                                                valor_float = float(valor)
                                                if valor_float >= 100.0:
                                                    return ['background-color: #b6fcb6'] * len(row)
                                                else:
                                                    return [''] * len(row)
                                            except:
                                                return [''] * len(row)

                                        try:
                                            st.dataframe(
                                                df_display.style.apply(highlight_alcance_meta, axis=1),
                                                use_container_width=True,
                                                hide_index=True
                                            )
                                        except Exception as e:
                                            st.dataframe(
                                                df_display,
                                                use_container_width=True,
                                                hide_index=True
                                            )
                                        
                                        # Cartão com soma de Vendas Luck Sem Adicionais
                                        if tipo in ['Transferistas', 'Guias'] and 'Vendas Luck Sem Adicionais' in df_display.columns:
                                            try:
                                                # Extrair valores numéricos e somar
                                                def extrair_valor_vendas(valor_str):
                                                    try:
                                                        if pd.isna(valor_str):
                                                            return 0.0
                                                        valor_limpo = str(valor_str).replace('R$', '').replace('.', '').replace(',', '.').strip()
                                                        return float(valor_limpo) if valor_limpo else 0.0
                                                    except:
                                                        return 0.0
                                                
                                                total_vendas = df_display['Vendas Luck Sem Adicionais'].apply(extrair_valor_vendas).sum()
                                                total_vendas_formatado = f"R$ {total_vendas:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                
                                                col1, col2, col3, col4, col5 = st.columns(5)
                                                
                                                with col1:
                                                    st.markdown(f"""
                                                    <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin-top: 20px;">
                                                        <h3 style="margin: 0; color: #0e1117;">💰 Total de Vendas Luck Sem Adicionais</h3>
                                                        <h2 style="margin: 10px 0 0 0; color: #1f77b4;">{total_vendas_formatado}</h2>
                                                    </div>
                                                    """, unsafe_allow_html=True)
                                                
                                                with col2:
                                                    # Calcular total de Paxs In
                                                    if 'Paxs In' in df_display.columns:
                                                        try:
                                                            def extrair_valor_paxs(valor_str):
                                                                try:
                                                                    if pd.isna(valor_str):
                                                                        return 0.0
                                                                    valor_limpo = str(valor_str).replace(',', '.').strip()
                                                                    return float(valor_limpo) if valor_limpo else 0.0
                                                                except:
                                                                    return 0.0
                                                            
                                                            total_paxs = df_display['Paxs In'].apply(extrair_valor_paxs).sum()
                                                            total_paxs_formatado = f"{total_paxs:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                            
                                                            st.markdown(f"""
                                                            <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin-top: 20px;">
                                                                <h3 style="margin: 0; color: #0e1117;">👥 Total de Paxs In</h3>
                                                                <h2 style="margin: 10px 0 0 0; color: #1f77b4;">{total_paxs_formatado}</h2>
                                                            </div>
                                                            """, unsafe_allow_html=True)
                                                        except Exception as e:
                                                            st.error(f"Erro ao calcular total de Paxs In: {e}")
                                                
                                                with col3:
                                                    # Calcular Ticket Médio
                                                    if 'Paxs In' in df_display.columns:
                                                        try:
                                                            if total_paxs > 0:
                                                                ticket_medio = total_vendas / total_paxs
                                                                ticket_medio_formatado = f"R$ {ticket_medio:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                            else:
                                                                ticket_medio_formatado = "R$ 0,00"
                                                            
                                                            st.markdown(f"""
                                                            <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin-top: 20px;">
                                                                <h3 style="margin: 0; color: #0e1117;">🎟️ Ticket Médio</h3>
                                                                <h2 style="margin: 10px 0 0 0; color: #1f77b4;">{ticket_medio_formatado}</h2>
                                                            </div>
                                                            """, unsafe_allow_html=True)
                                                        except Exception as e:
                                                            st.error(f"Erro ao calcular Ticket Médio: {e}")
                                                
                                                with col4:
                                                    # Pegar valor da Meta (primeiro valor da coluna)
                                                    if 'Meta' in df_display.columns:
                                                        try:
                                                            # Pegar o primeiro valor não nulo da coluna Meta
                                                            meta_valor = df_display['Meta'].iloc[0] if len(df_display) > 0 else "R$ 0,00"
                                                            
                                                            st.markdown(f"""
                                                            <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin-top: 20px;">
                                                                <h3 style="margin: 0; color: #0e1117;">🎯 Meta</h3>
                                                                <h2 style="margin: 10px 0 0 0; color: #1f77b4;">{meta_valor}</h2>
                                                            </div>
                                                            """, unsafe_allow_html=True)
                                                        except Exception as e:
                                                            st.error(f"Erro ao exibir Meta: {e}")
                                                
                                                with col5:
                                                    # Calcular Alcance da Meta (Ticket Médio / Meta)
                                                    if 'Meta' in df_display.columns and 'Paxs In' in df_display.columns:
                                                        try:
                                                            # Extrair valor numérico da Meta
                                                            meta_str = str(meta_valor).replace('R$', '').replace('.', '').replace(',', '.').strip()
                                                            meta_float = float(meta_str) if meta_str not in ['', '0', '0,00'] else 0
                                                            
                                                            if meta_float > 0:
                                                                alcance_meta = (ticket_medio / meta_float) * 100
                                                                alcance_meta_formatado = f"{alcance_meta:.2f}%".replace('.', ',')
                                                            else:
                                                                alcance_meta_formatado = "0,00%"
                                                            
                                                            st.markdown(f"""
                                                            <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin-top: 20px;">
                                                                <h3 style="margin: 0; color: #0e1117;">📊 Alcance da Meta</h3>
                                                                <h2 style="margin: 10px 0 0 0; color: #1f77b4;">{alcance_meta_formatado}</h2>
                                                            </div>
                                                            """, unsafe_allow_html=True)
                                                        except Exception as e:
                                                            st.error(f"Erro ao calcular Alcance da Meta: {e}")
                                            except Exception as e:
                                                st.error(f"Erro ao calcular totais: {e}")
                                        
                                        # ==================== SEGUNDO GRID PARA TRANSFERISTAS E GUIAS ====================
                                        if tipo in ['Transferistas', 'Guias']:
                                            st.markdown("---")
                                            st.subheader(f"📋 Grid All Inclusive - {tipo}")
                                            
                                            # Criar grid simplificado com apenas Vendedor e Tipo de Vendedor
                                            df_simples = df_tipo[['Nome Do Vendedor', 'Tipo de Vendedor']].drop_duplicates().reset_index(drop=True)
                                            df_simples = df_simples.rename(columns={'Nome Do Vendedor': 'Vendedor'})
                                            
                                            # Adicionar coluna "Vendas Luck Sem Adicionais All Inclusive"
                                            if not df_vendas.empty:
                                                try:
                                                    vendedores_list = df_simples['Vendedor'].tolist()
                                                    vendas_luck_ai = calcular_vendas_luck_all_inclusive(
                                                        df_vendas, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final
                                                    )
                                                    
                                                    # Adicionar coluna ao dataframe
                                                    df_simples['Vendas Luck Sem Adicionais All Inclusive'] = df_simples['Vendedor'].map(vendas_luck_ai).fillna(0)
                                                    
                                                    # Formatar valores como moeda
                                                    df_simples['Vendas Luck Sem Adicionais All Inclusive'] = df_simples['Vendas Luck Sem Adicionais All Inclusive'].apply(
                                                        lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                    )
                                                    
                                                except Exception as e:
                                                    st.error(f"Erro ao calcular Vendas Luck All Inclusive: {e}")
                                            
                                            # Adicionar coluna "Paxs In All Inclusive"
                                            if not df_paxs_in.empty:
                                                try:
                                                    vendedores_list = df_simples['Vendedor'].tolist()
                                                    paxs_in_ai = calcular_paxs_in_all_inclusive(
                                                        df_paxs_in, vendedores_list, dia_inicial, mes_inicial, ano_inicial, dia_final, mes_final, ano_final
                                                    )
                                                    
                                                    # Adicionar coluna ao dataframe
                                                    df_simples['Paxs In All Inclusive'] = df_simples['Vendedor'].map(paxs_in_ai).fillna(0)
                                                    
                                                    # Formatar valores como decimal com 1 casa decimal (corrigindo divisão por 100)
                                                    df_simples['Paxs In All Inclusive'] = df_simples['Paxs In All Inclusive'].apply(
                                                        lambda x: f"{float(x)/100:.1f}".replace('.', ',') if x > 0 else "0"
                                                    )
                                                    
                                                except Exception as e:
                                                    st.error(f"Erro ao calcular Paxs In All Inclusive: {e}")
                                            
                                            # Adicionar coluna "Ticket Médio All Inclusive"
                                            if 'Vendas Luck Sem Adicionais All Inclusive' in df_simples.columns and 'Paxs In All Inclusive' in df_simples.columns:
                                                try:
                                                    def calcular_ticket_medio_ai(row):
                                                        try:
                                                            # Converter Vendas Luck Sem Adicionais All Inclusive para float
                                                            vendas_str = str(row['Vendas Luck Sem Adicionais All Inclusive']).replace('R$', '').replace('.', '').replace(',', '.').strip()
                                                            vendas_float = float(vendas_str) if vendas_str not in ['', '0', '0,00'] else 0
                                                            
                                                            # Converter Paxs In All Inclusive para float
                                                            paxs_str = str(row['Paxs In All Inclusive']).replace(',', '.').strip()
                                                            paxs_float = float(paxs_str) if paxs_str not in ['', '0', '0,0'] else 0
                                                            
                                                            # Calcular ticket médio All Inclusive
                                                            if paxs_float > 0:
                                                                ticket_medio_ai = vendas_float / paxs_float
                                                                return f"R$ {ticket_medio_ai:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                            else:
                                                                return "R$ 0,00"
                                                        except:
                                                            return "R$ 0,00"
                                                    
                                                    # Aplicar cálculo de ticket médio All Inclusive
                                                    df_simples['Ticket Médio All Inclusive'] = df_simples.apply(calcular_ticket_medio_ai, axis=1)
                                                    
                                                except Exception as e:
                                                    st.error(f"Erro ao calcular Ticket Médio All Inclusive: {e}")
                                            
                                            # Adicionar coluna "Meta All Inclusive" apenas para Transferistas e Guias
                                            try:
                                                def buscar_meta_ai_por_vendedor(row):
                                                    vendedor = row['Vendedor']
                                                    # Busca e soma metas da coluna "Meta All Inclusive" na aba Vendedores
                                                    df_vendedor = df_vendedores[df_vendedores['Nome Do Vendedor'] == vendedor].copy()
                                                    df_vendedor['Ano'] = pd.to_numeric(df_vendedor['Ano'], errors='coerce')
                                                    df_vendedor['mês'] = pd.to_numeric(df_vendedor['mês'], errors='coerce')
                                                    df_vendedor = df_vendedor.dropna(subset=['Ano', 'mês'])
                                                    df_periodo = df_vendedor[
                                                        (df_vendedor['Ano'] >= ano_inicial) & 
                                                        (df_vendedor['Ano'] <= ano_final) &
                                                        (df_vendedor['mês'] >= mes_inicial) & 
                                                        (df_vendedor['mês'] <= mes_final)
                                                    ]
                                                    if 'Meta All Inclusive' in df_periodo.columns:
                                                        metas_ai = df_periodo['Meta All Inclusive'].apply(formatar_meta)
                                                        meta_total_ai = metas_ai.sum()
                                                        if meta_total_ai > 0:
                                                            return f"R$ {float(meta_total_ai):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                        else:
                                                            return "R$ 0,00"
                                                    else:
                                                        return "R$ 0,00"
                                                # Aplicar busca de meta all inclusive
                                                df_simples['Meta All Inclusive'] = df_simples.apply(buscar_meta_ai_por_vendedor, axis=1)
                                            except Exception as e:
                                                st.error(f"Erro ao buscar Meta All Inclusive: {e}")

                                            # Adicionar coluna "Alcance de Meta All Inclusive" apenas para Transferistas e Guias
                                            try:
                                                def calcular_alcance_meta_ai(row):
                                                    ticket_str = str(row.get('Ticket Médio All Inclusive', 'R$ 0,00')).replace('R$', '').replace('.', '').replace(',', '.').strip()
                                                    meta_str = str(row.get('Meta All Inclusive', 'R$ 0,00')).replace('R$', '').replace('.', '').replace(',', '.').strip()
                                                    try:
                                                        ticket = float(ticket_str) if ticket_str not in ['', '0', '0,00'] else 0
                                                        meta = float(meta_str) if meta_str not in ['', '0', '0,00'] else 0
                                                        if meta > 0:
                                                            alcance = ticket / meta
                                                            return f"{alcance:.2%}".replace('.', ',')
                                                        else:
                                                            return "0,00%"
                                                    except:
                                                        return "0,00%"
                                                # Aplicar cálculo de alcance de meta all inclusive
                                                df_simples['Alcance de Meta All Inclusive'] = df_simples.apply(calcular_alcance_meta_ai, axis=1)
                                            except Exception as e:
                                                st.error(f"Erro ao calcular Alcance de Meta All Inclusive: {e}")

                                            # Adicionar coluna "Premiação All Inclusive" apenas para Transferistas
                                            if tipo == 'Transferistas':
                                                try:
                                                    def calcular_premiacao_ai_grid(row):
                                                        try:
                                                            # Calcular baseado no Alcance de Meta All Inclusive atual (sem mapeamento fixo)
                                                            alcance_str = str(row.get('Alcance de Meta All Inclusive', '0,00%')).replace('%', '').replace(',', '.')
                                                            alcance_float = float(alcance_str) if alcance_str not in ['', '0', '0,00'] else 0
                                                            
                                                            # Aplicar lógica de premiação All Inclusive baseada no alcance real
                                                            if alcance_float >= 150.0:
                                                                return '5%'
                                                            elif alcance_float >= 120.0:
                                                                return '4%'
                                                            elif alcance_float >= 100.0:
                                                                return '3%'
                                                            elif alcance_float >= 90.0:
                                                                return '2%'
                                                            elif alcance_float >= 80.0:
                                                                return '1%'
                                                            else:
                                                                return '0%'
                                                                
                                                        except Exception:
                                                            return '0%'
                                                    
                                                    # Aplicar cálculo de premiação All Inclusive
                                                    df_simples['Premiação All Inclusive'] = df_simples.apply(calcular_premiacao_ai_grid, axis=1)
                                                    
                                                    # Armazenar valores de Premiação All Inclusive em variável global para uso no Grid Detalhes
                                                    globals()['premiacao_ai_por_vendedor'] = dict(zip(df_simples['Vendedor'], df_simples['Premiação All Inclusive']))
                                                except Exception as e:
                                                    st.error(f"Erro ao calcular Premiação All Inclusive: {e}")
                                            st.write(f"**Total de vendedores:** {len(df_simples)}")
                                            # Mostrar grid simplificado
                                            def highlight_alcance_meta_ai(row):
                                                valor = str(row.get('Alcance de Meta All Inclusive', '0,00%')).replace('%', '').replace(',', '.')
                                                try:
                                                    valor_float = float(valor)
                                                    if valor_float >= 100.0:
                                                        return ['background-color: #b6fcb6'] * len(row)
                                                    else:
                                                        return [''] * len(row)
                                                except:
                                                    return [''] * len(row)

                                            try:
                                                st.dataframe(
                                                    df_simples.style.apply(highlight_alcance_meta_ai, axis=1),
                                                    use_container_width=True,
                                                    hide_index=True
                                                )
                                            except Exception as e:
                                                st.dataframe(
                                                    df_simples,
                                                    use_container_width=True,
                                                    hide_index=True
                                                )
                                            
                                            # Cartão com soma de Vendas Luck Sem Adicionais All Inclusive
                                            if 'Vendas Luck Sem Adicionais All Inclusive' in df_simples.columns:
                                                try:
                                                    # Extrair valores numéricos e somar
                                                    def extrair_valor_vendas_ai(valor_str):
                                                        try:
                                                            if pd.isna(valor_str):
                                                                return 0.0
                                                            valor_limpo = str(valor_str).replace('R$', '').replace('.', '').replace(',', '.').strip()
                                                            return float(valor_limpo) if valor_limpo else 0.0
                                                        except:
                                                            return 0.0
                                                    
                                                    total_vendas_ai = df_simples['Vendas Luck Sem Adicionais All Inclusive'].apply(extrair_valor_vendas_ai).sum()
                                                    total_vendas_ai_formatado = f"R$ {total_vendas_ai:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                    
                                                    col1_ai, col2_ai, col3_ai, col4_ai, col5_ai = st.columns(5)
                                                    
                                                    with col1_ai:
                                                        st.markdown(f"""
                                                        <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin-top: 20px;">
                                                            <h3 style="margin: 0; color: #0e1117;">💰 Total de Vendas Luck Sem Adicionais All Inclusive</h3>
                                                            <h2 style="margin: 10px 0 0 0; color: #1f77b4;">{total_vendas_ai_formatado}</h2>
                                                        </div>
                                                        """, unsafe_allow_html=True)
                                                    
                                                    with col2_ai:
                                                        # Calcular total de Paxs In All Inclusive
                                                        if 'Paxs In All Inclusive' in df_simples.columns:
                                                            try:
                                                                def extrair_valor_paxs_ai(valor_str):
                                                                    try:
                                                                        if pd.isna(valor_str):
                                                                            return 0.0
                                                                        valor_limpo = str(valor_str).replace(',', '.').strip()
                                                                        return float(valor_limpo) if valor_limpo else 0.0
                                                                    except:
                                                                        return 0.0
                                                                
                                                                total_paxs_ai = df_simples['Paxs In All Inclusive'].apply(extrair_valor_paxs_ai).sum()
                                                                total_paxs_ai_formatado = f"{total_paxs_ai:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                                
                                                                st.markdown(f"""
                                                                <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin-top: 20px;">
                                                                    <h3 style="margin: 0; color: #0e1117;">👥 Total de Paxs In All Inclusive</h3>
                                                                    <h2 style="margin: 10px 0 0 0; color: #1f77b4;">{total_paxs_ai_formatado}</h2>
                                                                </div>
                                                                """, unsafe_allow_html=True)
                                                            except Exception as e:
                                                                st.error(f"Erro ao calcular total de Paxs In All Inclusive: {e}")
                                                    
                                                    with col3_ai:
                                                        # Calcular Ticket Médio All Inclusive
                                                        if 'Paxs In All Inclusive' in df_simples.columns:
                                                            try:
                                                                if total_paxs_ai > 0:
                                                                    ticket_medio_ai = total_vendas_ai / total_paxs_ai
                                                                    ticket_medio_ai_formatado = f"R$ {ticket_medio_ai:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                                                                else:
                                                                    ticket_medio_ai_formatado = "R$ 0,00"
                                                                
                                                                st.markdown(f"""
                                                                <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin-top: 20px;">
                                                                    <h3 style="margin: 0; color: #0e1117;">🎟️ Ticket Médio All Inclusive</h3>
                                                                    <h2 style="margin: 10px 0 0 0; color: #1f77b4;">{ticket_medio_ai_formatado}</h2>
                                                                </div>
                                                                """, unsafe_allow_html=True)
                                                            except Exception as e:
                                                                st.error(f"Erro ao calcular Ticket Médio All Inclusive: {e}")
                                                    
                                                    with col4_ai:
                                                        # Pegar valor da Meta All Inclusive (primeiro valor da coluna)
                                                        if 'Meta All Inclusive' in df_simples.columns:
                                                            try:
                                                                # Pegar o primeiro valor não nulo da coluna Meta All Inclusive
                                                                meta_ai_valor = df_simples['Meta All Inclusive'].iloc[0] if len(df_simples) > 0 else "R$ 0,00"
                                                                
                                                                st.markdown(f"""
                                                                <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin-top: 20px;">
                                                                    <h3 style="margin: 0; color: #0e1117;">🎯 Meta All Inclusive</h3>
                                                                    <h2 style="margin: 10px 0 0 0; color: #1f77b4;">{meta_ai_valor}</h2>
                                                                </div>
                                                                """, unsafe_allow_html=True)
                                                            except Exception as e:
                                                                st.error(f"Erro ao exibir Meta All Inclusive: {e}")
                                                    
                                                    with col5_ai:
                                                        # Calcular Alcance da Meta All Inclusive (Ticket Médio AI / Meta AI)
                                                        if 'Meta All Inclusive' in df_simples.columns and 'Paxs In All Inclusive' in df_simples.columns:
                                                            try:
                                                                # Extrair valor numérico da Meta All Inclusive
                                                                meta_ai_str = str(meta_ai_valor).replace('R$', '').replace('.', '').replace(',', '.').strip()
                                                                meta_ai_float = float(meta_ai_str) if meta_ai_str not in ['', '0', '0,00'] else 0
                                                                
                                                                if meta_ai_float > 0:
                                                                    alcance_meta_ai_calc = (ticket_medio_ai / meta_ai_float) * 100
                                                                    alcance_meta_ai_formatado = f"{alcance_meta_ai_calc:.2f}%".replace('.', ',')
                                                                else:
                                                                    alcance_meta_ai_formatado = "0,00%"
                                                                
                                                                st.markdown(f"""
                                                                <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin-top: 20px;">
                                                                    <h3 style="margin: 0; color: #0e1117;">📊 Alcance da Meta All Inclusive</h3>
                                                                    <h2 style="margin: 10px 0 0 0; color: #1f77b4;">{alcance_meta_ai_formatado}</h2>
                                                                </div>
                                                                """, unsafe_allow_html=True)
                                                            except Exception as e:
                                                                st.error(f"Erro ao calcular Alcance da Meta All Inclusive: {e}")
                                                except Exception as e:
                                                    st.error(f"Erro ao calcular total de vendas All Inclusive: {e}")

                                            # ========== TERCEIRO GRID: COMISSÃO ==========
                                            st.markdown("---")
                                            st.subheader(f"💰 Comissão - {tipo}")

                                            # Pegar vendedores dos dois grids
                                            vendedores_tipo = df_display['Vendedor'].tolist() if 'Vendedor' in df_display.columns else []
                                            vendedores_ai = df_simples['Vendedor'].tolist() if 'Vendedor' in df_simples.columns else []
                                            vendedores_comissao = list(set(vendedores_tipo + vendedores_ai))
                                            
                                            # Criar grid de detalhes da comissão (apenas para Transferistas e Guias)
                                            if not df_comissao.empty and vendedores_comissao:
                                                st.subheader(f"📋 Detalhes da Comissão - {tipo}")
                                                
                                                # Filtrar dados de comissão por período e vendedores
                                                # Passar df_vendas globalmente para uso na função de busca All Inclusive
                                                globals()['df_vendas'] = df_vendas
                                                comissao_detalhes = filtrar_comissao_por_periodo_vendedor(
                                                    df_comissao, 
                                                    vendedores_comissao,
                                                    dia_inicial, mes_inicial, ano_inicial, 
                                                    dia_final, mes_final, ano_final
                                                )
                                                
                                                if not comissao_detalhes.empty:
                                                    # Para Guias, remover as colunas Premiação e Premiação All Inclusive
                                                    if tipo == 'Guias':
                                                        colunas_ocultar = ['Premiação', 'Premiação All Inclusive', 'Valor Comissão Premiação', 'Valor Comissão Premiação All Inclusive']
                                                        comissao_display = comissao_detalhes.drop(columns=[col for col in colunas_ocultar if col in comissao_detalhes.columns])
                                                    else:
                                                        comissao_display = comissao_detalhes
                                                    
                                                    st.dataframe(
                                                        comissao_display,
                                                        use_container_width=True,
                                                        hide_index=True
                                                    )
                                                    st.info(f"📊 Total de registros de comissão: {len(comissao_detalhes)}")
                                                    
                                                    # Resumo por vendedor
                                                    if len(comissao_detalhes) > 0:
                                                        st.markdown("---")
                                                        st.subheader(f"📈 Resumo de Comissão por Vendedor - {tipo}")
                                                        
                                                        # Criar coluna numérica temporária para soma
                                                        def extrair_valor_numerico(valor_str):
                                                            try:
                                                                if pd.isna(valor_str):
                                                                    return 0.0
                                                                valor_limpo = str(valor_str).replace('R$', '').replace('.', '').replace(',', '.').strip()
                                                                return float(valor_limpo) if valor_limpo else 0.0
                                                            except:
                                                                return 0.0
                                                        
                                                        comissao_detalhes['Valor da Venda Numerico'] = comissao_detalhes['Valor da Venda'].apply(extrair_valor_numerico)
                                                        comissao_detalhes['Valor Comissão Luck Numerico'] = comissao_detalhes['Valor Comissão Luck'].apply(extrair_valor_numerico)
                                                        comissao_detalhes['Valor Comissão Terceiros Numerico'] = comissao_detalhes['Valor Comissão Terceiros'].apply(extrair_valor_numerico)
                                                        comissao_detalhes['Valor Comissão Premiação Numerico'] = comissao_detalhes['Valor Comissão Premiação'].apply(extrair_valor_numerico)
                                                        comissao_detalhes['Valor Comissão Premiação AI Numerico'] = comissao_detalhes['Valor Comissão Premiação All Inclusive'].apply(extrair_valor_numerico)
                                                        comissao_detalhes['Valor Total de Comissão Numerico'] = comissao_detalhes['Valor Total de Comissão'].apply(extrair_valor_numerico)
                                                        
                                                        resumo_vendedor = comissao_detalhes.groupby('Vendedor').agg({
                                                            'Valor da Venda Numerico': 'sum',
                                                            'Valor Comissão Luck Numerico': 'sum',
                                                            'Valor Comissão Terceiros Numerico': 'sum',
                                                            'Valor Comissão Premiação Numerico': 'sum',
                                                            'Valor Comissão Premiação AI Numerico': 'sum',
                                                            'Valor Total de Comissão Numerico': 'sum'
                                                        }).reset_index()
                                                        
                                                        resumo_vendedor.columns = ['Vendedor', 'Valor Total de Venda', 'Valor Total Comissão Luck', 'Valor Total Comissão Terceiros', 'Valor Total Comissão Premiação', 'Valor Total Comissão Premiação All Inclusive', 'Valor Total de Comissão']
                                                        
                                                        # Formatar valores como moeda
                                                        resumo_vendedor['Valor Total de Venda'] = resumo_vendedor['Valor Total de Venda'].apply(
                                                            lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notna(x) else 'R$ 0,00'
                                                        )
                                                        resumo_vendedor['Valor Total Comissão Luck'] = resumo_vendedor['Valor Total Comissão Luck'].apply(
                                                            lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notna(x) else 'R$ 0,00'
                                                        )
                                                        resumo_vendedor['Valor Total Comissão Terceiros'] = resumo_vendedor['Valor Total Comissão Terceiros'].apply(
                                                            lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notna(x) else 'R$ 0,00'
                                                        )
                                                        resumo_vendedor['Valor Total Comissão Premiação'] = resumo_vendedor['Valor Total Comissão Premiação'].apply(
                                                            lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notna(x) else 'R$ 0,00'
                                                        )
                                                        resumo_vendedor['Valor Total Comissão Premiação All Inclusive'] = resumo_vendedor['Valor Total Comissão Premiação All Inclusive'].apply(
                                                            lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notna(x) else 'R$ 0,00'
                                                        )
                                                        resumo_vendedor['Valor Total de Comissão'] = resumo_vendedor['Valor Total de Comissão'].apply(
                                                            lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notna(x) else 'R$ 0,00'
                                                        )
                                                        
                                                        # Ocultar colunas de premiação para Guias
                                                        if tipo == 'Guias':
                                                            colunas_ocultar_resumo = ['Valor Total Comissão Premiação', 'Valor Total Comissão Premiação All Inclusive']
                                                            resumo_display = resumo_vendedor.drop(columns=[col for col in colunas_ocultar_resumo if col in resumo_vendedor.columns])
                                                        else:
                                                            resumo_display = resumo_vendedor
                                                        
                                                        st.dataframe(
                                                            resumo_display,
                                                            use_container_width=True,
                                                            hide_index=True
                                                        )
                                                else:
                                                    st.info(f"📋 Nenhuma comissão encontrada para {tipo} no período selecionado")
                                            else:
                                                # Grid simples como fallback
                                                df_comissao_simples = pd.DataFrame({'Vendedor': vendedores_comissao}) if vendedores_comissao else pd.DataFrame()
                                                if not df_comissao_simples.empty:
                                                    st.write(f"**Total de vendedores:** {len(df_comissao_simples)}")
                                                    st.dataframe(
                                                        df_comissao_simples,
                                                        use_container_width=True,
                                                        hide_index=True
                                                    )

                                            # ========== GRÁFICOS DE TICKET MÉDIO ========== 
                                            import matplotlib.pyplot as plt
                                            import matplotlib.cm as cm

                                            # Título do período
                                            periodo_titulo = f"Período: {dia_inicial:02d}/{mes_inicial:02d}/{ano_inicial} a {dia_final:02d}/{mes_final:02d}/{ano_final}"

                                            # Gráfico 1: Ticket Médio
                                            if 'Ticket Médio' in df_display.columns:
                                                st.markdown("---")
                                                st.subheader(f"🎟️ Ticket Médio - {tipo} ({periodo_titulo})")
                                                vendedores = df_display['Vendedor'].tolist()
                                                valores = [float(str(v).replace('R$', '').replace('.', '').replace(',', '.').strip()) if str(v) not in ['', '0', '0,00'] else 0 for v in df_display['Ticket Médio']]
                                                largura = max(8, len(vendedores) * 0.6)
                                                fig, ax = plt.subplots(figsize=(largura, 4))
                                                colors = cm.get_cmap('tab10', len(vendedores))
                                                bars = ax.bar(vendedores, valores, color=[colors(i) for i in range(len(vendedores))])
                                                ax.set_ylabel('Ticket Médio (R$)')
                                                ax.set_xlabel('Vendedor')
                                                ax.set_title(f'Ticket Médio por Vendedor - {tipo}')
                                                ax.set_xticklabels(vendedores, rotation=45, ha='right')
                                                # Adicionar legenda com valor
                                                for bar, valor in zip(bars, valores):
                                                    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height(), f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'), ha='center', va='bottom', fontsize=9)
                                                st.pyplot(fig)

                                            # Gráfico 2: Ticket Médio All Inclusive
                                            if 'Ticket Médio All Inclusive' in df_simples.columns:
                                                st.markdown("---")
                                                st.subheader(f"🎟️ Ticket Médio All Inclusive - {tipo} ({periodo_titulo})")
                                                vendedores_ai = df_simples['Vendedor'].tolist()
                                                valores_ai = [float(str(v).replace('R$', '').replace('.', '').replace(',', '.').strip()) if str(v) not in ['', '0', '0,00'] else 0 for v in df_simples['Ticket Médio All Inclusive']]
                                                largura_ai = max(8, len(vendedores_ai) * 0.6)
                                                fig2, ax2 = plt.subplots(figsize=(largura_ai, 4))
                                                colors_ai = cm.get_cmap('tab20', len(vendedores_ai))
                                                bars_ai = ax2.bar(vendedores_ai, valores_ai, color=[colors_ai(i) for i in range(len(vendedores_ai))])
                                                ax2.set_ylabel('Ticket Médio All Inclusive (R$)')
                                                ax2.set_xlabel('Vendedor')
                                                ax2.set_title(f'Ticket Médio All Inclusive por Vendedor - {tipo}')
                                                ax2.set_xticklabels(vendedores_ai, rotation=45, ha='right')
                                                # Adicionar legenda com valor
                                                for bar, valor in zip(bars_ai, valores_ai):
                                                    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height(), f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'), ha='center', va='bottom', fontsize=9)
                                                st.pyplot(fig2)
                                        
                                        # ========== ARMAZENAR DADOS NO SESSION STATE ==========
                                        # Salvar dados dos grids para geração de relatórios
                                        if 'dados_relatorios' not in st.session_state:
                                            st.session_state['dados_relatorios'] = {}
                                        
                                        if tipo not in st.session_state['dados_relatorios']:
                                            st.session_state['dados_relatorios'][tipo] = {}
                                        
                                        # Armazenar dados do Grid 1 (Vendas Luck)
                                        for idx, row in df_display.iterrows():
                                            vendedor_nome = row['Vendedor']
                                            if vendedor_nome not in st.session_state['dados_relatorios'][tipo]:
                                                st.session_state['dados_relatorios'][tipo][vendedor_nome] = {}
                                            
                                            st.session_state['dados_relatorios'][tipo][vendedor_nome]['grid1'] = row.to_dict()
                                        
                                        # Armazenar dados do Grid 2 (All Inclusive)
                                        if 'df_simples' in locals() and not df_simples.empty:
                                            for idx, row in df_simples.iterrows():
                                                vendedor_nome = row['Vendedor']
                                                if vendedor_nome not in st.session_state['dados_relatorios'][tipo]:
                                                    st.session_state['dados_relatorios'][tipo][vendedor_nome] = {}
                                                
                                                st.session_state['dados_relatorios'][tipo][vendedor_nome]['grid2'] = row.to_dict()
                                        
                                        # Armazenar dados do Resumo de Comissão
                                        if 'resumo_vendedor' in locals() and not resumo_vendedor.empty:
                                            for idx, row in resumo_vendedor.iterrows():
                                                vendedor_nome = row['Vendedor']
                                                if vendedor_nome not in st.session_state['dados_relatorios'][tipo]:
                                                    st.session_state['dados_relatorios'][tipo][vendedor_nome] = {}
                                                
                                                st.session_state['dados_relatorios'][tipo][vendedor_nome]['resumo'] = row.to_dict()
                                        
                                        # Armazenar dados do Grid Detalhes da Comissão
                                        if 'comissao_detalhes' in locals() and not comissao_detalhes.empty:
                                            for vendedor_nome in comissao_detalhes['Vendedor'].unique():
                                                if vendedor_nome not in st.session_state['dados_relatorios'][tipo]:
                                                    st.session_state['dados_relatorios'][tipo][vendedor_nome] = {}
                                                
                                                vendedor_detalhes = comissao_detalhes[comissao_detalhes['Vendedor'] == vendedor_nome]
                                                st.session_state['dados_relatorios'][tipo][vendedor_nome]['detalhes'] = vendedor_detalhes
                                        
                                        # Armazenar informações do período
                                        st.session_state['periodo_texto'] = f"{dia_inicial:02d}/{mes_inicial:02d}/{ano_inicial} a {dia_final:02d}/{mes_final:02d}/{ano_final}"
                    
                    # ========== SEÇÃO DE GERAÇÃO DE RELATÓRIOS ==========
                    if 'dados_relatorios' in st.session_state and st.session_state['dados_relatorios']:
                        st.markdown("---")
                        st.markdown("---")
                        st.subheader("📄 Geração de Relatórios em PDF")
                        st.info("💡 Gere relatórios individuais por vendedor sem recarregar os dados")
                        
                        col_rel1, col_rel2, col_rel3, col_rel4 = st.columns(4)
                        
                        with col_rel1:
                            # Filtro de Tipo de Vendedor (obrigatório)
                            tipos_disponiveis = list(st.session_state['dados_relatorios'].keys())
                            tipo_vendedor_filtro = st.selectbox(
                                "Tipo de Vendedor *",
                                options=sorted(tipos_disponiveis),
                                key="tipo_vendedor_relatorio"
                            )
                        
                        with col_rel2:
                            # Coletar vendedores do tipo selecionado
                            vendedores_do_tipo = []
                            if tipo_vendedor_filtro in st.session_state['dados_relatorios']:
                                vendedores_do_tipo = list(st.session_state['dados_relatorios'][tipo_vendedor_filtro].keys())
                            
                            # Adicionar "Todos" como primeira opção
                            opcoes_vendedor = ["Todos"] + sorted(vendedores_do_tipo)
                            
                            vendedor_selecionado = st.selectbox(
                                "Selecione o Vendedor",
                                options=opcoes_vendedor,
                                key="vendedor_relatorio"
                            )
                        
                        with col_rel3:
                            tipo_relatorio = st.selectbox(
                                "Tipo de Relatório",
                                options=["Estatístico", "Comissão"],
                                key="tipo_relatorio"
                            )
                        
                        with col_rel4:
                            st.write("")
                            st.write("")
                            if st.button("📥 Gerar PDF", type="primary", key="gerar_pdf_btn"):
                                periodo_texto = st.session_state.get('periodo_texto', 'Período não especificado')
                                
                                # Determinar lista de vendedores para processar
                                if vendedor_selecionado == "Todos":
                                    vendedores_para_processar = vendedores_do_tipo
                                else:
                                    vendedores_para_processar = [vendedor_selecionado]
                                
                                # Armazenar PDFs gerados para download em lote
                                pdfs_gerados = []
                                
                                # Gerar PDF para cada vendedor
                                for vendedor_atual in vendedores_para_processar:
                                    # Buscar dados do vendedor no tipo selecionado
                                    dados_vendedor = st.session_state['dados_relatorios'][tipo_vendedor_filtro].get(vendedor_atual, None)
                                    
                                    if dados_vendedor:
                                        try:
                                            if tipo_relatorio == "Estatístico":
                                                # Gerar PDF Estatístico
                                                pdf_buffer = gerar_pdf_estatistico(
                                                    vendedor=vendedor_atual,
                                                    periodo_texto=periodo_texto,
                                                    dados_grid1=dados_vendedor.get('grid1', {}),
                                                    dados_grid2=dados_vendedor.get('grid2', {}),
                                                    dados_resumo=dados_vendedor.get('resumo', {})
                                                )
                                                
                                                # Formatar período para nome do arquivo (sem barras)
                                                periodo_arquivo = periodo_texto.replace('/', '-').replace(' ', '_')
                                                nome_arquivo = f"Relatorio_Estatistico_{vendedor_atual.replace(' ', '_')}_{periodo_arquivo}.pdf"
                                                
                                                # Armazenar para download individual
                                                pdfs_gerados.append({
                                                    'vendedor': vendedor_atual,
                                                    'buffer': pdf_buffer,
                                                    'nome_arquivo': nome_arquivo,
                                                    'tipo': 'Estatístico'
                                                })
                                                
                                                st.download_button(
                                                    label=f"⬇️ Download Relatório Estatístico - {vendedor_atual}",
                                                    data=pdf_buffer,
                                                    file_name=nome_arquivo,
                                                    mime="application/pdf",
                                                    key=f"download_estatistico_{vendedor_atual}"
                                                )
                                                st.success(f"✅ Relatório Estatístico gerado com sucesso para {vendedor_atual}!")
                                            
                                            else:  # Comissão
                                                # Gerar PDF Comissão
                                                pdf_buffer = gerar_pdf_comissao(
                                                    vendedor=vendedor_atual,
                                                    periodo_texto=periodo_texto,
                                                    dados_detalhes=dados_vendedor.get('detalhes', None),
                                                    dados_resumo=dados_vendedor.get('resumo', {}),
                                                    tipo_vendedor=tipo_vendedor_filtro
                                                )
                                                
                                                # Formatar período para nome do arquivo (sem barras)
                                                periodo_arquivo = periodo_texto.replace('/', '-').replace(' ', '_')
                                                nome_arquivo = f"Relatorio_Comissao_{vendedor_atual.replace(' ', '_')}_{periodo_arquivo}.pdf"
                                                
                                                # Armazenar para download individual
                                                pdfs_gerados.append({
                                                    'vendedor': vendedor_atual,
                                                    'buffer': pdf_buffer,
                                                    'nome_arquivo': nome_arquivo,
                                                    'tipo': 'Comissão'
                                                })
                                                
                                                st.download_button(
                                                    label=f"⬇️ Download Relatório de Comissão - {vendedor_atual}",
                                                    data=pdf_buffer,
                                                    file_name=nome_arquivo,
                                                    mime="application/pdf",
                                                    key=f"download_comissao_{vendedor_atual}"
                                                )
                                                st.success(f"✅ Relatório de Comissão gerado com sucesso para {vendedor_atual}!")
                                        
                                        except Exception as e:
                                            st.error(f"❌ Erro ao gerar PDF para {vendedor_atual}: {str(e)}")
                                    else:
                                        st.warning(f"⚠️ Dados do vendedor {vendedor_atual} não encontrados")
                                
                                # Se gerou múltiplos PDFs, oferecer download em lote (ZIP)
                                if len(pdfs_gerados) > 1:
                                    st.markdown("---")
                                    st.info(f"📦 {len(pdfs_gerados)} relatórios gerados. Baixe todos de uma vez:")
                                    
                                    import zipfile
                                    zip_buffer = io.BytesIO()
                                    
                                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                        for pdf_info in pdfs_gerados:
                                            # Adicionar cada PDF ao ZIP
                                            zip_file.writestr(pdf_info['nome_arquivo'], pdf_info['buffer'].getvalue())
                                    
                                    zip_buffer.seek(0)
                                    
                                    # Formatar período para nome do arquivo ZIP
                                    periodo_arquivo = periodo_texto.replace('/', '-').replace(' ', '_')
                                    tipo_rel_nome = tipo_relatorio.replace(' ', '_')
                                    nome_zip = f"Relatorios_{tipo_rel_nome}_{tipo_vendedor_filtro}_{periodo_arquivo}.zip"
                                    
                                    st.download_button(
                                        label=f"📦 Download TODOS os {len(pdfs_gerados)} Relatórios (ZIP)",
                                        data=zip_buffer,
                                        file_name=nome_zip,
                                        mime="application/zip",
                                        key="download_todos_zip",
                                        type="primary"
                                    )
                    
                    # Nenhum tipo de vendedor encontrado - silencioso
                else:
                    st.error("❌ Colunas necessárias não encontradas na planilha!")
        else:
            st.error("❌ Colunas 'mês' e 'Ano' não encontradas na planilha!")
else:
    # Área para exibir os dados (placeholder)
    st.markdown("---")
    st.subheader("📈 Visualização de Dados")
    st.info("💡 Clique no botão 'Carregar Dados' para visualizar os grids por tipo de vendedor.")
