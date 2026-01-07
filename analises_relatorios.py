import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import warnings
warnings.filterwarnings('ignore')


# RELATÓRIOS INDIVIDUAIS EM PPT

# Carregando o dataframe processado
df_total = pd.read_csv('dados/processado_dados_avaliacao.csv')


# Configurações de texto
MAX_CARACTERES_POR_SLIDE = 2000  # Limite de caracteres por slide
TAMANHO_FONTE_TEXTO_LONGO = 10  # Tamanho da fonte para textos longos

# =================== CONFIGURAÇÕES ===================
DIRETORIO_BASE = r"D:\Scripts_Python\FGV\Avaliacoes_de_Desempenho_2025"
CAMINHO_TEMPLATE = os.path.join(DIRETORIO_BASE, "templates", "[GARN] Modelo de apresentação de slides.pptx")
DIRETORIO_SAIDA = os.path.join(DIRETORIO_BASE, "relatorios_gerados")
DIRETORIO_GRAFICOS_TEMP = os.path.join(DIRETORIO_BASE, "graficos_temp")

# Criar diretórios se não existirem
os.makedirs(DIRETORIO_SAIDA, exist_ok=True)
os.makedirs(DIRETORIO_GRAFICOS_TEMP, exist_ok=True)

# =================== DEFINIÇÕES DAS COLUNAS ===================
# Questões gerais
COLUNAS_GERAIS = [
    'Visao Sistemica_Identifica desafios',
    'Visao Sistematica_Compartilha ferramentas',
    'Visao Sistemica_Colabora com as demais frentes',
    'Visao Sistemica_Compreensao do impacto',
    'Gestao e Lideranca_Canal aberto',
    'Gestao e Lideranca_Incentiva e colabora',
    'Gestao e Lideranca_Autogestao de tempo',
    'Relacionamento_Cativa parceiros',
    'Relacionamento_Cria vinculos',
    'Comunicacao_Assertivo',
    'Comunicacao_Nao Violenta',
    'Comunicacao_Postura Profissional',
    'Aprendizagem_Compartilha experiencia',
    'Aprendizagem_Conhecimentos de fontes internas e externas',
    'Aprendizagem_Busca se desenvolver',
    'Aprendizagem_Autonomia',
    'Execucao_Resolve as dificuldades',
    'Execucao_Imprevistos e alteracoes',
    'Execucao_Ideias em Acoes'
]

# Questões de liderança
COLUNAS_LIDERANCA = [
    'Lideranca_Desenvolvimento de Pessoas',
    'Lideranca_Visão Estratégica',
    'Lideranca_Delegação',
    'Lideranca_Gerenciamento de Riscos',
    'Lideranca_Monitoramento de Resultados'
]

# Mapeamento de questões para placeholders gráficos
MAPEAMENTO_QUESTOES_GRAFICOS = {
    'Visao Sistemica_Identifica desafios': 'VS_1',
    'Visao Sistematica_Compartilha ferramentas': 'VS_2',
    'Visao Sistemica_Colabora com as demais frentes': 'VS_3',
    'Visao Sistemica_Compreensao do impacto': 'VS_4',
    'Gestao e Lideranca_Canal aberto': 'GL_1',
    'Gestao e Lideranca_Incentiva e colabora': 'GL_2',
    'Gestao e Lideranca_Autogestao de tempo': 'GL_3',
    'Relacionamento_Cativa parceiros': 'REL_1',
    'Relacionamento_Cria vinculos': 'REL_2',
    'Comunicacao_Assertivo': 'COM_1',
    'Comunicacao_Nao Violenta': 'COM_2',
    'Comunicacao_Postura Profissional': 'COM_3',
    'Aprendizagem_Compartilha experiencia': 'AD_1',
    'Aprendizagem_Conhecimentos de fontes internas e externas': 'AD_2',
    'Aprendizagem_Busca se desenvolver': 'AD_3',
    'Aprendizagem_Autonomia': 'AD_4',
    'Execucao_Resolve as dificuldades': 'EX_1',
    'Execucao_Imprevistos e alteracoes': 'EX_2',
    'Execucao_Ideias em Acoes': 'EX_3'
}

# Mapeamento de tipos para nomes amigáveis
TIPOS_AMIGAVEIS = {
    'autoavaliacao': 'Autoavaliação',
    'avaliacao_pelo_lider': 'Avaliação do Líder',
    'avaliacao_pelo_liderado': 'Avaliação dos Liderados'
}

CORES_TIPOS = {
    'autoavaliacao': '#00a6dc',  # Azul
    'avaliacao_pelo_lider': '#58a75b',  # Verde
    'avaliacao_pelo_liderado': '#cc8a42'  # Marrom
}

# =================== FUNÇÕES AUXILIARES ===================
def formatar_nota(valor):
    """Formata uma nota com 1 casa decimal"""
    if pd.isna(valor):
        return "Não se aplica"
    return f"{valor:.1f}".replace('.', ',')

def filtrar_dados_pessoa(df, nome):
    """Filtra o dataframe para uma pessoa específica"""
    return df[df['Nome'] == nome]

def calcular_medias_tipo(df_pessoa, tipo, colunas):
    """Calcula a média para um tipo específico e um conjunto de colunas"""
    df_tipo = df_pessoa[df_pessoa['Tipo'] == tipo]
    if df_tipo.empty:
        return None
    return df_tipo[colunas].mean()

def criar_grafico_barras_questao(df_pessoa, questao, tipo_placeholder):
    """Cria um gráfico de barras para uma questão específica"""
    # Coletar médias por tipo
    dados = {}
    for tipo in ['autoavaliacao', 'avaliacao_pelo_lider', 'avaliacao_pelo_liderado']:
        df_tipo = df_pessoa[df_pessoa['Tipo'] == tipo]
        if not df_tipo.empty and questao in df_tipo.columns:
            media = df_tipo[questao].mean()
            if not pd.isna(media):
                dados[tipo] = media
    
    if not dados:
        return None
    
    # Criar gráfico
    fig, ax = plt.subplots(figsize=(4, 3))
    tipos_presentes = list(dados.keys())
    valores = [dados[t] for t in tipos_presentes]
    cores = [CORES_TIPOS[t] for t in tipos_presentes]
    nomes = [TIPOS_AMIGAVEIS[t] for t in tipos_presentes]
    
    bars = ax.bar(range(len(tipos_presentes)), valores, color=cores)
    ax.set_ylim(0, 6)  # Escala de 0-6
    
    # Adicionar valores nas barras
    for bar, valor in zip(bars, valores):
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height,
                f'{valor:.1f}', ha='center', va='bottom')
    
    ax.set_xticks(range(len(tipos_presentes)))
    ax.set_xticklabels(nomes, rotation=45, ha='right')
    ax.set_ylabel('Média')
    ax.set_title(f'{questao.split("_")[0]}', fontsize=10)
    
    plt.tight_layout()
    
    # Salvar gráfico
    caminho_grafico = os.path.join(DIRETORIO_GRAFICOS_TEMP, f"{tipo_placeholder}_{id(df_pessoa)}.png")
    fig.savefig(caminho_grafico, dpi=150, bbox_inches='tight')
    plt.close(fig)
    
    return caminho_grafico

def criar_radar_geral(df_pessoa):
    """Cria gráfico de radar para competências gerais"""
    # Definir categorias e questões
    categorias = {
        'Visão Sistêmica': [
            'Visao Sistemica_Identifica desafios',
            'Visao Sistematica_Compartilha ferramentas',
            'Visao Sistemica_Colabora com as demais frentes',
            'Visao Sistemica_Compreensao do impacto'
        ],
        'Gestão e Liderança': [
            'Gestao e Lideranca_Canal aberto',
            'Gestao e Lideranca_Incentiva e colabora',
            'Gestao e Lideranca_Autogestao de tempo'
        ],
        'Relacionamento': [
            'Relacionamento_Cativa parceiros',
            'Relacionamento_Cria vinculos'
        ],
        'Comunicação': [
            'Comunicacao_Assertivo',
            'Comunicacao_Nao Violenta',
            'Comunicacao_Postura Profissional'
        ],
        'Aprendizagem e Desenvolvimento': [
            'Aprendizagem_Compartilha experiencia',
            'Aprendizagem_Conhecimentos de fontes internas e externas',
            'Aprendizagem_Busca se desenvolver',
            'Aprendizagem_Autonomia'
        ],
        'Execução': [
            'Execucao_Resolve as dificuldades',
            'Execucao_Imprevistos e alteracoes',
            'Execucao_Ideias em Acoes'
        ]
    }
    
    # Calcular médias por categoria para cada tipo
    dados_radar = {}
    for tipo in ['autoavaliacao', 'avaliacao_pelo_lider', 'avaliacao_pelo_liderado']:
        df_tipo = df_pessoa[df_pessoa['Tipo'] == tipo]
        if not df_tipo.empty:
            medias_categoria = []
            for categoria, questoes in categorias.items():
                # Filtrar questões que existem no dataframe
                questoes_existentes = [q for q in questoes if q in df_tipo.columns]
                if questoes_existentes:
                    media = df_tipo[questoes_existentes].mean().mean()
                    medias_categoria.append(media if not pd.isna(media) else 0)
                else:
                    medias_categoria.append(0)
            dados_radar[tipo] = medias_categoria
    
    if not dados_radar:
        return None
    
    # Configurar radar
    N = len(categorias)
    angulos = [n / float(N) * 2 * np.pi for n in range(N)]
    angulos += angulos[:1]
    
    # Criar figura com mais espaço para a legenda
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(projection='polar'))
    
    # Plotar cada tipo
    for tipo in dados_radar.keys():
        valores = dados_radar[tipo]
        valores += valores[:1]
        ax.plot(angulos, valores, 'o-', linewidth=2, 
                label=TIPOS_AMIGAVEIS[tipo], color=CORES_TIPOS[tipo])
        ax.fill(angulos, valores, alpha=0.1, color=CORES_TIPOS[tipo])
    
    # Configurar
    ax.set_xticks(angulos[:-1])
    
    # Configurar labels com melhor espaçamento
    xtick_labels = list(categorias.keys())
    ax.set_xticklabels(xtick_labels, fontsize=11)
    
    # Ajustar posição das labels para ficarem fora do gráfico
    for label, angle in zip(ax.get_xticklabels(), angulos[:-1]):
        texto = label.get_text()

        # Regra geral baseada no ângulo
        if 0 <= angle < np.pi:
            alinhamento = 'left'
        else:
            alinhamento = 'right'

        # Exceções específicas
        if texto == 'Relacionamento':
            alinhamento = 'right'

        if texto == 'Execução':
            alinhamento = 'left'

        label.set_horizontalalignment(alinhamento)
        label.set_rotation(angle * 180 / np.pi - 90)


    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels(['1', '2', '3', '4', '5'], fontsize=9)
    
    # Adicionar grade para melhor visualização
    ax.grid(True, alpha=0.3)
    
    # Posicionar legenda fora do gráfico
    ax.legend(loc='upper left', bbox_to_anchor=(-0.35, 1.25), 
              fontsize=10, framealpha=0.9)
    
    # Ajustar layout para dar mais espaço
    plt.tight_layout(rect=[0, 0, 1, 1]) 
    
    # Salvar gráfico
    caminho_grafico = os.path.join(DIRETORIO_GRAFICOS_TEMP, f"radar_geral_{id(df_pessoa)}.png")
    fig.savefig(caminho_grafico, dpi=150, bbox_inches='tight', pad_inches=0.5)
    plt.close(fig)
    
    return caminho_grafico

def criar_radar_lideranca(df_pessoa):
    """Cria gráfico de radar para competências de liderança"""
    # Calcular médias por questão para cada tipo
    dados_radar = {}
    for tipo in ['autoavaliacao', 'avaliacao_pelo_lider', 'avaliacao_pelo_liderado']:
        df_tipo = df_pessoa[df_pessoa['Tipo'] == tipo]
        if not df_tipo.empty:
            medias = []
            for questao in COLUNAS_LIDERANCA:
                if questao in df_tipo.columns:
                    media = df_tipo[questao].mean()
                    medias.append(media if not pd.isna(media) else 0)
                else:
                    medias.append(0)
            dados_radar[tipo] = medias
    
    if not dados_radar:
        return None
    
    # Configurar radar
    N = len(COLUNAS_LIDERANCA)
    angulos = [n / float(N) * 2 * np.pi for n in range(N)]
    angulos += angulos[:1]
    
    # Criar figura com mais espaço para a legenda
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(projection='polar'))
    
    # Encurtar nomes para o radar
    labels = [q.replace('Lideranca_', '').replace('_', ' ') for q in COLUNAS_LIDERANCA]
    
    # Plotar cada tipo
    for tipo in dados_radar.keys():
        valores = dados_radar[tipo]
        valores += valores[:1]
        ax.plot(angulos, valores, 'o-', linewidth=2, 
                label=TIPOS_AMIGAVEIS[tipo], color=CORES_TIPOS[tipo])
        ax.fill(angulos, valores, alpha=0.1, color=CORES_TIPOS[tipo])
    
    # Configurar
    ax.set_xticks(angulos[:-1])
    
    # Configurar labels com melhor espaçamento
    ax.set_xticklabels(labels, fontsize=10)
    
    # Ajustar posição das labels para ficarem fora do gráfico
    for label, angle in zip(ax.get_xticklabels(), angulos[:-1]):
        texto = label.get_text()
    
        # Regra geral baseada no ângulo
        if 0 <= angle < np.pi:
            alinhamento = 'left'
        else:
            alinhamento = 'right'

        # Exceções específicas
        if texto == 'Delegação':
            alinhamento = 'right'

        label.set_horizontalalignment(alinhamento)
        label.set_rotation(angle * 180 / np.pi - 90)


    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels(['1', '2', '3', '4', '5'], fontsize=9)
    
    # Adicionar grade para melhor visualização
    ax.grid(True, alpha=0.3)
    
    # Posicionar legenda fora do gráfico
    ax.legend(loc='upper left', bbox_to_anchor=(-0.35, 1.25), 
              fontsize=10, framealpha=0.9)
    
    # Ajustar layout para dar mais espaço
    plt.tight_layout()  
    
    # Salvar gráfico
    caminho_grafico = os.path.join(DIRETORIO_GRAFICOS_TEMP, f"radar_lideranca_{id(df_pessoa)}.png")
    fig.savefig(caminho_grafico, dpi=150, bbox_inches='tight', pad_inches=0.5)
    plt.close(fig)
    
    return caminho_grafico

def limpar_placeholders(slide):
    """Remove os placeholders {{...}} do slide"""
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            if "{{" in shape.text:
                shape.text = shape.text.replace("{{", "").replace("}}", "")

def adicionar_imagem_no_placeholder(slide, placeholder_nome, caminho_imagem, posicao=None, tamanho=None):
    """Adiciona uma imagem no lugar de um placeholder"""
    for shape in slide.shapes:
        if hasattr(shape, "text") and f"{{{{{placeholder_nome}}}}}" in shape.text:
            # Se encontramos um placeholder de texto
            left = shape.left
            top = shape.top
            width = shape.width
            height = shape.height
            
            # Remover a forma de texto
            sp = shape._element
            sp.getparent().remove(sp)
            
            # Adicionar imagem
            slide.shapes.add_picture(caminho_imagem, left, top, width, height)
            return True
    return False

def substituir_texto_no_slide(slide, placeholder, valor, max_caracteres=None):
    """Substitui texto em um placeholder"""
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            if f"{{{{{placeholder}}}}}" in shape.text:
                # Limitar texto se max_caracteres for especificado
                if max_caracteres and len(str(valor)) > max_caracteres:
                    valor = str(valor)[:max_caracteres-3] + "..."
                shape.text = shape.text.replace(f"{{{{{placeholder}}}}}", str(valor))
                return True
    return False

def substituir_texto_formatado(slide, placeholder, valor, tamanho_fonte=14, cor_hex="#CC8A42", centralizar=True, negrito=True):
    """Substitui texto em um placeholder com formatação específica"""
    for shape in slide.shapes:
        if hasattr(shape, "text") and f"{{{{{placeholder}}}}}" in shape.text:
            shape.text = shape.text.replace(f"{{{{{placeholder}}}}}", str(valor))
            
            # Formatação do texto
            for paragraph in shape.text_frame.paragraphs:
                if centralizar:
                    paragraph.alignment = PP_ALIGN.CENTER
                
                for run in paragraph.runs:
                    run.font.size = Pt(tamanho_fonte)
                    run.font.bold = negrito
                    
                    # Converter cor hexadecimal para RGB
                    if cor_hex:
                        cor_hex = cor_hex.lstrip('#')
                        r = int(cor_hex[0:2], 16)
                        g = int(cor_hex[2:4], 16)
                        b = int(cor_hex[4:6], 16)
                        run.font.color.rgb = RGBColor(r, g, b)
            
            return True
    return False


def gerar_relatorio_pessoa(nome_pessoa, df_total):
    """Gera relatório para uma pessoa específica"""
    print(f"Gerando relatório para: {nome_pessoa}")
    
    # Filtrar dados da pessoa
    df_pessoa = filtrar_dados_pessoa(df_total, nome_pessoa)
    
    if df_pessoa.empty:
        print(f"  ⚠️  Nenhum dado encontrado para {nome_pessoa}")
        return None
    
    # Carregar template
    prs = Presentation(CAMINHO_TEMPLATE)


    # =================== SLIDE 1: CAPA ===================
    slide0 = prs.slides[0]

   # Procurar pelo placeholder {{nome}} e formatá-lo
    for shape in slide0.shapes:
        if hasattr(shape, "text") and "{{nome}}" in shape.text:
            # Substituir o texto
            shape.text = shape.text.replace("{{nome}}", nome_pessoa)
            
            # Formatar a fonte - tamanho 18 e negrito
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(18)  # Tamanho 18
                    run.font.bold = True    # Negrito
            
            break  # Parar após encontrar o primeiro placeholder {{nome}}
    
    # =================== SLIDE 2: COMPETÊNCIAS GERAIS ===================
    slide1 = prs.slides[1]
    
    # Calcular médias gerais
    nota_auto = calcular_medias_tipo(df_pessoa, 'autoavaliacao', COLUNAS_GERAIS)
    nota_lider = calcular_medias_tipo(df_pessoa, 'avaliacao_pelo_lider', COLUNAS_GERAIS)
    nota_liderado = calcular_medias_tipo(df_pessoa, 'avaliacao_pelo_liderado', COLUNAS_GERAIS)
    
    cor_nota = "#CC8A42"

    # Substituir placeholders de notas
    substituir_texto_formatado(slide1, "nota_auto", 
                            formatar_nota(nota_auto.mean() if nota_auto is not None else None),
                            tamanho_fonte=14, cor_hex=cor_nota, centralizar=True)
    substituir_texto_formatado(slide1, "nota_lider", 
                            formatar_nota(nota_lider.mean() if nota_lider is not None else None),
                            tamanho_fonte=14, cor_hex=cor_nota, centralizar=True)
    substituir_texto_formatado(slide1, "nota_liderado", 
                            formatar_nota(nota_liderado.mean() if nota_liderado is not None else None),
                            tamanho_fonte=14, cor_hex=cor_nota, centralizar=True)
    
    # Gerar e adicionar radar geral
    radar_geral_path = criar_radar_geral(df_pessoa)
    if radar_geral_path:
        # Substituir placeholder de radar
        for shape in slide1.shapes:
            if hasattr(shape, "text") and "{{radar_geral}}" in shape.text:
                left = shape.left
                top = shape.top
                width = shape.width
                height = shape.height
                
                # Remover placeholder
                sp = shape._element
                sp.getparent().remove(sp)
                
                # Adicionar imagem do radar
                slide1.shapes.add_picture(radar_geral_path, left, top, width, height)
                break
    
    # =================== SLIDE 3: COMPETÊNCIAS DE LIDERANÇA ===================
    slide2 = prs.slides[2]
    
    # Calcular médias de liderança
    nota_lider_auto = calcular_medias_tipo(df_pessoa, 'autoavaliacao', COLUNAS_LIDERANCA)
    nota_lider_lider = calcular_medias_tipo(df_pessoa, 'avaliacao_pelo_lider', COLUNAS_LIDERANCA)
    nota_lider_liderado = calcular_medias_tipo(df_pessoa, 'avaliacao_pelo_liderado', COLUNAS_LIDERANCA)
    
    # Substituir placeholders
    substituir_texto_formatado(slide2, "nota_lider_auto", 
                            formatar_nota(nota_lider_auto.mean() if nota_lider_auto is not None else None),
                            tamanho_fonte=14, cor_hex=cor_nota, centralizar=True)
    substituir_texto_formatado(slide2, "nota_lider_lider", 
                            formatar_nota(nota_lider_lider.mean() if nota_lider_lider is not None else None),
                            tamanho_fonte=14, cor_hex=cor_nota, centralizar=True)
    substituir_texto_formatado(slide2, "nota_lider_liderado", 
                            formatar_nota(nota_lider_liderado.mean() if nota_lider_liderado is not None else None),
                            tamanho_fonte=14, cor_hex=cor_nota, centralizar=True)
    
    # Gerar e adicionar radar de liderança
    radar_lideranca_path = criar_radar_lideranca(df_pessoa)
    if radar_lideranca_path:
        # Substituir placeholder de radar (segundo slide)
        for shape in slide2.shapes:
            if hasattr(shape, "text") and "{{radar_geral}}" in shape.text:
                left = shape.left
                top = shape.top
                width = shape.width
                height = shape.height
                
                # Remover placeholder
                sp = shape._element
                sp.getparent().remove(sp)
                
                # Adicionar imagem do radar
                slide2.shapes.add_picture(radar_lideranca_path, left, top, width, height)
                break
    
    # =================== SLIDES 5-...: DETALHAMENTO DAS COMPETÊNCIAS ===================
    # Para cada questão, gerar gráfico e substituir no slide correspondente
    for questao, placeholder in MAPEAMENTO_QUESTOES_GRAFICOS.items():
        # Encontrar qual slide tem este placeholder
        for slide_idx, slide in enumerate(prs.slides):
            for shape in slide.shapes:
                if hasattr(shape, "text") and f"{{{{{placeholder}}}}}" in shape.text:
                    # Gerar gráfico
                    grafico_path = criar_grafico_barras_questao(df_pessoa, questao, placeholder)
                    
                    if grafico_path:
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height
                        
                        # Remover placeholder
                        sp = shape._element
                        sp.getparent().remove(sp)
                        
                        # Adicionar imagem do gráfico
                        slide.shapes.add_picture(grafico_path, left, top, width, height)
                    
                    break
    
    # =================== SLIDES FINAIS: QUESTÕES ABERTAS ===================
    # Extrair respostas abertas
    respostas_por_tipo = {}
    for tipo in ['autoavaliacao', 'avaliacao_pelo_lider', 'avaliacao_pelo_liderado']:
        df_tipo = df_pessoa[df_pessoa['Tipo'] == tipo]
        if not df_tipo.empty:
            pontos_fortes = df_tipo['Pontos Fortes'].dropna().unique()
            oportunidades = df_tipo['Oportunidades de Desenvolvimento'].dropna().unique()
            
            # Juntar todas as respostas
            pf_texto = '; '.join([str(pf).strip() for pf in pontos_fortes if str(pf).strip() != ''])
            od_texto = '; '.join([str(od).strip() for od in oportunidades if str(od).strip() != ''])
            
            # Limitar se for muito longo (para visualização inicial)
            if len(pf_texto) > 800:
                pf_texto = pf_texto[:797] + "..."
            if len(od_texto) > 800:
                od_texto = od_texto[:797] + "..."
            
            respostas_por_tipo[tipo] = {
                'PF': pf_texto if pf_texto else "Não informado",
                'OD': od_texto if od_texto else "Não informado"
            }
        else:
            respostas_por_tipo[tipo] = {
                'PF': "Não se aplica",
                'OD': "Não se aplica"
            }
    
    # Mapeamento dos placeholders para cada slide (agora com índice correto)
    # VERIFIQUE OS ÍNDICES CORRETOS COM O CÓDIGO DE DEBUG ANTERIOR
    mapeamento_placeholders = [
        (24, "PF_auto", respostas_por_tipo['autoavaliacao']['PF']),
        (25, "PF_lider", respostas_por_tipo['avaliacao_pelo_lider']['PF']),
        (26, "PF_liderado", respostas_por_tipo['avaliacao_pelo_liderado']['PF']),
        (27, "OD_auto", respostas_por_tipo['autoavaliacao']['OD']),
        (28, "OD_lider", respostas_por_tipo['avaliacao_pelo_lider']['OD']),
        (29, "OD_liderado", respostas_por_tipo['avaliacao_pelo_liderado']['OD'])
    ]
    
    # Primeiro, substituir todos os placeholders
    for slide_idx, placeholder, valor in mapeamento_placeholders:
        if slide_idx < len(prs.slides):
            substituir_texto_no_slide(prs.slides[slide_idx], placeholder, valor, max_caracteres=800)
    
    # Verificar textos muito longos e dividir em slides adicionais
    slides_para_processar = []
    for slide_idx, placeholder, valor in mapeamento_placeholders:
        if slide_idx < len(prs.slides) and len(str(valor)) > 800:
            slides_para_processar.append((slide_idx, placeholder, valor))
    
    # Processar divisão de slides (em ordem reversa para não afetar índices)
    slides_para_processar.sort(reverse=True)  # Do último para o primeiro
    for slide_idx, placeholder, valor in slides_para_processar:
        if len(str(valor)) > 800:
            dividir_texto_em_slides(prs, str(valor), slide_idx, placeholder, max_caracteres=800)
    
    # Limpar todos os placeholders restantes
    for slide in prs.slides:
        limpar_placeholders(slide)
    
    # Salvar apresentação
    nome_arquivo = f"Relatorio_{nome_pessoa.replace(' ', '_').replace('/', '_')}.pptx"
    caminho_saida = os.path.join(DIRETORIO_SAIDA, nome_arquivo)
    prs.save(caminho_saida)
    
    print(f"  ✅ Relatório salvo em: {caminho_saida}")
    
    # Limpar gráficos temporários
    limpar_arquivos_temporarios(df_pessoa)
    
    return caminho_saida

def limpar_arquivos_temporarios(df_pessoa):
    """Limpa arquivos temporários gerados para uma pessoa"""
    import glob
    import os
    
    # Padrão de arquivos temporários
    padrao = os.path.join(DIRETORIO_GRAFICOS_TEMP, f"*_{id(df_pessoa)}.png")
    
    for arquivo in glob.glob(padrao):
        try:
            os.remove(arquivo)
        except:
            pass

def dividir_texto_em_slides(prs, texto, slide_template_idx, placeholder, max_caracteres=800):
    """
    Divide texto longo em múltiplos slides.
    Retorna a lista de slides criados (incluindo o original modificado)
    """
    if len(texto) <= max_caracteres:
        return [slide_template_idx]
    
    slides_criados = [slide_template_idx]
    partes = []
    
    # Dividir o texto em parágrafos primeiro
    paragrafos = texto.split('; ')
    
    parte_atual = ""
    for paragrafo in paragrafos:
        if len(parte_atual) + len(paragrafo) + 2 <= max_caracteres:
            if parte_atual:
                parte_atual += "; " + paragrafo
            else:
                parte_atual = paragrafo
        else:
            if parte_atual:
                partes.append(parte_atual)
            parte_atual = paragrafo
    
    if parte_atual:
        partes.append(parte_atual)
    
    # Se só tem uma parte, retorna o slide original
    if len(partes) <= 1:
        return [slide_template_idx]
    
    # Pegar o slide original como template
    slide_original = prs.slides[slide_template_idx]
    layout_original = slide_original.slide_layout
    
    # Modificar o slide original com a primeira parte
    substituir_texto_no_slide(slide_original, placeholder, partes[0] + " (continua...)")
    
    # Criar slides adicionais para as partes restantes
    for i, parte in enumerate(partes[1:], 1):
        novo_slide = prs.slides.add_slide(layout_original)
        slides_criados.append(len(prs.slides) - 1)
        
        # Copiar o título do slide original
        for shape_orig in slide_original.shapes:
            if shape_orig.has_text_frame and shape_orig.text_frame.text:
                # Verificar se é o título (geralmente é a primeira forma com texto)
                for shape_novo in novo_slide.shapes:
                    if shape_novo.has_text_frame:
                        # Manter o título original
                        if "(continuação" not in shape_novo.text_frame.text:
                            # Adicionar indicação de continuação
                            for paragraph in shape_novo.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    if f"{{{{{placeholder}}}}}" not in run.text:
                                        run.text = run.text + f" (continuação {i+1})"
        
        # Substituir o placeholder no novo slide
        if i == len(partes) - 2:  # Penúltima parte
            substituir_texto_no_slide(novo_slide, placeholder, parte)
        else:
            substituir_texto_no_slide(novo_slide, placeholder, parte + " (continua...)")
    
    return slides_criados

# =================== FUNÇÃO PRINCIPAL ===================
def gerar_relatorios_todos(df_total):
    """Gera relatórios para todas as pessoas no dataframe"""
    print("=" * 60)
    print("INICIANDO GERAÇÃO DE RELATÓRIOS")
    print("=" * 60)
    
    # Verificar se há dados
    if df_total.empty:
        print("Dataframe vazio!")
        return
    
    # Listar pessoas únicas
    pessoas = df_total['Nome'].unique()
    print(f"Total de pessoas encontradas: {len(pessoas)}")
    
    # Gerar relatório para cada pessoa
    relatorios_gerados = []
    
    for i, nome_pessoa in enumerate(pessoas, 1):
        print(f"\n[{i}/{len(pessoas)}] ", end="")
        caminho_relatorio = gerar_relatorio_pessoa(nome_pessoa, df_total)
        if caminho_relatorio:
            relatorios_gerados.append(caminho_relatorio)
    
    # Resumo
    print("\n" + "=" * 60)
    print("RESUMO DA GERAÇÃO")
    print("=" * 60)
    print(f"Total de pessoas processadas: {len(pessoas)}")
    print(f"Relatórios gerados com sucesso: {len(relatorios_gerados)}")
    print(f"Diretório de saída: {DIRETORIO_SAIDA}")
    
    return relatorios_gerados

# =================== EXECUÇÃO ===================
if __name__ == "__main__":
    
    # Verificar estrutura do dataframe
    print("Colunas disponíveis no dataframe:")
    print(df_total.columns.tolist())
    print(f"\nTotal de registros: {len(df_total)}")
    print(f"Tipos de avaliação: {df_total['Tipo'].unique()}")
    
    # Gerar todos os relatórios
    relatorios = gerar_relatorios_todos(df_total)
    
    # Limpar diretório de gráficos temporários completamente
    try:
        import shutil
        shutil.rmtree(DIRETORIO_GRAFICOS_TEMP)
        os.makedirs(DIRETORIO_GRAFICOS_TEMP, exist_ok=True)
    except:
        pass




