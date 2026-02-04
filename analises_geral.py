from glob import glob
import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import warnings
warnings.filterwarnings('ignore')
import glob
import matplotlib.pyplot as plt


# Carregando o dataframe processado
df_total = pd.read_csv('dados/processado_dados_avaliacao.csv')

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


# Dataframe temporário
df_temp = df_total.copy()

# Média das competências gerais por linha
df_temp['Media_Geral'] = df_temp[COLUNAS_GERAIS].mean(axis=1)

# Agrupar por Nome e Tipo e calcular média
df_agrupado = (
    df_temp
    .groupby(['Nome', 'Tipo'])['Media_Geral']
    .mean()
    .reset_index()
)

# Pivotar Tipo para colunas
df_resultado = (
    df_agrupado
    .pivot(index='Nome', columns='Tipo', values='Media_Geral')
    .reset_index()
)

# Criar novas colunas para o dataframe
# Média entre Auto + Líder + Liderado
df_resultado['Media Geral (Auto + Lider + Liderado)'] = df_resultado[
    ['autoavaliacao', 'avaliacao_pelo_lider', 'avaliacao_pelo_liderado']
].mean(axis=1)

# Média entre Líder + Liderado
df_resultado['Media Terceiros (lider + liderado)'] = df_resultado[
    ['avaliacao_pelo_lider', 'avaliacao_pelo_liderado']
].mean(axis=1)

# Média entre Auto + Média de terceiros
df_resultado['Media_Auto e Terceiros'] = df_resultado[
    ['autoavaliacao', 'Media Terceiros (lider + liderado)']
].mean(axis=1)

# Arredondar para 2 casas decimais
df_resultado = df_resultado.round(2)

# Salvar o dataframe em Excel
df_resultado.to_excel('dados/analise_geral_competencias.xlsx', index=False)


################ MATRIZ DE CORRELAÇÃO ENTRE ALGUMAS MÉTRICAS
def plot_matriz_9box(df, eixo_x, eixo_y, titulo, coluna_cor='Nome', salvar_pdf=None):

    fig, ax = plt.subplots(figsize=(10, 7))

    # Limites
    ax.set_xlim(0, 5)
    ax.set_ylim(0, 5)

    # Cortes
    cut_x1, cut_x2 = 1.67, 3.33
    cut_y1, cut_y2 = 1.67, 3.33

    # ======================
    # FUNDO DOS QUADRANTES
    # ======================

    # Inferior
    ax.axvspan(0, cut_x1, ymin=0, ymax=cut_y1/5, alpha=0.4, color='#c58b9b')
    ax.axvspan(cut_x1, cut_x2, ymin=0, ymax=cut_y1/5, alpha=0.4, color='#e5d7a5')
    ax.axvspan(cut_x2, 5, ymin=0, ymax=cut_y1/5, alpha=0.4, color='#8fb7cc')

    # Meio
    ax.axvspan(0, cut_x1, ymin=cut_y1/5, ymax=cut_y2/5, alpha=0.4, color='#e5d7a5')
    ax.axvspan(cut_x1, cut_x2, ymin=cut_y1/5, ymax=cut_y2/5, alpha=0.4, color='#8fb7cc')
    ax.axvspan(cut_x2, 5, ymin=cut_y1/5, ymax=cut_y2/5, alpha=0.4, color='#9ccdbf')

    # Superior
    ax.axvspan(0, cut_x1, ymin=cut_y2/5, ymax=1, alpha=0.4, color='#8fb7cc')
    ax.axvspan(cut_x1, cut_x2, ymin=cut_y2/5, ymax=1, alpha=0.4, color='#9ccdbf')
    ax.axvspan(cut_x2, 5, ymin=cut_y2/5, ymax=1, alpha=0.4, color='#8ecbb8')

    # ======================
    # LINHAS DE CORTE
    # ======================
    ax.axvline(cut_x1, linestyle='--', color='black')
    ax.axvline(cut_x2, linestyle='--', color='black')
    ax.axhline(cut_y1, linestyle='--', color='black')
    ax.axhline(cut_y2, linestyle='--', color='black')

    # ======================
    # DEFINIR CORES
    # ======================

    # Converte categoria para números
    categorias = df[coluna_cor].astype('category')
    cores = categorias.cat.codes

    scatter = ax.scatter(
        df[eixo_x],
        df[eixo_y],
        c=cores,
        cmap='tab20',   # paleta com várias cores distintas
        s=70,
        edgecolors='black'
    )

    # ======================
    # TEXTOS DOS QUADRANTES
    # ======================
    labels = [
        ("Insuficiente", 0.8, 0.4),
        ("Eficaz", 2.5, 0.4),
        ("Comprometido", 4.2, 0.4),

        ("Questionável", 0.8, 2.4),
        ("Mantenedor", 2.5, 2.4),
        ("Forte desempenho", 4.2, 2.4),

        ("Dúvida", 0.8, 4.4),
        ("Forte desempenho", 2.5, 4.4),
        ("Alto potencial", 4.2, 4.4),
    ]

    for text, x, y in labels:
        ax.text(x, y, text, ha='center', va='center', fontsize=10, weight='bold')

    # ======================
    # TÍTULO
    # ======================
    ax.set_title(titulo, fontsize=14, weight='bold')
    ax.set_xlabel(eixo_x)
    ax.set_ylabel(eixo_y)

    # Salvar PDF se caminho for informado
    if salvar_pdf:
        plt.savefig(salvar_pdf, bbox_inches='tight')

    plt.show()

# 1º Gráfico: Média de Autoavaliação e Terceiros x Média de Terceiros
plot_matriz_9box(
    df_resultado,
    'Media_Auto e Terceiros',
    'Media Terceiros (lider + liderado)',
    'Matriz: Média entre Autoavaliação + Terceiros vs Avaliação de Terceiros',
    salvar_pdf='matriz1_Media entre Auto e Terceiros.pdf'
)

# 2º Gráfico: Média Geral (Auto + Líder + Liderado) x Média de Terceiros
plot_matriz_9box(
    df_resultado,
    'Media Geral (Auto + Lider + Liderado)',
    'Media Terceiros (lider + liderado)',
    'Matriz: Média Geral (Auto + Líder + Liderado) vs Avaliação de Terceiros',
    salvar_pdf='matriz2_Media Geral.pdf'
)

# 3º Gráfico: Autoavaliação x Média de Terceiros
plot_matriz_9box(
    df_resultado,
    'autoavaliacao',
    'Media Terceiros (lider + liderado)',
    'Matriz: Autoavaliação vs Avaliação de Terceiros',
    salvar_pdf='matriz3_Autoavaliacao.pdf'
)








