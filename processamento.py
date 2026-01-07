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


# PROCESSAMENTO DOS DADOS

# Carregar dados da autoavaliação e avaliação de líderes e liderados
df_autoavaliacao = pd.read_excel('dados/Autoavaliação da equipe GA RN 2025 (respostas).xlsx')

df_avaliacao = pd.read_excel('dados/Avaliação da equipe GA RN (Líder __ Liderado) (respostas).xlsx')

# Limpeza das colunas desnecessárias
df_autoavaliacao = df_autoavaliacao = df_autoavaliacao.drop(columns=
                                 ['Carimbo de data/hora', 
                                  'Endereço de e-mail', 
                                  'Qual seu e-mail com domínio @FGV:', 
                                  'Qual seu e-mail com domínio @SEED (@educar):',
                                  'Qual sua Frente de atuação?',
                                  'Qual seu cargo/função?',
                                  ' Data de entrada no projeto',
                                  'Quais foram suas principais conquistas desde a sua entrada no projeto?',
                                  'Quais foram seus principais desafios durante o projeto?',
                                  'Como a FGV DGPE, neste projeto, pode apoiar nesse desenvolvimento?'
                                  ])

df_avaliacao = df_avaliacao.drop(columns=
                                 ['Carimbo de data/hora',
                                  'Endereço de e-mail', 
                                  'Avaliador(a), selecione o seu nome completo:',
                                  'Sua Frente de atuação:',
                                  'Caso necessário, deixe aqui sugestões para o desenvolvimento do colaborador avaliado:'
                                  ])


# Separação do df_avaliação em 2 dataframes: líderes e liderados
df_lider = df_avaliacao[df_avaliacao["Relação com o avaliado(a): "] == "Avaliado(a) é meu liderado"]
df_lider = df_lider.reset_index(drop=True)

df_liderado = df_avaliacao[df_avaliacao["Relação com o avaliado(a): "] == "Avaliado(a) é meu líder"]
df_liderado = df_liderado.reset_index(drop=True)


# Criar variável para identificar o tipo de avaliação
df_autoavaliacao['tipo'] = 'autoavaliacao'
df_lider['tipo'] = 'avaliacao_pelo_lider'
df_liderado['tipo'] = 'avaliacao_pelo_liderado'

# Preparar os dataframes para a concatenação
# Retirar colunas desnecessárias
df_lider = df_lider.drop(columns=['Relação com o avaliado(a): '])
df_liderado = df_liderado.drop(columns=['Relação com o avaliado(a): '])

# Trocar os nomes das colunas para facilitar a concatenação
nomes_colunas = ['Nome',
                            'Lideranca_Desenvolvimento de Pessoas', 
                            'Lideranca_Visão Estratégica',
                            'Lideranca_Delegação',
                            'Lideranca_Gerenciamento de Riscos',
                            'Lideranca_Monitoramento de Resultados',
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
                            'Execucao_Ideias em Acoes',
                            'Pontos Fortes',
                            'Oportunidades de Desenvolvimento',
                            'Tipo'
                            ]

df_autoavaliacao.columns = nomes_colunas
df_lider.columns = nomes_colunas
df_liderado.columns = nomes_colunas


# Juntar os três em um único dataframe, mostrando a diferença entre os 3 tipos: autoavaliação, avlaiação pelo liderado e avaliação pelo líder
df_total = pd.concat([df_autoavaliacao, df_lider, df_liderado], ignore_index=True)

# Trocar os valores textuais por valores númericos (1 a 5) das questões de liderança
map_desenv_pessoas = {
    "Não demonstra interesse em desenvolver os outros. Centraliza decisões e responsabilidades. Não oferece feedbacks ou oportunidades de assumir novos desafios.": 1,
    "Reconhece a importância de desenvolver a equipe, mas faz isso de forma esporádica. Falta constância no incentivo à autonomia e no apoio ao aprendizado.": 2,
    "Identifica e fortalece talentos. Estimula o crescimento e desenvolvimento profissional da equipe, fornecendo suporte, direcionamento e oportunidades de aprendizado contínuo, incentivando autonomia e preparando novas lideranças.": 3,
    "Desenvolve ativamente as pessoas da equipe com planos claros e feedbacks frequentes. Estimula alto nível de autonomia e crescimento contínuo. Cria oportunidades reais de crescimento, de maior visibilidade e está sempre pensando em como preparar as pessoas para desafios maiores.": 4,
    "Forma novas lideranças, inspira outras pessoas a também cuidarem do desenvolvimento de suas equipes e deixa um legado claro por onde passa. Vai além do seu time, fomentando o desenvolvimento por meio de trocas, mentorias, rodas de conversa, palestras e outras iniciativas que fortalecem a cultura de aprendizagem da organização.": 5
}

map_visao_estrategica = {
    "Realiza suas atividades de forma automática e com uma preocupação apenas no operacional. Não relaciona suas ações aos objetivos maiores da sua área ou projeto. Tem dificuldade em compreender ou comunicar a estratégia para a equipe gerando inseguranças, retrabalhos e desmotivação": 1,
    "Demonstra alguma percepção sobre objetivos finais da área ou projeto, mas tem dificuldade em conectar o trabalho do dia a dia com essa estratégia maior. Costuma agir de forma imediatista e reativa tomando decisões desalinhadas ou  perdendo oportunidades por falta de análise mais ampla dos cenários e contextos.": 2,
    "Analisa cenários para antecipar tendências e oportunidades, tomando decisões fundamentadas e alinhadas aos objetivos da área ou projeto. Conecta atividades diárias à estratégia da empresa, promovendo inovação e sustentabilidade.": 3,
    "Toma decisões alinhadas à estratégia e incentiva o time a pensar com visão sistêmica. Considera impactos sustentáveis (sociais, ambientais, financeiros) nos projetos ou área. Atua como ponte entre níveis operacionais e estratégicos.": 4,
    "Possui uma visão ampla dos cenários internos e externos, antecipando questões que possam impactar o futuro da organização. Engaja colegas e lideranças com clareza de propósito e traduz a estratégia organizacional em ações concretas, promovendo uma cultura de adaptação contínua e sustentável.": 5
}

map_delegacao = {
    "Demonstra baixa confiança nas pessoas, evitando delegar e centralizando decisões. Resiste à autonomia da equipe, recorrendo ao microgerenciamento. Essa postura limita a colaboração, gera insegurança e impacta negativamente a moral do time.": 1,
    "Tem dificuldade para distribuir responsabilidades, mantendo um controle rigoroso sobre o trabalho da equipe. Fornece pouca liberdade para que as pessoas encontrem suas próprias soluções, o que diminui o engajamento e gera frustração por falta de oportunidades para crescer e desenvolver habilidades.": 2,
    "Confia na equipe e distribui responsabilidades estrategicamente, alinhando expectativas e acompanhando entregas. Oferece apoio, recursos e desafios que estimulem o crescimento profissional, promovendo autonomia e desenvolvimento contínuo.": 3,
    "Cria um ambiente de confiança, fornecendo condições e informações para que a equipe atue com autonomia e eficiência. Incentiva a iniciativa e o desenvolvimento profissional, promovendo entregas de qualidade e valorizando os pontos fortes individuais de maneira inteligente..": 4,
    "Confia na equipe e delega com equilíbrio, garantindo resultados. Forma líderes autônomos e responsáveis, difundindo essa cultura na área. Promove ativamente a gestão do conhecimento e a troca de boas práticas, ampliando o impacto e gerando ganhos em escala.": 5
}

map_riscos = {
    "Ignora riscos e mudanças de cenário, apego às atividades rotineiras mesmo quando ineficazes. Reage mal à escassez de recursos, não aprende com erros e toma decisões precipitadas ou paralisadas por excesso de cautela. Dificulta colaborações e retroalimenta um ambiente rígido, com baixo diálogo e pouca abertura à melhoria.": 1,
    "Reconhece riscos, mas age tardiamente. Tem dificuldade em mobilizar pessoas, priorizar ações e transformar experiências em aprendizados. Atua de forma reativa, sem visão sistêmica ou planejamento preventivo consistente.": 2,
    "Identifica e monitora riscos que possam impactar a equipe, os resultados e a organização, propondo soluções preventivas e ajustando estratégias, conforme necessário. Trabalha colaborativamente na criação de planos de contingência para garantir a estabilidade e segurança dos projetos ou das atividades das áreas.": 3,
    "Integra a gestão de riscos ao planejamento desde o início. Prevê obstáculos com base em múltiplas fontes e atua de forma colaborativa, envolvendo diferentes áreas. Estimula a equipe a adotar uma postura preventiva, mantendo um ambiente de confiança, agilidade e segurança.": 4,
    "Age de forma preventiva e estratégica, antecipando riscos complexos que poderiam impactar a equipe, os resultados ou a organização como um todo. Constrói uma cultura sólida de avaliação e mitigação de riscos, formando outras pessoas com esse olhar e incentivando que práticas de prevenção sejam incorporadas por todas as camadas da equipe.": 5
}

map_resultados = {
    "Demonstra pouca ou nenhuma clareza na definição de prazos, metas e padrões de qualidade. Não realiza acompanhamento sistemático das atividades, o que leva a desvios recorrentes nas entregas. Há resistência ou falta de iniciativa para revisar processos, mesmo diante de falhas evidentes. A cultura de melhoria contínua não é percebida na prática, gerando retrabalho e desalinhamento com os padrões DGPE.": 1,
    "Ainda apresenta dificuldades em comunicar expectativas de forma clara e consistente. O monitoramento das atividades ocorre de forma esporádica ou superficial, limitando a capacidade de corrigir desvios a tempo. Ajustes de processo são pontuais e geralmente reativos, sem abordagem estruturada de melhoria. A prática de melhoria contínua não é incorporada como parte da rotina.": 2,
    "Estabelece expectativas claras sobre prazos, qualidade e resultados, acompanhando as atividades para garantir alinhamento ao padrão DGPE. Ajusta processos conforme necessário para otimizar entregas, promovendo uma cultura de melhoria contínua.": 3,
    "Comunica com clareza objetivos, prazos e critérios de qualidade. Monitora atividades de forma sistemática, intervindo quando necessário. Revê e aprimora processos proativamente e estimula a melhoria contínua, compartilhando boas práticas alinhadas aos padrões DGPE.": 4,
    "Mantém alto padrão no acompanhamento de entregas, assegurando prazos, qualidade e alinhamento estratégico. Desenvolve lideranças autônomas e fortalece a cultura de responsabilidade e melhoria contínua. Compartilha boas práticas e contribui para a gestão orientada a resultados na instituição.": 5
}


# Mapeamento por coluna
df_total["Lideranca_Desenvolvimento de Pessoas"] = (
    df_total["Lideranca_Desenvolvimento de Pessoas"].map(map_desenv_pessoas))

df_total["Lideranca_Visão Estratégica"] = (
    df_total["Lideranca_Visão Estratégica"].map(map_visao_estrategica))

df_total["Lideranca_Delegação"] = (
    df_total["Lideranca_Delegação"].map(map_delegacao))

df_total["Lideranca_Gerenciamento de Riscos"] = (
    df_total["Lideranca_Gerenciamento de Riscos"].map(map_riscos))

df_total["Lideranca_Monitoramento de Resultados"] = (
    df_total["Lideranca_Monitoramento de Resultados"].map(map_resultados))


# Salvar o dataframe processado em um novo arquivo csv para análises e relatórios
df_total.to_csv('dados/processado_dados_avaliacao.csv', index=False)








