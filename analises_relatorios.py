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















