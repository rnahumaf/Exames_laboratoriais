from PyPDF2 import PdfReader
import re
import easygui
import win32clipboard
import win32con

# Função para copiar texto RTF para a Área de Transferência
def copy_to_clipboard(html_text, rtf_text, plain_text):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    # HTML
    html_format = win32clipboard.RegisterClipboardFormat("HTML Format")
    win32clipboard.SetClipboardData(html_format, html_text)
    # RTF
    rtf_format = win32clipboard.RegisterClipboardFormat("Rich Text Format")
    win32clipboard.SetClipboardData(rtf_format, rtf_text)
    # Texto simples
    win32clipboard.SetClipboardData(win32con.CF_TEXT, plain_text)
    win32clipboard.CloseClipboard()

# Função para formatar valores com base no formato desejado
def format_value(exam, value, format_type):
    if exam in reference_values:
        low, high = reference_values[exam]
        if not low <= value <= high:
            if format_type == 'rtf':
                return f"{{\\b {value}}}"  # Formata em negrito no formato RTF
            elif format_type == 'html':
                return f"<strong>{value}</strong>"  # Formata em negrito no formato HTML
    return value

def format_html_for_clipboard(html_body):
    # Cabeçalho necessário para o formato HTML na área de transferência do Windows
    html_header = (
        "Version:0.9\r\n"
        "StartHTML:00000097\r\n"
        "EndHTML:{end_html:09d}\r\n"
        "StartFragment:00000133\r\n"
        "EndFragment:{end_fragment:09d}\r\n"
        "<html><body>\r\n"
        "<!--StartFragment-->"
    )
    html_footer = "<!--EndFragment-->\r\n</body></html>"

    # Montando o HTML completo
    html_complete = html_header + html_body + html_footer

    # Calculando onde o HTML e o fragmento terminam
    end_html = len(html_complete)
    end_fragment = html_complete.find(html_footer)

    # Incluindo os comprimentos no cabeçalho
    html_complete = html_complete.format(end_html=end_html, end_fragment=end_fragment)
    return html_complete.encode('utf-8')

# Função para remover zeros isolados após a vírgula em números
def remove_isolated_zeros(text):
    # Esta regex identifica números com um zero após a vírgula seguido por um espaço ou fim de string
    return re.sub(r'\.0+\s', ' ', text)

# Função para extrair valor de exame de um texto.
def extract_value(text, exam_name, pattern):
    # Procurando pelo padrão.
    match = re.search(pattern, text, re.I)
    if match:
        return match.group(1).strip()  # Retorna o valor do exame.
    else:
        return 'Não encontrado'  # Retorna esta string caso o exame não seja encontrado.

# Função para ajustar separadores numéricos 
def convert_to_float(number_str):
    # Remove o separador de milhares (ponto)
    without_thousands = number_str.replace('.', '')
    # Converte a vírgula em ponto decimal
    with_decimal_point = without_thousands.replace(',', '.')
    return float(with_decimal_point)

# Lista de exames e seus padrões regex.
exams = {
    'b-hCG': r'HCG.+?Resultado\s*:?\s*(\S+)',
    'ABO': r'GRUPO SANG.+?Grupo ABO\s*:?\s*(\S+)',
    'Rh': r'GRUPO SANG.+?Fator RH\s*:?\s*(\S+)',

    'Hb': r'HEMOGLOBINA\.{2,}:\s+(\d+(?:,\d*)?)\s+g/dl',
    'Ht': r'HEMATÓCRITO\.{2,}:\s+(\d+(?:,\d*)?)\s+%',
    'VCM': r'V\.?C\.?M\.{2,}:\s+(\d+(?:,\d*)?)\s+fL',
    'HCM': r'H\.?C\.?M\.{2,}:\s+(\d+(?:,\d*)?)\s+pg',
    'RDW': r'R\.?D\.?W\.{2,}:\s+(\d+(?:,\d*)?)\s+%',
    'Leuco': r'(?<!sedimento )Leucócitos.*?:.*?(\d+\.?\d*)',
    'Bast': r'Bastonetes.*?:.*?(?<=%)(\d+\.?\d*)',
    'Segm': r'entados.*?:.*?(?<=%)(\d+\.?\d*)',
    'Eosi': r'Eosinófilos.*?:.*?(?<=%)(\d+\.?\d*)',
    'Linf': r'Linfócitos.*?:.*?(?<=%)(\d+\.?\d*)',
    'Plaq': r'Plaquetas.*?:.*?(\d+.\d*)',

    'Retic': r'RETICULOCITOS.+?Resultado:?\s+(.*?)\s+%',
    'Ferro': r'FERRO SERICO.+?Resultado:?\s+(.*?)\s+[µu]g/dL',
    'Ferritina': r'FERRITINA.+?Resultado:?\s+(.*?)\s+ng/mL',
    'Transferrina': r'TRANSFERRINA.+?Resultado:?\s+(.*?)\s+mg/dL',
    'TSAT': r'(\d+,\d+,\d+,\d+)',
    'TIBC': r'CAPACIDADE.+?Resultado.+?(\d+(:?,\d*)?)',

    'PSAT': r'(?:PSA LIVRE/TOTAL|PSA TOTAL ULTRA).+?PSA TOTAL:?\s+(.*?)\s+ng/mL',
    'PSAL': r'PSA LIVRE/TOTAL.+?PSA LIVRE:?\s+(.*?)\s+ng/mL',
    'R': r'PSA LIVRE/TOTAL.+?RELAÇÃO:?.+?(.*?)\s*%',
    'PSOF1': r'SANGUE OCULTO DIRETO.+?Resultado:?\s*(\S+)',
    'PSOF2': r'FEZES 2.+?Resultado:?\s*(\S+)',
    'PSOF3': r'FEZES 3.+?Resultado:?\s*(\S+)',

    'Cr': r'CREATININA.+?Material: Soro.+?Resultado:?\s+(.*?)\s+mg/dL',
    'Clearance Cr. corrigido': r'CLEARANCE.+?Clearance corrigido:?\s*(\S+)',
    'Albuminúria 24h': r'MICROALBUMINURIA.+?Resultado:?\s+(.*?)\s+mg/24',
    'Proteinúria 24h': r'PROTE[IÍ]NA URIN.+?(\d+(:?.\d*)?)\s+mg/24',
    'Ur': r'UR[EÉ]IA.+?:.+?(\d+(:?.\d*)?)(?=\s+mg/dL)',
    'Cálcio total': r'(?<! de )C[AÁ]LCIO(?!\s+urin[áa]rio)(?!\s+I[OÔ]NICO).*?Resultado.*?(\d+(?:,\d*)?)',
    'Cálcio iônico': r'(?<! de )C[AÁ]LCIO I[OÔ]NICO.*?Resultado.*?(\d+(?:,\d*)?)',
    'Magnésio': r'Magn[ée]sio.*?Resultado.*?(\d+(?:,\d*)?)',
    'Fósforo': r'F[óo]sforo.*?Resultado:?\s*(\d+(?:,\d*)?)',
    'Na': r'SODIO.*?(\d+(?:,\d*)?)\s+(?:mg/dL|MEQ/L)',
    'K': r'POTASSIO.+?Resultado.*?(\d+(?:,\d*)?)',

    'GJ': r'GLICOSE JEJUM.+?(\d+(?:,\d*)?)(?=\s+mg/dL)',
    "HbA1c": r"(?<=A1C)[^%]*?(\d+(?:,\d*)?)(?=\s*%)",
    "A1a": r"(?<=A1a:)\s*(.*?)%",
    "A1b": r"(?<=A1b:)\s*(.*?)%",
    "HbF": r"(?<=Hb F:)\s*(.*?)%",
    "A1c lábil": r"(?<=A1c lábil:)\s*(.*?)%",
    "HbA": r"(?<=Hb A:)\s*(.*?)%",
 
    'TSH': r'TSH.+?(\d+(?:,\d*)?)(?=\s+[uµ]UI/mL)',
    'T4L': r'T4 LIVRE.+?Resultado:?\s+(.*?)\s+ng/dL',
    'Anti-TPO': r"(?:ANTI-PEROXIDASE|Anti-tireoperoxidase|Anti-TPO|Anti TPO).+?Resultado:?\s+(.*?)\s+UI/mL",
    'PTH': r'(?:PARATORM[OÔ]NIO|PTH INTACTO).+?Resultado:?\s+(.*?)\s+pg/mL',

    'CT': r'COLESTEROL TOTAL.+?(\d+(?:,\d*)?)(?=\s+mg/dL)',
    'HDL': r'HDL.+?(\d+(?:,\d*)?)(?=\s+mg/dL)(?!\s+mg/dL Resultado)',
    'LDL': r'LDL.+?(\d+(?:,\d*)?)(?=\s+mg/dL)',
    'TG': r'TRIGLIC.+?(\d+(?:,\d*)?)(?=\s+mg/dL)',
    'CPK': r'CREATINO FOSFOQUINASE-CPK.+?Resultado:?\s+(.*?)\s+U/L',

    'AST': r'TGO.+?(\d+(?:,\d*)?)(?=\s+U/L)',
    'ALT': r'TGP.+?(\d+(?:,\d*)?)(?=\s+U/L)',
    'GGT': r'GLUTAMIL.+?Resultado.*?(\d+(?:,\d*)?)',
    'FA': r'FOSFATASE ALCALINA.+?Resultado.*?(\d+(?:,\d*)?)',
    'BT': r'(?<!Ausentes )BILIRRUBINA.+?(?<=Método).*?Total.*?:?\s+(\d+(?:,\d*)?)\s+mg/dL',
    'BD': r'(?<!Ausentes )BILIRRUBINA.+?Direta.*?:?\s+(\d+(?:,\d*)?)\s+mg/dL',
    'BI': r'(?<!Ausentes )BILIRRUBINA.+?Indireta.*?:?\s+(\d+(?:,\d*)?)\s+mg/dL',
    'TP': r'PROTROMBINA.+?PROTROMBINA:?\s+(.*?)\s+seg',
    'INR': r'PROTROMBINA.+?RNI:?\s*(\S+)',
    'TTPA': r'TROMBOPLASTINA.+?Parcial:?\s+(.*?)\s+segundos',
    'Proteínas totais': r'PROTE[IÍ]NAS TOTAIS E.+?Prote[ií]nas totais:?\s+(.*?)\s+g/dL',
    'Albumina': r'PROTE[IÍ]NAS TOTAIS E.+?Albumina.*?(\d+(?:,\d*)?)',
    'Globulina': r'PROTE[IÍ]NAS TOTAIS E.+?Globulina.*?(\d+(?:,\d*)?)',    
    'Relacao A/G': r'PROTE[IÍ]NAS TOTAIS E.+?Relação A.*?(\d+(?:,\d*)?)',
    'Alfa-fetoproteína': r'FETOPROTE.*?Resultado.*?(\d+(?:,\d*)?)',

    'PCR': r'PROTEINA C REATIVA.+?Resultado:?\s+(\d+,?\d*)',
    'VHS': r'HEMOSSEDIMENTACAO.+?Resultado.:?\s+(\d+)\s+mm/1h',
    'AUR': r'ÁCIDO ÚRICO(?!\s+urin[áa]rio).+?Resultado:?\s+(.*?)\s+mg/dL',
    'FR': r'FATOR REUMATOIDE.+?Resultado:?\s+(.*?)\s+UI/mL',
    'FAN': r'FATOR ANTI.+?T[IÍ]TULO\s*:?\s*(\S+)',
    'Padrão': r'FATOR ANTI.+?PADRÃO\s*:\s*([^:]+?)\s*?T[íi]tulo',
    'Anti-núcleo': r'Anticorpos Nucleares\s*:\s*([^:]+?)\s*?Não reagente',
    'Anti-envelope': r'Envelope Nuclear\s*:\s*([^:]+?)\s*?Não reagente',
    'Anti-nucléolo': r'Nucleolares\s*:\s*([^:]+?)\s*?Não reagente',
    'Anti-citoplasma': r'Citoplasm[áa]ticos\s*:\s*([^:]+?)\s*?Não reagente',
    'Anti-mitótico': r'Aparelho Mit[óo]tico\s*:\s*([^:]+?)\s*?Não reagente',
    'Anti-placa metafásica': r'Placa Metaf[áa]sica Cromoss[ôo]mica\s*:\s*([^:]+?)\s*?Não reagente',
    'IgA total':r'Imunog.*?IgA.*?Resultado.*?(\d+(:?.\d*)?)',
    'IgG total':r'Imunog.*?IgG.*?Resultado.*?(\d+(:?.\d*)?)',
    'Eletroforese de proteínas-PT':r'Eletroforese[ | de ]prote[ií]na.*?Prote[ií]nas totais.*?(\d+(?:,\d*)?)',
    'Eletroforese-Alb':r'Eletroforese[ | de ]prote[ií]na.*?Albumina.*?(\d+(?:,\d*)?)',
    'Alfa-1-globulina':r'Eletroforese[ | de ]prote[ií]na.*?Alfa 1.*?(\d+(?:,\d*)?)',
    'Alfa-2-globulina':r'Eletroforese[ | de ]prote[ií]na.*?Alfa 2.*?(\d+(?:,\d*)?)',
    'Beta-1-globulina':r'Eletroforese[ | de ]prote[ií]na.*?Beta 1.*?(\d+(?:,\d*)?)',
    'Beta-2-globulina':r'Eletroforese[ | de ]prote[ií]na.*?Beta 2.*?(\d+(?:,\d*)?)',
    'Gama-globulina':r'Eletroforese[ | de ]prote[ií]na.*?Gama.*?(\d+(?:,\d*)?)',
    'Pico monoclonal':r'Eletroforese[ | de ]prote[ií]na.*?Pico mono.*?(?<=:)(.*?)g/dL',
    'Eletroforese-A/G':r'Eletroforese[ | de ]prote[ií]na.*?Relação A.*?(\d+(?:,\d*)?)',

    'B12': r'VITAMINA B12.+?(\d+(?:,\d*)?)(?=\s+pg/mL)',
    'Ác. fólico': r'ACIDO FOLICO.+?Resultado:?\s*(.*?)\s+ng/mL',
    'Vit. D': r'25 HIDROXI.*?Resultado.*?(\d+(?:,\d*)?)(?=\s+ng/mL)',
 
    'Testo total': r'TESTOSTERONA TOTAL.+?Resultado:\s+(.*?)\s+ng/dL',
    'Testo livre': r'TESTOSTERONA LIVRE.+?Testosterona Livre:\s+(.*?)\s+ng/dL',
    'SHBG': r'TESTOSTERONA LIVRE.+?SHBG:\s+(.*?)\s+nmol/L',
    'PRL': r'PROLACTINA.+?Resultado:?\s+(.*?)\s+ng/mL',
    'Estradiol': r'ESTRADIOL.+?Resultado:?\s+(.*?)\s+pg/mL',
    'FSH': r'FOL[ÍI]CULO ESTIMULANTE.+?Resultado:?\s+(.*?)\s+mUI/mL',
    'LH': r'HORM[OÔ]NIO LUTEINIZANTE.+?Resultado:?\s+(.*?)\s+mUI/mL',

    'VDRL': r'VDRL.+?Resultado:?\s*(\S+)',
    'HBsAg': r'HBsAg.+?Resultado:?\s*(\S+)',
    'HCV': r'HCV.+?Resultado:?\s*(\S+)',
    'Anti-HIV': r'HIV.+?(?:Resultado:|Leitura:)\s*(\S+)',
    'HAV-IgG': r'Hepatite A IgG.+?(?:Resultado|Leitura):?\s*(\S+)',
    'HAV-IgM': r'Hepatite A IgM.+?(?:Resultado|Leitura):?\s*(\S+)',
}

eas = {
    'EAS-pH': r'URINA.+?pH\s*:?\s*(\S+)',
    'Proteinúria': r'URINA.+?PROTE[IÍ]NAS\s*:?\s*(\S+)',
    'Glicosúria': r'URINA.+?GLICOSE\s*:?\s*(\S+)',
    'Cetonúria': r'URINA.+?CORPOS CETONICOS\s*:?\s*(\S+)',
    'Cetonuria': r'URINA.+?CETONAS\s*:?\s*(\S+)',
    #'Bilirrubinúria': r'URINA.+?BILIRRUBINA\s*:?\s*(\S+)',
    'Bilirrubinuria': r'URINA.+?biliares\s*(\S+)',
    'Nitrito': r'URINA.+?NITRITO\s*:?\s*(\S+)',
    'Leucocitúria': r'URINA.+?LEUC[OÓ]CITOS\s*:?\s*(\S+)',
    'Hematúria': r'URINA.+?HEMACIAS\s*:?\s*(\S+)',
    'Hematuria': r'URINA.+?ERITR[OÓ]CITOS\s*:?\s*(\S+)',
    'Bacteriúria': r'URINA.+?FLORA BACTERIANA\s*:?\s*(\S+)',
    'Leveduras': r'URINA.+?LEVEDURAS\s*:?\s*(\S+)',
    'Escamosas': r'URINA.+?C[EÉ]LULAS EPITELIAIS ESCAMOSAS\s*:?\s*(\S+)',
    'Não-escamosas': r'URINA.+?C[EÉ]LULAS EPITELIAIS NÃO ESCAMOSAS\s*:?\s*(\S+)',
    'Cristais': r'URINA.+?CRISTAIS\s*:?\s*(\S+)',
    'Cilindros': r'URINA.+?CILINDROS\s*:?\s*(\S+)',
    'Muco': r'URINA.+?FILAMENTOS? DE MUCO\s*:?\s*(\S+)',
    'Bacteriuria': r'URINA.+?Bact[ée]rias\s*:?\s*(\S+)',
}

reference_values = {
    'Hb': (12.0, 16.0), 
    'Ht': (35.0, 50.0), 
    'VCM': (80.0, 100.0), 
    'HCM': (26.0, 34.0), 
    'RDW': (10.0, 15.0), 
    'Leuco': (4500.0, 11000.0), 
    'Bast': (0.0, 0.0), 
    'Segm': (2000.0, 7500.0),
    'Eosi': (100.0, 400.0),
    'Linf': (900.0, 4400.0),
    'Plaq': (150000.0, 400000.0),

    'Retic': (0.5, 2.0),
    'Ferro': (50.0, 175.0),
    'Ferritina': (15.0, 446.0),
    'Transferrina': (200.0, 360.0),
    'TSAT': (15.0, 50.0),
    'TIBC': (250.0, 425.0),

    'PSAT': (0.0, 2.5),
    'R': (20.0, 60.0),

    'Cr': (0.0, 1.2),
    'Clearance Cr. corrigido': (90.0, 150.0),
    'Albuminúria 24h': (0.0, 30.0),
    'Ur': (10.0, 50.0),
    'Cálcio total': (8.6, 10.3),
    'Cálcio iônico': (1.11, 1.4),
    'Magnésio': (1.6, 2.6),
    'Fósforo': (2.5, 4.5),
    'Na': (136, 145),
    'K': (3.5, 5.1),

    'GJ': (0, 99.9),
    "HbA1c": (0, 5.69),
 
    'TSH': (0.45, 4.5),
    'T4L': (0.9, 1.8),
    'Anti-TPO': (0, 9),
    'PTH': (16, 84),

    'CT': (0, 200),
    'HDL': (40, 120),
    'LDL': (0, 100),
    'TG': (0, 149),
    'CPK': (0, 174),

    'AST': (0, 40),
    'ALT': (0, 41),
    'GGT': (8, 73),
    'FA': (35, 129),
    'BT': (0.2, 1.10),
    'BD': (0, 0.3),
    'BI': (0.2, 0.8),
    'TP': (10.1, 12.8),
    'INR': (0.9, 1.1),
    'TTPA': (25.4, 33.4),
    'Proteínas totais': (6.8, 8.1),
    'Albumina': (3.5, 5.2),
    'Globulina': (1.7, 3.5),
    'Relacao A/G': (0.9, 2.0),
    'Alfa-fetoproteína': (0, 7.5),

    'PCR': (0, 5.0),
    'VHS': (2, 36),
    'AUR': (0, 8),
    'FR': (0, 14),
    'IgA total': (50, 400),
    'IgG total': (600, 1500),
    'Eletroforese de proteínas-PT': (6.5, 8.1), #100%
    'Eletroforese-Alb': (3.5, 4.85), # 54 a 60%
    'Alfa-1-globulina': (0.22, 0.43), # 3,4 a 5,3%
    'Alfa-2-globulina': (0.55, 1.08), # 8,4 a 13,3%
    'Beta-1-globulina': (0.32, 0.54), # 4,9 a 6,6%
    'Beta-2-globulina': (0.24, 0.54), # 3,6 a 6,6%
    'Gama-globulina': (0.74, 1.75), # 11,4 a 21,6%

    'B12': (300, 1500),
    'Ác. fólico': (4, 30),
    'Vit. D': (20, 60),
 
    'Testo total': (350, 816), #Feminino até 50
    'Testo livre': (6.5, 20),
    'SHBG': (18, 77), # Feminino de 27 a 128
    'PRL': (0, 31), # Masculino até 24
    'Estradiol': (39, 440), # Homens até 29
    'FSH': (0, 25), # Menopausa acima de 30 // Homens até 10
    'LH': (0, 100) # Menopausa acima de 15 // Homens até 9
}


# Caminho para o arquivo PDF.
#pdf_path = 'Resultado.pdf'
pdf_path = easygui.fileopenbox(title="Selecione o arquivo PDF", filetypes=["*.pdf"])

#for i in range(1, 10):
#    pdf_path = 'C:\\Users\\rnahu\\Downloads\\' + str(i) + '.pdf'

# Extraindo o texto do PDF.
with open(pdf_path, 'rb') as file:
    pdf = PdfReader(file)
    text = ' '.join(page.extract_text() for page in pdf.pages)

# Processando o texto: substituindo múltiplos espaços em branco por um único espaço.
text = re.sub(r'\s+', ' ', text)
print(text)

match = re.search(r"Data Coleta:\s+([\d/]+)", text)
if match:
    cadastro_date = match.group(1)
else:
    cadastro_date = "DATA"

# Extraindo os valores dos exames.
results = {exam: extract_value(text, exam, pattern) for exam, pattern in exams.items()}
results_eas = {exam: extract_value(text, exam, pattern) for exam, pattern in eas.items()}

# Verificar se os valores necessários foram extraídos corretamente e calcular LDL
if all(key in results and results[key] != 'Não encontrado' for key in ['CT', 'HDL', 'TG']):
    try:
        ldl_value = round(convert_to_float(results['CT']) - convert_to_float(results['HDL']) - (convert_to_float(results['TG']) / 5))
        results['LDL'] = str(ldl_value)  # Atualize o valor de 'LDL' no dicionário de resultados
    except ValueError:  # Caso algum valor não possa ser convertido para float
        pass

# Verificar se os valores necessários foram extraídos corretamente e calcular TSAT
if all(key in results and results[key] != 'Não encontrado' for key in ['Ferro', 'Transferrina']):
    try:
        TSAT_value = round((convert_to_float(results['Ferro']) / convert_to_float(results['Transferrina'])) * 70.9)
        results['TSAT'] = str(TSAT_value)  # Atualize o valor de 'LDL' no dicionário de resultados
    except ValueError:  # Caso algum valor não possa ser convertido para float
        pass

# Verificar e atualizar os resultados de sorologias
sorologias = ['VDRL', 'Anti-HIV', 'HCV', 'HBsAg']
for key in results:
    if key in sorologias and results[key] == 'Não':
        results[key] = 'NR'

# Montando a string do laudo.
laudo_header = f"Exames laboratoriais {cadastro_date}: "

# Extraindo os valores dos exames e formatando
laudo_parts_rtf = []
laudo_parts_html = []
eas_parts = []

for exam, result in results.items():
    if result not in ['Não encontrado', 'Não reagente']:
        try:
            value = convert_to_float(result)
            formatted_value_rtf = format_value(exam, value, 'rtf')
            formatted_value_html = format_value(exam, value, 'html')
            laudo_parts_rtf.append(f"{exam}: {formatted_value_rtf}")
            laudo_parts_html.append(f"{exam}: {formatted_value_html}")
        except ValueError:
            laudo_parts_rtf.append(f"{exam}: {result}")
            laudo_parts_html.append(f"{exam}: {result}")

for exam, result in results_eas.items():
    if result not in ['Negativo', 'AUSENTES', 'Ausentes', 'Ausente', 'Não encontrado']:
        eas_parts.append(f"{exam}: {result}")

# Montando os laudos
rtf_body = remove_isolated_zeros(laudo_header + " // ".join(laudo_parts_rtf))
html_body = remove_isolated_zeros(laudo_header + " // ".join(laudo_parts_html))

# RTF completo
rtf_header = r"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss Arial;}}\viewkind4\uc1\pard\f0\fs20 "
rtf_footer = r"}"
rtf_text = (rtf_header + rtf_body + rtf_footer).encode('utf-8')

# Texto simples
plain_text = rtf_body.encode('utf-8')

# Texto HTML completo
html_header = "<html><body>"
html_footer = "</body></html>"
html_text = format_html_for_clipboard(html_body)

# Copiar os textos para a Área de Transferência
copy_to_clipboard(html_text, rtf_text, plain_text)