from PyPDF2 import PdfReader
import re
import pyperclip
import easygui

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
    'Hb': r'HEMOGLOBINA :(\d+,\d*)\s+g/dl',
    'Ht': r'HEMATÓCRITO :(\d+,\d*)\s+%',
    'VCM': r'V\.?C\.?M\.? :(\d+,\d*)\s+fL',
    'HCM': r'H\.?C\.?M\.? :(\d+,\d*)\s+pg',
    'RDW': r'R.D.W. - CV :(\d+,\d*)\s+%',
    'Retic': r'RETICULOCITOS.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+%',
    'Ferro': r'FERRO S[EÉ]RICO.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+µg/dL',
    'Ferritina': r'FERRITINA.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+ng/mL',
    'Transferrina': r'(?<!saturação de )TRANSFERRINA.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+mg/dL',
    'TIBC': r'CAPACIDADE TOTAL.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+μg/dL',
    'TSAT': r'Saturação de Transferrina.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+%',
    'Leuco': r'Leucócitos\s*:\s*(\d+\.?\d*)',
    'Bast': r'Bastonetes\s*:\s*\d+,\d+\s*(\d+\.?\d*)',
    'Segm': r'entados\s*:\s*\d+,\d+\s*(\d+\.?\d*)',
    'Eosi': r'Eosinófilos\s*:\s*\d+,\d+\s*(\d+\.?\d*)',
    'Linf': r'Linfócitos Típicos\s*:\s*\d+,\d+\s*(\d+\.?\d*)',
    'Plaq': r'Plaquetas :\s+(\d+.\d*)\s+/ uL',
    'PSAT': r'(?:PSA LIVRE/TOTAL|PSA TOTAL ULTRA).+?PSA TOTAL:?\s+(.*?)\s+ng/mL',
    'PSAL': r'PSA LIVRE/TOTAL.+?PSA LIVRE:?\s+(.*?)\s+ng/mL',
    'R': r'PSA LIVRE/TOTAL.+?RELAÇÃO:?.+?(.*?)\s*%',
    'PSOF1': r'SANGUE OCULTO DIRETO.+?Resultado:?\s*(\S+)',
    'PSOF2': r'FEZES 2.+?Resultado:?\s*(\S+)',
    'PSOF3': r'FEZES 3.+?Resultado:?\s*(\S+)',
    'Cr': r'CREATININA.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+mg/dL',
    'Clearance Cr. corrigido': r'CLEARANCE.+?Clearance corrigido:?\s*(\S+)',
    'Albuminúria 24h': r'MICROALBUMINURIA.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+mg/24',
    'Albuminúria Isolada': r'MICROALBUMIN[UÚ]RIA.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+mg/g',
    #'Ur': r'UR[EÉ]IA.+?Resultado:?\s+(.*?)\s+mg/dL',
    'Ur': r'UR[EÉ]IA.+?(\d+,\d*|\d+)(?=\s+mg/dL)',
    'Ca': r'(?<!de )C[AÁ]LCIO(?!\s+urin[áa]rio).+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+mg/dL',
    'Na': r'S[OÓ]DIO.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+mEq/L',
    'K': r'POT[AÁ]SSIO.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+mEq/L',
    #'GJ': r'GLICOSE JEJUM.+?Resultado:?\s+(.*?)\s+mg/dL',
    'GJ': r'Glicemia.+?(\d+(?:,\d*)?)(?=\s+mg/dL)',
    #'GLICADA': r'Hb A1c:\s+(.*?)\s*%',
    #'A1C': r'HEMOGLOBINA GLICADA - A1C:\s+(.*?)',
    #'HbA1c': r'HEMOGLOBINA GLICADA.+?HBGLI\s+(.*?)\s*%',
    "HbA1c": r"(?<=GLICADA)[^%]*?(\d+,\d+)(?=\s*%)",
    #'TSH': r'TSH ULTRA SENSIVEL.+?Resultado:?\s+(.*?)\s+[uµ]UI/mL',
    'TSH': r'TSH.+?(\d+(?:,\d*)?)(?=\s+[uµ]UI/mL)',
    'T4L': r'T4 LIVRE.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+ng /dL',
    'Anti-TPO': r"(?:ANTI-PEROXIDASE|Anti-tireoperoxidase|Anti-TPO|Anti TPO).+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+UI/mL",
    'PTH': r'(?:PARATORM[OÔ]NIO|PTH INTACTO).+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+pg/mL',
    #'CT': r'COLESTEROL TOTAL.*?Resultado:?\s+(\d+(?:,\d*)?)\s+mg/dL',
    'CT': r'COLESTEROL TOTAL.+?(\d+(?:,\d*)?)(?=\s+mg/dL)',
    #'HDL': r'HDL.*?(\d+)\s+mg/dL',
    #'HDL': r'HDL.*?Resultado:?\s+(\d+(?:,\d*)?)\s+mg/dL',
    'HDL': r'HDL.+?(\d+(?:,\d*)?)(?=\s+mg/dL)',
    #'LDL': r'LDL.*?(\d+)\s+mg/dL',
    #'LDL': r'LDL.*?Resultado:?\s+(\d+(?:,\d*)?)\s+mg/dL',
    'LDL': r'LDL.+?(\d+(?:,\d*)?)(?=\s+mg/dL)',
    #'TG': r'TRIGLICERIDEOS.+?RESULTADO:?\s*(\S+)',
    'TG': r'Triglic.+?(\d+(?:,\d*)?)(?=\s+mg/dL)',
    'CPK': r'CK Total.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+U/L',
    #'AST': r'TRANSAMINASE OXALÁCETICA \(TGO/AST\).+?Resultado:?\s+(.*?)\s+U/L',
    'AST': r'TGO.+?(\d+(?:,\d*)?)(?=\s+U/L)',
    #'ALT': r'TRANSAMINASE PIRUVICA \(TGP/ALT\).+?Resultado:?\s+(.*?)\s+U/L',
    'ALT': r'TGP.+?(\d+(?:,\d*)?)(?=\s+U/L)',
    #'GGT': r'GLUTAMIL.+?Resultado:?\s+(.*?)\s+U/L',
    'GGT': r'Gama GT.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+U/L',
    'FA': r'Fosfatase Alcalina.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+U/L',
    'PT': r'PROTE[IÉ]NAS TOTAIS E.+?Prote[ií]nas totais:?\s+(.*?)\s+g/dL',
    'Albumina': r'PROTE[IÉ]NAS TOTAIS E.+?Albumina:?\s+(.*?)\s+g/dL',
    'BT': r'Bilirrubina Total(?!\s+e fra).+?Total\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+mg/dL',
    'BD': r'Bilirrubina Direta.+?Direta\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+mg/dL',
    'BI': r'Bilirrubina Indireta.+?Indireta\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+mg/dL',
    'TP': r'PROTROMBINA.+?PROTROMBINA:?\s+(.*?)\s+seg',
    'INR': r'PROTROMBINA.+?RNI:?\s*(\S+)',
    'TTPA': r'TROMBOPLASTINA.+?Parcial:?\s+(.*?)\s+segundos',
    'PCR': r'PROTEINA C REATIVA.+?Resultado:?\s+(\d+,?\d*)\s+mg/dL',
    'VHS': r'HEMOSSEDIMENTACAO.+?Resultado.:?\s+(\d+)\s+mm/1h',
    'AUR': r'ÁCIDO ÚRICO liberado.+?(\d+(?:,\d*)?)(?=\s+mg/dL)',
    #'FR': r'FATOR REUMATOIDE.+?Resultado:......\s+(.*?)\s+Ul/mL',
    'FR': r'FATOR REUMAT[OÓ]IDE.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+UI/mL',
    'FAN': r'FATOR ANTI.+?T[IÍ]TULO\s*:?\s*(\S+)',
    'Padrão': r'FATOR ANTI.+?PADRÃO\s*:\s*([^:]+?)\s*?T[íi]tulo',
    'Anti-núcleo': r'Anticorpos Nucleares\s*:\s*([^:]+?)\s*?Não reagente',
    'Anti-envelope': r'Envelope Nuclear\s*:\s*([^:]+?)\s*?Não reagente',
    'Anti-nucléolo': r'Nucleolares\s*:\s*([^:]+?)\s*?Não reagente',
    'Anti-citoplasma': r'Citoplasm[áa]ticos\s*:\s*([^:]+?)\s*?Não reagente',
    'Anti-mitótico': r'Aparelho Mit[óo]tico\s*:\s*([^:]+?)\s*?Não reagente',
    'Anti-placa metafásica': r'Placa Metaf[áa]sica Cromoss[ôo]mica\s*:\s*([^:]+?)\s*?Não reagente',
    #'B12': r'VITAMINA B12.+?Resultado:?\s+(.*?)\s+pg/mL',
    'B12': r'VITAMINA B12.+?(\d+(?:,\d*)?)(?=\s+pg/mL)',
    'Ác. fólico': r'[AÁ]CIDO F[OÓ]LICO.+?Resultado\s*:\s*(\d+|\d+,\d+|>\d+|<\d+|\d+.\d+)\s+ng/mL',
    #'Vit.D': r'25 HIDROXI.+?Resultado:?\s+(.*?)\s+ng/mL',
    'Vit. D': r'25 OH.+?(\d+(?:,\d*)?)(?=\s+ng/mL)',
    'Testo total': r'TESTOSTERONA TOTAL.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+ng/dL',
    'Testo livre': r'TESTOSTERONA LIVRE.+?Testosterona Livre\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+ng/dL',
    'SHBG': r'TESTOSTERONA LIVRE.+?SHBG\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+nmol/L',
    'PRL': r'PROLACTINA.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+ng/mL',
    'Estradiol': r'ESTRADIOL.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+pg/mL',
    'FSH': r'FOL[ÍI]CULO ESTIMULANTE.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+mUI/mL',
    'LH': r'HORM[OÔ]NIO LUTEINIZANTE.+?Resultado\s*:\s*(\d+|\d+,\d+|<\d+|\d+.\d+)\s+mUl/mL',
    'VDRL': r'VDRL.+?Interpretação:?\s*(\S+)',
    'Sorologia sífilis': r'Sorologia para sífilis.+?Interpretação:?\s*(\S+)',
    'HBsAg': r'HBsAg.+?Interpretação:?\s*(\S+)',
    'Anti-HBs': r'Anti HBs.+?Interpretação:?\s*(\S+)',
    'Anti-HBc': r'HBc.+?Interpretação:?\s*(\S+)',
    'HCV': r'Hepatite C.+?Interpretação:?\s*(\S+)',
    'Anti-HIV': r'HIV.+?Interpretação:?\s*(\S+)',
}

eas = {
    'EAS-pH': r'URINA.+?pH\s*:?\s*(\S+)',
    'EAS-proteinúria': r'URINA.+?PROTE[IÍ]NAS\s*:?\s*(\S+)',
    'EAS-glicosúria': r'URINA.+?GLICOSE\s*:?\s*(\S+)',
    'EAS-cetonúria': r'URINA.+?CORPOS CETONICOS\s*:?\s*(\S+)',
    'EAS-cetonas': r'URINA.+?CETONAS\s*:?\s*(\S+)',
    'EAS-bilirrubinúria': r'URINA.+?BILIRRUBINA\s*:?\s*(\S+)',
    'EAS-bilirrubina': r'URINA.+?BILIARES\s*:?\s*(\S+)',
    'EAS-nitrito': r'URINA.+?NITRITO\s*:?\s*(\S+)',
    'EAS-leucocitúria': r'URINA.+?LEUC[OÓ]CITOS\s*:?\s*(\S+)',
    'EAS-hematúria': r'URINA.+?HEM[AÁ]CIAS\s*:?\s*(\S+)',
    'EAS-hemácias': r'URINA.+?ERITR[OÓ]CITOS\s*:?\s*(\S+)',
    'EAS-bacteriúria': r'URINA.+?FLORA BACTERIANA\s*:?\s*(\S+)',
    'EAS-leveduras': r'URINA.+?LEVEDURAS\s*:?\s*(\S+)',
    'EAS-escamosas': r'URINA.+?Células\s*:?\s*(\S+)',
    'EAS-não-escamosas': r'URINA.+?Células NÃO ESCAMOSAS\s*:?\s*(\S+)',
    'EAS-oxalato': r'URINA.+?Oxalato de cálcio\s*:?\s*(\S+)',
    'EAS-fosfato': r'URINA.+?Fosfato Triplo\s*:?\s*(\S+)',
    'EAS-urato': r'URINA.+?Ácido Úrico\s*:?\s*(\S+)',
    'EAS-grân. amorfos': r'URINA.+?Grânuos amorfos\s*:?\s*(\S+)',
    'EAS-cilindros hialinos': r'URINA.+?Cil. Hialinos\s*:?\s*(\S+)',
    'EAS-cilindros patológicos': r'URINA.+?Cil. Patológicos\s*:?\s*(\S+)',
    'EAS-muco': r'URINA.+?FIL. DE MUCO\s*:?\s*(\S+)',
    'EAS-bactérias': r'URINA.+?Bact[ée]rias\s*:?\s*(\S+)',
}


# Caminho para o arquivo PDF.
#for i in range(1,21):
#    pdf_path = str(i) + '.pdf'

pdf_path = easygui.fileopenbox(title="Selecione o arquivo PDF", filetypes=["*.pdf"])


# Extraindo o texto do PDF.
with open(pdf_path, 'rb') as file:
    pdf = PdfReader(file)
    text = ' '.join(page.extract_text() for page in pdf.pages)

# Processando o texto: substituindo múltiplos espaços em branco por um único espaço.
text = re.sub(r'\s+', ' ', text)

match = re.search(r"Cadastro:\s+([\d/]+)", text)
if match:
    cadastro_date = match.group(1)
else:
    cadastro_date = "DATA"

# Extraindo os valores dos exames.
results = {exam: extract_value(text, exam, pattern) for exam, pattern in exams.items()}
results_eas = {exam: extract_value(text, exam, pattern) for exam, pattern in eas.items()}

# Verificar se os valores necessários foram extraídos corretamente e calcular o LDL
if all(key in results and results[key] != 'Não encontrado' for key in ['CT', 'HDL', 'TG']):
    try:
        ldl_value = round(convert_to_float(results['CT']) - convert_to_float(results['HDL']) - (convert_to_float(results['TG']) / 5))
        results['LDL'] = str(ldl_value)  # Atualize o valor de 'LDL' no dicionário de resultados
    except ValueError:  # Caso algum valor não possa ser convertido para float
        pass

# Montando a string do laudo.
laudo_parts = [f"{cadastro_date}:"]
for exam, result in results.items():
    if result not in ['Não encontrado', 'Não reagente']:
        laudo_parts.append(f"{exam} {result}")
        
eas_parts = []
for exam, result in results_eas.items():
    if result not in ['Negativo', 'AUSENTES', 'Ausentes', 'Ausente', 'Não encontrado']:
        eas_parts.append(f"{exam} {result}")

laudo = "Exames laboratoriais " + " // ".join(laudo_parts) + " // " + " // ".join(eas_parts)
pyperclip.copy(laudo)

#print(str(i) + ' ' + laudo)

## Para transformar em arquivo executável, digitar este comando num Prompt de Comando (CMD):
## pyinstaller --onefile Exames_Campinas.py