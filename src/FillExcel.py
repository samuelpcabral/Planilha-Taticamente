from openpyxl import load_workbook
from pandas import read_html


class CreateExcel:

    def __init__(self):
        pass

    def read_attributes_html(self, html_file):
        player_attributes = dict()
        tables = read_html(html_file, encoding='UTF-8')
        table_tecnicos = tables[0]
        dict_tecnicos = {
            'Cabeceamento': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Cabeceamento", "Unnamed: 2"]),
            'Cantos': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Cantos", "Unnamed: 2"]),
            'Cruzamentos': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Cruzamentos", "Unnamed: 2"]),
            'Desarme': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Desarme", "Unnamed: 2"]),
            'Finalização': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Finalização", "Unnamed: 2"]),
            'Finta': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Finta", "Unnamed: 2"]),
            'Lançamentos Longos': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Lançamentos Longos", "Unnamed: 2"]),
            'Livres': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Livres", "Unnamed: 2"]),
            'Marcação': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Marcação", "Unnamed: 2"]),
            'Marcação de Penáltis': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Marcação de Penáltis", "Unnamed: 2"]),
            'Passe': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Passe", "Unnamed: 2"]),
            'Primeiro Toque': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Primeiro Toque", "Unnamed: 2"]),
            'Remates de Longe': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Remates de Longe", "Unnamed: 2"]),
            'Técnica': int(table_tecnicos.loc[table_tecnicos["Técnicos"] == "Técnica", "Unnamed: 2"])
        }
        player_attributes['Tecnicos'] = dict_tecnicos
        table_mentais = tables[1]
        dict_mentais = {
            'Agressividade': int(table_mentais.loc[table_mentais["Mentais"] == "Agressividade", "Unnamed: 2"]),
            'Antecipação': int(table_mentais.loc[table_mentais["Mentais"] == "Antecipação", "Unnamed: 2"]),
            'Bravura': int(table_mentais.loc[table_mentais["Mentais"] == "Bravura", "Unnamed: 2"]),
            'Compostura': int(table_mentais.loc[table_mentais["Mentais"] == "Compostura", "Unnamed: 2"]),
            'Concentração': int(table_mentais.loc[table_mentais["Mentais"] == "Concentração", "Unnamed: 2"]),
            'Decisões': int(table_mentais.loc[table_mentais["Mentais"] == "Decisões", "Unnamed: 2"]),
            'Determinação': int(table_mentais.loc[table_mentais["Mentais"] == "Determinação", "Unnamed: 2"]),
            'Imprevisibilidade': int(table_mentais.loc[table_mentais["Mentais"] == "Imprevisibilidade", "Unnamed: 2"]),
            'Índice de Trabalho': int(table_mentais.loc[table_mentais["Mentais"] == "Índice de Trabalho", "Unnamed: 2"]),
            'Liderança': int(table_mentais.loc[table_mentais["Mentais"] == "Liderança", "Unnamed: 2"]),
            'Posicionamento': int(table_mentais.loc[table_mentais["Mentais"] == "Posicionamento", "Unnamed: 2"]),
            'Sem Bola': int(table_mentais.loc[table_mentais["Mentais"] == "Sem Bola", "Unnamed: 2"]),
            'Trabalho de Equipa': int(table_mentais.loc[table_mentais["Mentais"] == "Trabalho de Equipa", "Unnamed: 2"]),
            'Visão de Jogo': int(table_mentais.loc[table_mentais["Mentais"] == "Visão de Jogo", "Unnamed: 2"])
        }
        player_attributes['Mentais'] = dict_mentais
        table_fisicos = tables[2]
        dict_fisicos = {
            "Aceleração": int(table_fisicos.loc[table_fisicos["Físicos"] == "Aceleração", "Unnamed: 2"]),
            "Agilidade": int(table_fisicos.loc[table_fisicos["Físicos"] == "Agilidade", "Unnamed: 2"]),
            "Aptidão Física": int(table_fisicos.loc[table_fisicos["Físicos"] == "Aptidão Física", "Unnamed: 2"]),
            "Equilíbrio": int(table_fisicos.loc[table_fisicos["Físicos"] == "Equilíbrio", "Unnamed: 2"]),
            "Força": int(table_fisicos.loc[table_fisicos["Físicos"] == "Força", "Unnamed: 2"]),
            "Impulsão": int(table_fisicos.loc[table_fisicos["Físicos"] == "Impulsão", "Unnamed: 2"]),
            "Resistência": int(table_fisicos.loc[table_fisicos["Físicos"] == "Resistência", "Unnamed: 2"]),
            "Velocidade": int(table_fisicos.loc[table_fisicos["Físicos"] == "Velocidade", "Unnamed: 2"])
        }
        player_attributes['Fisicos'] = dict_fisicos
        return player_attributes

    def fill_excel(self, template, new_name, attributes):
        wb = load_workbook(template)
        atributos = wb['Atributos']
        wb.active = atributos
        atributos['B3'] = attributes['Tecnicos']['Cabeceamento']
        atributos['B4'] = attributes['Tecnicos']['Cantos']
        atributos['B5'] = attributes['Tecnicos']['Cruzamentos']
        atributos['B6'] = attributes['Tecnicos']['Desarme']
        atributos['B7'] = attributes['Tecnicos']['Finalização']
        atributos['B8'] = attributes['Tecnicos']['Finta']
        atributos['B9'] = attributes['Tecnicos']['Lançamentos Longos']
        atributos['B10'] = attributes['Tecnicos']['Livres']
        atributos['B11'] = attributes['Tecnicos']['Marcação']
        atributos['B12'] = attributes['Tecnicos']['Marcação de Penáltis']
        atributos['B13'] = attributes['Tecnicos']['Passe']
        atributos['B14'] = attributes['Tecnicos']['Primeiro Toque']
        atributos['B15'] = attributes['Tecnicos']['Remates de Longe']
        atributos['B16'] = attributes['Tecnicos']['Técnica']

        atributos['B18'] = attributes['Mentais']['Agressividade']
        atributos['B19'] = attributes['Mentais']['Antecipação']
        atributos['B20'] = attributes['Mentais']['Bravura']
        atributos['B21'] = attributes['Mentais']['Compostura']
        atributos['B22'] = attributes['Mentais']['Concentração']
        atributos['B23'] = attributes['Mentais']['Decisões']
        atributos['B24'] = attributes['Mentais']['Determinação']
        atributos['B25'] = attributes['Mentais']['Imprevisibilidade']
        atributos['B26'] = attributes['Mentais']['Índice de Trabalho']
        atributos['B27'] = attributes['Mentais']['Liderança']
        atributos['B28'] = attributes['Mentais']['Posicionamento']
        atributos['B29'] = attributes['Mentais']['Sem Bola']
        atributos['B30'] = attributes['Mentais']['Trabalho de Equipa']
        atributos['B31'] = attributes['Mentais']['Visão de Jogo']

        atributos['B33'] = attributes['Fisicos']['Aceleração']
        atributos['B34'] = attributes['Fisicos']['Agilidade']
        atributos['B35'] = attributes['Fisicos']['Aptidão Física']
        atributos['B36'] = attributes['Fisicos']['Equilíbrio']
        atributos['B37'] = attributes['Fisicos']['Força']
        atributos['B38'] = attributes['Fisicos']['Impulsão']
        atributos['B39'] = attributes['Fisicos']['Resistência']
        atributos['B40'] = attributes['Fisicos']['Velocidade']

        wb.save(f'{new_name}.xlsx')


def create_sheets(template, files):
    ce = CreateExcel()
    for file in files:
        attributes = ce.read_attributes_html(file)
        filename = file.rsplit('/')
        new_filename = filename[-1].replace('.html', '')
        ce.fill_excel(template, new_filename, attributes)
    return str(len(files))
