import math
import matplotlib.pyplot as plt
import pandas as pd
import os
import glob
from pathlib import Path
from docx import Document
import re

def convert_angle(angle):
    """Converte ângulo negativo para positivo equivalente (ex: -20 -> 340)"""
    return angle if angle >= 0 else 360 + angle

def detect_file_type(filename):
    """Detecta o tipo de arquivo baseado na extensão"""
    ext = Path(filename).suffix.lower()
    if ext in ['.xls', '.xlsx']:
        return 'excel'
    elif ext == '.docx':
        return 'word'
    else:
        return 'unknown'

def extract_ht_from_obs(obs, est, pv):
    """Extrai a altura total (HT) das observações"""
    if not obs:
        return 0
    
    # Procura por padrões como "Ht. T0 = 6m", "Ht T0=6m", etc.
    patterns = [
        r'Ht\.?\s*([A-Za-z0-9]+)\s*=\s*([0-9.,]+)\s*m',
        r'Ht\.?\s*=\s*([0-9.,]+)\s*m',
        r'Ht\s+([0-9.,]+)\s*m',
        r'Ht\s*([0-9.,]+)m'
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, obs, re.IGNORECASE)
        for match in matches:
            if len(match) == 2:
                point, ht_value = match
                # Se o ponto mencionado é o EST ou PV atual, usa esse valor
                if point == est or point == pv:
                    try:
                        return float(ht_value.replace(',', '.'))
                    except:
                        pass
            elif len(match) == 1:
                # Padrão sem especificar ponto, assume-se que é para o ponto atual
                try:
                    return float(match[0].replace(',', '.'))
                except:
                    pass
    
    # Se não encontrou HT específico, retorna 0
    return 0

def parse_excel_file(filename):
    """Lê o arquivo Excel e extrai os dados da planilha"""
    try:
        # Lê a planilha Excel
        df = pd.read_excel(filename, sheet_name='Plan1')
        
        # Remove linhas completamente vazias
        df = df.dropna(how='all')
        
        data = []
        
        for index, row in df.iterrows():
            # Pula o cabeçalho
            if str(row.iloc[0]).strip() == 'EST.':
                continue
                
            # Extrai os dados das colunas
            est = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            pv = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            
            # Verifica se é uma linha válida
            if not est or not pv:
                continue
                
            try:
                # DI - Distância Inclinada (coluna D)
                di = float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0
            except:
                di = 0
                
            try:
                # ALFA C - Ângulo vertical (coluna E)
                alpha_c = float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0
            except:
                alpha_c = 0
                
            # Extrai HB da coluna K
            try:
                hb = float(row.iloc[10]) if pd.notna(row.iloc[10]) else 0
            except:
                hb = 0
                
            # Extrai HT das observações (coluna L)
            obs = str(row.iloc[11]) if pd.notna(row.iloc[11]) else ""
            ht = extract_ht_from_obs(obs, est, pv)
            
            data.append({
                'EST': est,
                'PV': pv,
                'DI': di,
                'αc': alpha_c,
                'HT': ht,
                'HB': hb
            })
        
        return data
    except Exception as e:
        print(f"Erro ao ler arquivo Excel {filename}: {str(e)}")
        return []

def parse_word_file(filename):
    """Lê o arquivo Word e extrai os dados da tabela"""
    try:
        doc = Document(filename)
        table = doc.tables[0]
        
        data = []
        
        for i, row in enumerate(table.rows[1:], 1):  # Começa da linha 2 para pular o cabeçalho
            cells = row.cells
            
            # Verifica se tem células suficientes
            if len(cells) < 14:
                continue
                
            est = cells[0].text.strip()
            pv = cells[1].text.strip()
            
            # Verifica se é uma linha válida
            if not est or not pv or est == "EST.":
                continue
            
            # Processa ângulo αc (colunas 3 e 4) - positivo e negativo
            try:
                angulo_pos_text = cells[3].text.strip()
                angulo_neg_text = cells[4].text.strip()
                
                # Prioriza ângulo positivo se disponível, senão usa o negativo
                if angulo_pos_text:
                    alpha_c = float(angulo_pos_text)
                elif angulo_neg_text:
                    alpha_c = -float(angulo_neg_text)  # Converte para negativo
                else:
                    alpha_c = 0
            except:
                alpha_c = 0
                
            try:
                di = float(cells[6].text.strip() or 0)  # DI é a coluna 6
            except:
                di = 0
                
            try:
                ht = float(cells[12].text.strip() or 0)
            except:
                ht = 0
                
            try:
                hb = float(cells[13].text.strip() or 0)
            except:
                hb = 0

            data.append({
                'EST': est,
                'PV': pv,
                'DI': di,
                'αc': alpha_c,
                'HT': ht,
                'HB': hb
            })
        
        return data
    except Exception as e:
        print(f"Erro ao ler arquivo Word {filename}: {str(e)}")
        return []

def parse_file(filename):
    """Lê o arquivo baseado no tipo detectado"""
    file_type = detect_file_type(filename)
    
    if file_type == 'excel':
        return parse_excel_file(filename)
    elif file_type == 'word':
        return parse_word_file(filename)
    else:
        print(f"Tipo de arquivo não suportado: {filename}")
        return []

def calculate_coordinates(data):
    coordinates = {}
    connections = []
    
    # Primeiro, coletamos todos os dados de todos os pontos
    point_data = {}
    
    for row in data:
        est = row['EST']
        pv = row['PV']
        
        # Armazena dados do ponto EST
        point_data[est] = {
            'ht': row['HT'],
            'hb': row['HB']
        }
        
        # Para o PV, se não existe, inicializa com zeros
        if pv not in point_data:
            point_data[pv] = {
                'ht': 0,
                'hb': 0
            }
    
    # Agora, processamos as conexões para calcular coordenadas
    for row in data:
        est = row['EST']
        pv = row['PV']
        di = row['DI']
        angle_deg = convert_angle(row['αc'])
        
        # Se é o primeiro ponto, define como origem
        if est not in coordinates and len(coordinates) == 0:
            coordinates[est] = {
                'x': 0, 
                'y': 0, 
                'ht': point_data[est]['ht'], 
                'hb': point_data[est]['hb']
            }
        
        # Se EST não tem coordenadas, não podemos calcular
        if est not in coordinates:
            continue
            
        # Calcula posição do PV
        x0, y0 = coordinates[est]['x'], coordinates[est]['y']
        angle_rad = math.radians(angle_deg)
        x1 = x0 + di * math.cos(angle_rad)
        y1 = y0 + di * math.sin(angle_rad)
        
        # Se PV não existe nas coordenadas, adiciona com seus próprios dados
        if pv not in coordinates:
            coordinates[pv] = {
                'x': x1, 
                'y': y1, 
                'ht': point_data[pv]['ht'], 
                'hb': point_data[pv]['hb']
            }
        
        connections.append({
            'from': est,
            'to': pv,
            'coords': [(x0, y0), (x1, y1)],
            'DI': di,
            'angle': angle_deg
        })
    
    return coordinates, connections

def plot_topography(coordinates, connections, filename):
    """Gera gráfico para um arquivo específico"""
    if not coordinates:
        print(f"Nenhum ponto para plotar no arquivo {filename}!")
        return
        
    # Cria figura para este arquivo
    plt.figure(figsize=(16, 12))
    
    # Plota conexões horizontais
    for conn in connections:
        x_vals = [conn['coords'][0][0], conn['coords'][1][0]]
        y_vals = [conn['coords'][0][1], conn['coords'][1][1]]
        
        plt.plot(x_vals, y_vals, 'b-', alpha=0.7, linewidth=1)

    # Plotar pontos e linhas verticais HT/HB
    for point, data in coordinates.items():
        x, y = data['x'], data['y']
        ht, hb = data['ht'], data['hb']
        
        # Ponto central
        plt.plot(x, y, 'ro', markersize=4)
        
        # Linha vertical representando HT e HB
        if ht > 0 or hb > 0:
            # Linha vertical completa (do HB ao HT)
            plt.plot([x, x], [y - hb, y + ht], 'k-', linewidth=1.5, alpha=0.7)
            
            # Ponto de HT (teto)
            if ht > 0:
                plt.plot(x, y + ht, '^', color='red', markersize=5, alpha=0.8)
            
            # Ponto de HB (base)
            if hb > 0:
                plt.plot(x, y - hb, 'v', color='blue', markersize=5, alpha=0.8)
        
        # Nome do ponto (menor e mais discreto)
        plt.annotate(point, (x, y), 
                    textcoords="offset points", 
                    xytext=(2, 2),
                    ha='left', 
                    va='bottom', 
                    fontsize=6,
                    alpha=0.8,
                    color='darkred')

    # Adicionar legenda simplificada
    plt.plot([], [], 'ro', markersize=4, label='Ponto de Estação')
    plt.plot([], [], 'k-', linewidth=1.5, label='Altura (HT+HB)')
    plt.plot([], [], 'r^', markersize=5, label='HT (Teto)')
    plt.plot([], [], 'bv', markersize=5, label='HB (Base)')
    plt.plot([], [], 'b-', linewidth=1, label='Conexões')

    plt.grid(True, linestyle='--', alpha=0.3)
    
    # Usa o nome do arquivo no título
    file_title = Path(filename).stem
    plt.title(f'Topografia - {file_title}', fontsize=14, fontweight='bold')
    plt.xlabel('Distância Leste-Oeste (m)')
    plt.ylabel('Distância Norte-Sul (m)')
    plt.axis('equal')
    plt.legend(loc='upper left', fontsize=8)
    plt.tight_layout()
    
    # Mostra o gráfico
    plt.show()

def process_files(folder_path=None, specific_files=None):
    """Processa múltiplos arquivos"""
    files_to_process = []
    
    if specific_files:
        # Processa arquivos específicos fornecidos
        files_to_process = specific_files
    elif folder_path:
        # Processa todos os arquivos na pasta
        extensions = ['*.xls', '*.xlsx', '*.docx']
        for ext in extensions:
            files_to_process.extend(glob.glob(os.path.join(folder_path, ext)))
    else:
        # Procura arquivos no diretório atual
        extensions = ['*.xls', '*.xlsx', '*.docx']
        for ext in extensions:
            files_to_process.extend(glob.glob(ext))
    
    if not files_to_process:
        print("Nenhum arquivo encontrado para processar!")
        return
    
    print(f"Arquivos encontrados: {len(files_to_process)}")
    
    for i, filename in enumerate(files_to_process, 1):
        print(f"\n=== Processando arquivo {i}/{len(files_to_process)}: {filename} ===")
        
        try:
            # Extrai dados do arquivo
            data = parse_file(filename)
            
            if not data:
                print(f"Nenhum dado válido encontrado no arquivo {filename}.")
                continue
                
            print(f"Encontradas {len(data)} linhas de dados válidas")
            
            # Calcula coordenadas
            coordinates, connections = calculate_coordinates(data)
            
            if not coordinates:
                print(f"Nenhuma coordenada foi calculada para {filename}.")
                continue
            
            # Gera gráfico
            plot_topography(coordinates, connections, filename)
            
            # Imprime resumo
            print(f"=== RESUMO PARA {filename} ===")
            print(f"Total de pontos: {len(coordinates)}")
            print(f"Total de medições: {len(data)}")
            
        except Exception as e:
            print(f"Erro ao processar arquivo {filename}: {str(e)}")
            import traceback
            traceback.print_exc()

# Execução principal
if __name__ == "__main__":
    # Opção 1: Processar todos os arquivos na pasta atual
    process_files()
    
    # Opção 2: Processar arquivos em uma pasta específica
    # process_files(folder_path="caminho/para/sua/pasta")
    
    # Opção 3: Processar arquivos específicos
    # process_files(specific_files=["arquivo1.xlsx", "arquivo2.docx"])