# generate_test_file.py - GERA ARQUIVO DE TESTE COMPATÍVEL (CORRIGIDO)
import pandas as pd
import numpy as np
from datetime import datetime
import os

def gerar_arquivo_teste_compativel():
    """Gera um arquivo Excel com os dados fornecidos, formatado corretamente"""
    
    print("🚀 Gerando arquivo Excel compatível com seu sistema...")
    
    # Dados fornecidos - EXATAMENTE como você mostrou
    data = {
        'ID do Módulo': [
            '0001', '0002', '0003', '0004', '0005', '0006', '0007', '0008', '0009', '0010',
            '0011', '0012', '0013', '0014', '0015', '0016', '0017', '0018', '0019', '0020',
            '0021', '0022', '0023', '0024', '0025', '0026', '0027', '0028', '0029', '0030',
            '0031', '0032', '0033', '0034', '0035', '0036', '0037', '0038', '0039', '0040',
            '0041', '0042', '0043', '0044', '0045', '0046', '0047', '0048', '0049', '0050',
            '0051', '0052', '0053', '0054', '0055', '0056', '0057', '0058', '0059', '0060',
            '0061', '0062', '0063', '0064', '0065', '0066', '0067', '0068', '0069', '0070',
            '0071', '0072', '0073', '0074', '0075', '0076'
        ],
        'NS do Módulo': [
            '820', '821', '822', '823', '824', '825', '826', '827', '828', '829',
            '830', '831', '832', '833', '834', '835', '836', '837', '838', '839',
            '840', '841', '842', '843', '844', '845', '846', '847', '848', '849',
            '850', '851', '852', '853', '854', '855', '856', '857', '858', '859',
            '860', '861', '862', '863', '864', '865', '866', '867', '868', '869',
            '870', '871', '872', '873', '874', '875', 'A1', 'A10', 'A11', 'A12',
            'A13', 'A14', 'A15', 'A16', 'A17', 'A18', 'A19', 'A2', 'A20', 'A3',
            'A4', 'A5', 'A6', 'A7', 'A8', 'A9'
        ],
        'Fabricante': ['Solarex'] * 76,
        'Modelo': (['56W'] * 56) + (['77W'] * 20),
        'Potência do datasheet (W)': ([56.00] * 56) + ([77.00] * 20),
        'Voc Original (V)': ([20.80] * 56) + ([21.00] * 20),
        'Isc Original (A)': ([3.60] * 56) + ([5.00] * 20),
        'Ano': [1998] * 76,
        'Bifacial/Monofacial': ['Monofacial'] * 76,
        'Vidro Quebrado/Rachado?': [
            'Sim', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Sim', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Sim', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não'
        ],
        'Backsheet Danificado?': [
            'Sim', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Sim', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não'
        ],
        'Junction Box Danificado?': [
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Sim', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Sim', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não'
        ],
        'Cabos/Conectores Danificados?': [
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Sim', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não'
        ],
        'Defeito Reparável?': [
            'Não', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'Não', 'NA',
            'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'Não', 'NA', 'NA',
            'NA', 'NA', 'NA', 'NA', 'Não', 'NA', 'NA', 'Não', 'NA', 'NA',
            'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA',
            'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA',
            'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA',
            'NA', 'NA', 'NA', 'NA', 'Não', 'NA', 'NA', 'NA', 'NA', 'NA',
            'NA', 'NA', 'NA', 'NA', 'NA', 'NA'
        ],
        'Altura (m)': ([1.105] * 56) + ([1.461] * 20),
        'Largura (m)': ([0.610] * 56) + ([0.502] * 20),
        'Resistência Medida 1 min (MΩ)': [
            194.00, 189.00, 138.00, 159.00, 129.00, 633.00, 85.00, 0.00, 106.00, 200.00,
            182.00, 168.00, 227.00, 258.00, 282.00, 480.00, 0.00, 328.00, 29.90, 216.00,
            143.00, 308.00, 110.00, 0.00, 326.00, 222.00, 174.00, 127.00, 158.00, 79.70,
            148.00, 630.00, 1.61, 16.30, 158.00, 25.00, 198.00, 253.00, 34.60, 47.30,
            179.00, 444.00, 55.00, 603.00, 600.00, 217.00, 258.00, 184.00, 184.00, 192.00,
            71.60, 167.00, 238.00, 250.00, 352.00, 155.00, 155.00, 167.00, 164.00, 45.90,
            159.00, 20.00, 33.50, 6.00, 122.00, 57.00, 58.00, 44.00, 138.00, 125.00,
            100.00, 14.40, 97.00, 72.00, 94.00, np.nan
        ],
        'Resistência Medida 2 min (MΩ)': [
            227.00, 147.00, 124.00, 166.00, 146.00, 615.00, 94.00, 0.00, 130.00, 251.00,
            151.00, 182.00, 213.00, 250.00, 292.00, 464.00, 0.00, 365.00, 42.20, 209.00,
            167.00, 370.00, 133.00, 0.00, 344.00, 230.00, 168.00, 165.00, 147.00, 78.00,
            73.00, 600.00, 27.00, 44.00, 175.00, 36.00, 207.00, 261.00, 29.40, 40.40,
            180.00, 389.00, 61.00, 544.00, 607.00, 202.00, 257.00, 198.00, 198.00, 173.00,
            78.70, 187.00, 244.00, 261.00, 358.00, 165.00, 165.00, 172.00, 168.00, 19.00,
            166.00, 20.00, 34.70, 0.00, 127.00, 67.50, 63.50, 47.00, 137.00, 128.00,
            108.00, 15.00, 91.00, 76.00, 92.00, np.nan
        ],
        'Resistência Ôhmica Fabricante (MΩ·m²)': [
            130.77, 99.09, 83.58, 107.17, 86.95, 414.54, 57.29, 0.00, 71.45, 134.81,
            101.78, 113.24, 143.57, 168.51, 190.08, 312.76, 0.00, 221.09, 20.15, 140.88,
            96.39, 207.61, 74.15, 0.00, 219.74, 149.64, 113.24, 85.60, 99.09, 52.58,
            49.21, 404.43, 1.09, 10.99, 106.50, 16.85, 133.46, 170.53, 19.82, 27.23,
            120.65, 262.21, 37.07, 366.68, 404.43, 136.16, 173.23, 124.03, 124.03, 116.61,
            48.26, 112.57, 160.42, 168.51, 237.27, 104.48, 113.68, 122.48, 120.28, 13.94,
            116.61, 14.67, 24.57, 0.00, 89.48, 41.81, 42.54, 32.27, 100.48, 91.68,
            73.34, 10.56, 66.74, 52.81, 67.47, np.nan
        ],
        'Idade do Módulo Conhecida?': ['Sim'] * 76,
        'Voc Medido (V)': [
            6.30, 19.70, 19.80, 20.00, 19.60, 19.90, 20.00, 19.70, 19.60, 20.00,
            19.70, 19.60, 19.30, 19.70, 19.60, 20.40, 19.30, '#DIV/0!', 19.80, 20.00,
            19.70, 20.30, 19.80, 19.30, '#DIV/0!', 19.60, 19.70, 19.90, 19.90, 20.00,
            19.80, 19.80, 19.40, 19.90, 19.90, 19.60, 19.60, 19.70, 19.90, 18.90,
            19.50, 20.00, 19.80, 19.80, 19.30, 20.00, 19.60, 20.00, 19.40, 20.40,
            19.70, 19.30, 19.30, 19.60, 19.70, 19.50, 20.10, 20.70, 20.50, 19.20,
            18.80, 20.60, 20.70, 20.30, 8.70, 20.30, 20.50, 20.10, 20.40, 20.20,
            20.60, 20.40, 20.40, 20.20, 20.60, 20.50
        ],
        'Isc Medido (A)': [
            0.00, 3.30, 3.50, 3.60, 3.40, 3.30, 3.50, 3.40, 3.60, 3.40,
            3.40, 3.40, 3.40, 3.40, 3.40, 3.60, 3.50, '#DIV/0!', 3.40, 3.40,
            3.50, 3.50, 3.40, 3.40, '#DIV/0!', 3.50, 3.40, 3.60, 3.70, 3.50,
            3.40, 3.40, 3.40, 3.30, 3.50, 3.40, 3.50, 3.50, 3.50, 3.40,
            3.40, 3.50, 3.40, 3.50, 3.40, 3.70, 3.40, 3.50, 3.40, 3.70,
            3.50, 3.60, 3.50, 3.30, 3.50, 3.40, 4.90, 5.00, 4.90, 4.80,
            4.70, 5.00, 4.80, 5.00, 3.10, 4.90, 5.00, 4.80, 4.90, 4.80,
            5.10, 4.90, 4.90, 5.00, 4.90, 5.00
        ],
        'Pmáx Medido (W)': [
            2.40, 14.50, 14.60, 14.80, 14.20, 14.80, 14.50, 14.50, 13.90, 14.90,
            14.70, 14.50, 13.70, 14.70, 14.30, 15.40, 13.90, '#DIV/0!', 14.70, 14.20,
            14.00, 15.40, 14.50, 13.90, '#DIV/0!', 14.50, 14.80, 15.00, 14.30, 14.60,
            14.80, 14.70, 14.00, 15.20, 14.60, 14.50, 13.90, 14.40, 14.60, 6.10,
            14.00, 14.70, 14.70, 14.80, 13.70, 15.00, 14.60, 14.30, 13.90, 15.40,
            14.60, 14.40, 14.00, 14.70, 14.30, 14.60, 14.30, 15.10, 15.10, 5.90,
            5.30, 14.80, 15.20, 14.70, 5.90, 14.70, 14.60, 14.80, 15.10, 14.60,
            14.80, 14.70, 14.90, 14.40, 14.50, 14.80
        ],
        'Fill Factor Medido (%)': [
            8.00, 64.80, 62.80, 66.00, 62.80, 66.90, 64.70, 64.60, 61.20, 65.90,
            67.50, 65.80, 61.40, 65.60, 62.60, 67.10, 64.30, '#DIV/0!', 65.40, 60.70,
            60.60, 60.70, 63.80, 62.30, '#DIV/0!', 64.30, 68.40, 66.60, 63.00, 62.90,
            66.40, 65.20, 64.40, 64.60, 61.20, 64.50, 61.40, 63.10, 63.10, 28.90,
            63.50, 61.90, 66.30, 66.50, 62.70, 66.30, 66.50, 58.70, 63.60, 66.60,
            65.40, 66.80, 63.90, 65.90, 63.10, 67.30, 61.00, 64.40, 65.10, 26.30,
            23.40, 63.60, 61.60, 62.20, 57.80, 62.80, 63.10, 60.00, 65.00, 63.10,
            63.10, 62.70, 64.50, 61.10, 61.50, 63.80
        ],
        'Potência (% da original)': [
            4.29, 25.89, 26.07, 26.43, 25.36, 26.43, 25.89, 25.89, 24.82, 26.61,
            26.25, 25.89, 24.46, 26.25, 25.54, 27.50, 24.82, 'N/A', 26.25, 25.36,
            25.00, 27.50, 25.89, 24.82, 'N/A', 25.89, 26.43, 26.79, 25.54, 26.07,
            26.43, 26.25, 25.00, 27.14, 26.07, 25.89, 24.82, 25.71, 26.07, 10.89,
            25.00, 26.25, 26.25, 26.43, 24.46, 26.79, 26.07, 25.54, 24.82, 27.50,
            26.07, 25.71, 25.00, 26.25, 25.54, 26.07, 18.57, 19.61, 19.61, 7.66,
            6.88, 19.22, 19.74, 19.09, 7.66, 19.09, 18.96, 19.22, 19.61, 18.96,
            19.22, 19.09, 19.35, 18.70, 18.83, 19.22
        ],
        'Fill Factor Original (%)': [74.79] * 56 + [73.33] * 20,
        'Foi realizado Eletroluminescência?': ['Sim'] * 76,
        'Rachaduras Detectadas?': [
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Sim', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Sim', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Sim', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não'
        ],
        '>50% Células Danificadas?': [
            'Não', 'Não', 'Sim', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Sim', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não',
            'Não', 'Não', 'Não', 'Não', 'Não', 'Não'
        ]
    }
    
    # Criar DataFrame
    df = pd.DataFrame(data)
    
    # Ordem das colunas EXATAMENTE como seu sistema espera
    colunas_ordenadas = [
        'ID do Módulo', 'NS do Módulo', 'Fabricante', 'Modelo', 
        'Potência do datasheet (W)', 'Voc Original (V)', 'Isc Original (A)',
        'Ano', 'Bifacial/Monofacial',
        'Vidro Quebrado/Rachado?', 'Backsheet Danificado?', 
        'Junction Box Danificado?', 'Cabos/Conectores Danificados?',
        'Defeito Reparável?', 
        'Altura (m)', 'Largura (m)',
        'Resistência Medida 1 min (MΩ)', 'Resistência Medida 2 min (MΩ)', 
        'Resistência Ôhmica Fabricante (MΩ·m²)',
        'Idade do Módulo Conhecida?',
        'Voc Medido (V)', 'Isc Medido (A)', 'Pmáx Medido (W)', 
        'Fill Factor Medido (%)', 'Potência (% da original)',
        'Fill Factor Original (%)',
        'Foi realizado Eletroluminescência?', 'Rachaduras Detectadas?', 
        '>50% Células Danificadas?'
    ]
    
    df = df[colunas_ordenadas]
    
    # Nome do arquivo
    output_file = "Planilha_Teste_Completa_76_Modulos.xlsx"
    
    # Salvar em Excel COM FORMATAÇÃO DE CABEÇALHO
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Adicionar linhas de cabeçalho como no seu sistema
        workbook = writer.book
        
        # Criar sheet
        sheet_name = "Análise Módulos FV"
        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
        
        worksheet = writer.sheets[sheet_name]
        
        # Adicionar primeira linha de cabeçalho (etapas)
        etapas = [
            "Informações Gerais", "Informações Gerais", "Informações Gerais", 
            "Informações Gerais", "Informações Gerais", "Informações Gerais", 
            "Informações Gerais", "Informações Gerais", "Informações Gerais",
            "Inspeção Visual", "Inspeção Visual", "Inspeção Visual", 
            "Inspeção Visual", "Inspeção Visual",
            "Resistência de Isolamento", "Resistência de Isolamento", 
            "Resistência de Isolamento", "Resistência de Isolamento", 
            "Resistência de Isolamento",
            "Teste Curva IV", "Teste Curva IV", "Teste Curva IV", 
            "Teste Curva IV", "Teste Curva IV", "Teste Curva IV", 
            "Teste Curva IV", "Teste Curva IV",
            "Eletroluminescência", "Eletroluminescência", "Eletroluminescência"
        ]
        
        # Adicionar linha de etapas (linha 1)
        for col_idx, etapa in enumerate(etapas, 1):
            worksheet.cell(row=1, column=col_idx, value=etapa)
        
        # Ajustar largura das colunas
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 30)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print("=" * 70)
    print("✅ ARQUIVO EXCEL GERADO COM SUCESSO!")
    print("=" * 70)
    print(f"📁 Nome do arquivo: {output_file}")
    print(f"📊 Total de módulos: {len(df)}")
    print(f"📏 Área típica: {df['Altura (m)'].iloc[0]:.3f}m × {df['Largura (m)'].iloc[0]:.3f}m = {(df['Altura (m)'].iloc[0] * df['Largura (m)'].iloc[0]):.4f}m²")
    
    print("\n📋 ESTATÍSTICAS DOS DADOS:")
    print("-" * 40)
    
    # Análise dos dados
    print(f"1. Inspeção Visual:")
    print(f"   • Vidro quebrado: {(df['Vidro Quebrado/Rachado?'] == 'Sim').sum()} módulos")
    print(f"   • Backsheet danificado: {(df['Backsheet Danificado?'] == 'Sim').sum()} módulos")
    print(f"   • Junction box danificado: {(df['Junction Box Danificado?'] == 'Sim').sum()} módulos")
    print(f"   • Cabos danificados: {(df['Cabos/Conectores Danificados?'] == 'Sim').sum()} módulos")
    
    print(f"\n2. Teste de Resistência:")
    # Converter valores para numérico, tratando erros
    df['Resistência Medida 1 min (MΩ)'] = pd.to_numeric(df['Resistência Medida 1 min (MΩ)'], errors='coerce')
    df['Resistência Medida 2 min (MΩ)'] = pd.to_numeric(df['Resistência Medida 2 min (MΩ)'], errors='coerce')
    
    r1_media = df['Resistência Medida 1 min (MΩ)'].mean()
    r2_media = df['Resistência Medida 2 min (MΩ)'].mean()
    print(f"   • Resistência 1 min média: {r1_media:.1f} MΩ")
    print(f"   • Resistência 2 min média: {r2_media:.1f} MΩ")
    print(f"   • Módulos com resistência zero: {(df['Resistência Medida 1 min (MΩ)'] == 0).sum()}")
    
    print(f"\n3. Teste Curva IV:")
    # Converter potência para numérico
    potencia_numerica = []
    for val in df['Potência (% da original)']:
        if isinstance(val, (int, float)):
            potencia_numerica.append(float(val))
        elif isinstance(val, str) and val.replace('.', '', 1).isdigit():
            potencia_numerica.append(float(val))
        else:
            potencia_numerica.append(0.0)
    
    potencia_media = np.mean(potencia_numerica)
    potencia_min = np.min([p for p in potencia_numerica if p > 0])
    potencia_max = np.max(potencia_numerica)
    
    print(f"   • Potência média: {potencia_media:.1f}%")
    print(f"   • Potência mínima: {potencia_min:.1f}%")
    print(f"   • Potência máxima: {potencia_max:.1f}%")
    
    print(f"\n4. Eletroluminescência:")
    print(f"   • Rachaduras detectadas: {(df['Rachaduras Detectadas?'] == 'Sim').sum()} módulos")
    print(f"   • >50% células danificadas: {(df['>50% Células Danificadas?'] == 'Sim').sum()} módulos")
    
    print(f"\n🎯 RESULTADOS ESPERADOS (estimativa):")
    print("-" * 40)
    
    # Simulação simplificada da lógica
    aprovados_a = 0
    aprovados_b = 0
    reciclagem = 0
    manutencao = 0
    
    for idx, row in df.iterrows():
        # 1. Inspeção Visual
        if row['Vidro Quebrado/Rachado?'] == 'Sim':
            reciclagem += 1
            continue
        
        danos_visuais = (
            row['Backsheet Danificado?'] == 'Sim' or 
            row['Junction Box Danificado?'] == 'Sim' or 
            row['Cabos/Conectores Danificados?'] == 'Sim'
        )
        
        if danos_visuais:
            if row['Defeito Reparável?'] == 'Sim':
                manutencao += 1
            else:
                reciclagem += 1
            continue
        
        # 2. Teste de Resistência (simplificado)
        r1 = row['Resistência Medida 1 min (MΩ)']
        r2 = row['Resistência Medida 2 min (MΩ)']
        
        if pd.isna(r1) or pd.isna(r2) or r1 < 100 or r2 < 100:
            reciclagem += 1
            continue
        
        # 3. Teste de Potência
        potencia_str = str(row['Potência (% da original)'])
        if 'N/A' in potencia_str or '#DIV' in potencia_str:
            potencia = 0
        else:
            try:
                potencia = float(potencia_str)
            except:
                potencia = 0
        
        if potencia < 60:  # Como idade é conhecida, critério é < 60%
            reciclagem += 1
            continue
        
        # 4. Eletroluminescência
        if row['Rachaduras Detectadas?'] == 'Sim' or row['>50% Células Danificadas?'] == 'Sim':
            reciclagem += 1
            continue
        
        # 5. Classificação Final
        if potencia >= 90:
            aprovados_a += 1
        else:
            aprovados_b += 1
    
    total = aprovados_a + aprovados_b + reciclagem + manutencao
    
    print(f"   Classe A (≥90% potência): {aprovados_a} módulos ({aprovados_a/total*100:.1f}%)")
    print(f"   Classe B (aprovados): {aprovados_b} módulos ({aprovados_b/total*100:.1f}%)")
    print(f"   Reciclagem: {reciclagem} módulos ({reciclagem/total*100:.1f}%)")
    print(f"   Manutenção: {manutencao} módulos ({manutencao/total*100:.1f}%)")
    print(f"   Taxa de reuso: {(aprovados_a + aprovados_b)/total*100:.1f}%")
    
    print(f"\n" + "=" * 70)
    print("🎯 Use este arquivo para testar sua aplicação!")
    print("   Faça upload em: http://localhost:5000/upload")
    print("=" * 70)
    
    # Verificar se o arquivo foi criado
    if os.path.exists(output_file):
        tamanho_kb = os.path.getsize(output_file) / 1024
        print(f"\n📂 Tamanho do arquivo: {tamanho_kb:.1f} KB")
        print(f"📍 Localização: {os.path.abspath(output_file)}")
        print("✅ Arquivo criado com sucesso! Pronto para usar.")
    else:
        print(f"\n❌ Erro: Arquivo '{output_file}' não foi encontrado.")
    
    return df, data  # Retornar os dados para uso externo

# CÓDIGO ADICIONAL: Teste de cálculo de área
print("\n" + "=" * 70)
print("🔍 TESTE DE CÁLCULO DE ÁREA (para debug):")
print("=" * 70)

# Cálculo manual para verificar
altura = 1.105
largura = 0.610
area = altura * largura

print(f"Altura: {altura} m")
print(f"Largura: {largura} m")
print(f"Área calculada: {area:.4f} m²")
print(f"Área > 0? {area > 0}")
print(f"Área >= 0.01? {area >= 0.01}")
print(f"Área >= 0.1? {area >= 0.1}")

# Teste da lógica corrigida
if area < 0.01:
    print("❌ Área seria considerada inválida (muito pequena)")
elif area < 0.1:
    print("⚠️  Área seria considerada suspeita (pequena)")
else:
    print("✅ Área seria considerada válida")

# Executar a função
if __name__ == "__main__":
    try:
        print("🚀 Gerando arquivo Excel de teste compatível...")
        df, data_dict = gerar_arquivo_teste_compativel()
        
        print("\n" + "=" * 70)
        print("📏 ÁREAS DOS MÓDULOS (para debug):")
        print("-" * 40)
        # Calcular áreas para diferentes módulos
        areas = []
        for i in range(min(5, len(data_dict['Altura (m)']))):
            altura_i = data_dict['Altura (m)'][i]
            largura_i = data_dict['Largura (m)'][i]
            area_i = altura_i * largura_i
            areas.append(area_i)
            print(f"Módulo {i+1}: {altura_i:.3f}m × {largura_i:.3f}m = {area_i:.4f}m²")
        
        if areas:
            print(f"\nÁrea média (primeiros 5): {np.mean(areas):.4f} m²")
            print(f"Área mínima: {np.min(areas):.4f} m²")
            print(f"Área máxima: {np.max(areas):.4f} m²")
        
    except ImportError as e:
        print(f"\n❌ Erro de importação: {e}")
        print("\n📦 Instale as dependências necessárias:")
        print("   pip install pandas numpy openpyxl")
    except Exception as e:
        print(f"\n❌ Erro inesperado: {e}")
        import traceback
        traceback.print_exc()