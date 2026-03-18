# gerar_excel_teste.py - VERSÃO COMPLETA com dados conforme artigo
import pandas as pd
import numpy as np
from datetime import datetime
import os
import sys

def gerar_excel_testes_conforme_artigo(num_modules=20, incluir_exemplos=True):
    """Gera arquivo Excel com dados de teste conforme artigo científico"""
    
    np.random.seed(42)  # Para reprodutibilidade
    
    print("=" * 80)
    print("🚀 GERANDO ARQUIVO EXCEL DE TESTE - CONFORME ARTIGO CIENTÍFICO")
    print("=" * 80)
    
    # Funções auxiliares
    def random_sim_nao_na():
        return np.random.choice(["Sim", "Não", "NA"], p=[0.15, 0.15, 0.70])
    
    def random_sim_nao():
        return np.random.choice(["Sim", "Não"], p=[0.2, 0.8])
    
    def random_fabricante():
        fabricantes = ["Trina Solar", "Jinko Solar", "Canadian Solar", "Longi", "JA Solar", "Risen", "SunPower"]
        return np.random.choice(fabricantes)
    
    def random_modelo(fabricante):
        modelos = {
            "Trina Solar": ["Vertex S+", "Vertex N", "DEG15"],
            "Jinko Solar": ["Tiger Neo", "Cheetah HC", "Swan"],
            "Canadian Solar": ["HiKu7", "BiHiKu", "HiDM"],
            "Longi": ["Hi-MO 7", "Hi-MO 6", "Hi-MO 5"],
            "JA Solar": ["DeepBlue 4.0", "DeepBlue 3.0"],
            "Risen": ["Titan", "Hyper-ion"],
            "SunPower": ["Maxeon 6", "Performance"]
        }
        return np.random.choice(modelos.get(fabricante, ["Standard"]))
    
    # Dados conforme artigo (22 anos de operação)
    ano_base = 2000  # Artigo: módulos instalados em 2000
    
    # Criar lista de módulos
    modules_data = []
    
    for i in range(1, num_modules + 1):
        fabricante = random_fabricante()
        modelo = random_modelo(fabricante)
        
        # Ano de fabricação: entre 1998-2003 (conforme artigo)
        ano = np.random.choice([1998, 1999, 2000, 2001, 2002, 2003])
        idade = datetime.now().year - ano
        
        # Tipo de módulo
        bifacial = np.random.choice(["Bifacial", "Monofacial"], p=[0.3, 0.7])
        
        # Dimensões típicas (conforme artigo: 1.110m x 0.610m)
        altura = round(np.random.uniform(1.0, 1.2), 3)
        largura = round(np.random.uniform(0.5, 0.7), 3)
        area = altura * largura
        
        # Dados de inspeção visual
        # Artigo: 6 módulos com problemas visuais (de 76)
        if i <= int(num_modules * 0.08):  # ~8% com problemas
            vidro_quebrado = "Sim" if i == 1 else "Não"
            backsheet_danificado = "Sim" if i in [2, 3] else "Não"
            junction_danificado = "Sim" if i == 4 else "Não"
            cabos_danificados = "Sim" if i in [5, 6] else "Não"
        else:
            vidro_quebrado = "Não"
            backsheet_danificado = "Não"
            junction_danificado = "Não"
            cabos_danificados = "Não"
        
        # Defeito reparável
        if any(x == "Sim" for x in [vidro_quebrado, backsheet_danificado, junction_danificado, cabos_danificados]):
            defeito_reparavel = "Sim" if np.random.random() < 0.5 else "Não"
        else:
            defeito_reparavel = "NA"
        
        # Dados técnicos originais
        # Artigo usou módulos de 56W e 77W
        potencia_datasheet = np.random.choice([56, 77, 100, 150, 200, 250, 300, 350])
        voc_original = round(np.random.uniform(30, 50), 1)
        isc_original = round(np.random.uniform(8, 12), 2)
        
        # Fill Factor Original
        if voc_original > 0 and isc_original > 0:
            ff_original_calc = (potencia_datasheet / (voc_original * isc_original))
        else:
            ff_original_calc = round(np.random.uniform(0.70, 0.80), 4)
        
        # Resistência conforme artigo
        # Artigo: média ~131 MΩ·m², mínimo 40 MΩ·m²
        r_base = np.random.uniform(80, 250)
        r1 = round(r_base, 1)
        r2 = round(r1 * np.random.uniform(0.95, 1.05), 1)
        resistencia_fabricante_calc = min(r1, r2) * area
        
        # Artigo: 12 módulos falharam no teste de resistência
        if i <= int(num_modules * 0.16):  # ~16% falham (artigo: 12/76 = 15.8%)
            resistencia_fabricante_calc = np.random.uniform(10, 39)  # Abaixo de 40
        
        # Potência percentual conforme artigo
        # Artigo: 68% reuso, média de potência ~70-80%
        if np.random.random() < 0.68:  # 68% chance de ser reaproveitável
            # Módulos reaproveitáveis
            if np.random.random() < 0.5:  # 50% Classe A
                potencia_percentual = np.random.uniform(78, 95)  # ≥ esperada
            else:  # 50% Classe B
                potencia_percentual = np.random.uniform(68, 77.9)  # ≥ mínimo, < esperada
        else:
            # Módulos para reciclagem
            potencia_percentual = np.random.uniform(10, 67.9)
        
        # Dados medidos baseados na potência percentual
        pmax_medido = round(potencia_datasheet * (potencia_percentual / 100), 1)
        voc_medido = round(voc_original * np.random.uniform(0.95, 1.05), 1)
        isc_medido = round(isc_original * np.random.uniform(0.95, 1.05), 2)
        fill_factor_medido = round(ff_original_calc * 100 * (potencia_percentual/100) * np.random.uniform(0.9, 1.1), 1)
        
        # Idade conhecida? (conforme artigo, sim)
        idade_conhecida = "Sim"
        degradacao_anual = "1.00%"  # Conforme artigo: 1% anual
        
        # Eletroluminescência (conforme artigo: 9 módulos com problemas)
        eletroluminescencia = "Sim" if np.random.random() < 0.7 else "Não"
        if eletroluminescencia == "Sim":
            # Artigo: 9 módulos com problemas (de 76)
            if i <= int(num_modules * 0.12):  # ~12% com problemas
                rachaduras = "Sim" if np.random.random() < 0.7 else "Não"
                celulas_danificadas = "Sim" if np.random.random() < 0.3 else "Não"
            else:
                rachaduras = "Não"
                celulas_danificadas = "Não"
        else:
            rachaduras = "Não"
            celulas_danificadas = "Não"
        
        # Criar dicionário com dados
        module = {
            "ID do Módulo": f"M{i:03d}",
            "NS do Módulo": f"ART{i:03d}{ano}",
            "Fabricante": fabricante,
            "Modelo": modelo,
            "Potência do datasheet (W)": potencia_datasheet,
            "Voc Original (V)": voc_original,
            "Isc Original (A)": isc_original,
            "Ano": ano,
            "Bifacial/Monofacial": bifacial,
            "Vidro Quebrado/Rachado?": vidro_quebrado,
            "Backsheet Danificado?": backsheet_danificado,
            "Junction Box Danificado?": junction_danificado,
            "Cabos/Conectores Danificados?": cabos_danificados,
            "Defeito Reparável?": defeito_reparavel,
            "Altura (m)": altura,
            "Largura (m)": largura,
            "Resistência Medida 1 min (MΩ)": r1,
            "Resistência Medida 2 min (MΩ)": r2,
            "Resistência Ôhmica Fabricante (MΩ·m²)": round(resistencia_fabricante_calc, 2),
            "Idade do Módulo Conhecida?": idade_conhecida,
            "Degradação Anual Esperada (%)": degradacao_anual,
            "Voc Medido (V)": voc_medido,
            "Isc Medido (A)": isc_medido,
            "Pmáx Medido (W)": pmax_medido,
            "Fill Factor Medido (%)": fill_factor_medido,
            "Potência (% da original)": round(potencia_percentual/100, 4),
            "Fill Factor Original (%)": round(ff_original_calc, 4),
            "Foi realizado Eletroluminescência?": eletroluminescencia,
            "Rachaduras Detectadas?": rachaduras,
            ">50% Células Danificadas?": celulas_danificadas
        }
        
        modules_data.append(module)
    
    # Criar DataFrame
    df = pd.DataFrame(modules_data)
    
    # Ordem das colunas conforme template
    columns_order = [
        "ID do Módulo", "NS do Módulo", "Fabricante", "Modelo", 
        "Potência do datasheet (W)", "Voc Original (V)", "Isc Original (A)",
        "Ano", "Bifacial/Monofacial",
        "Vidro Quebrado/Rachado?", "Backsheet Danificado?", 
        "Junction Box Danificado?", "Cabos/Conectores Danificados?",
        "Defeito Reparável?", 
        "Altura (m)", "Largura (m)",
        "Resistência Medida 1 min (MΩ)", "Resistência Medida 2 min (MΩ)", "Resistência Ôhmica Fabricante (MΩ·m²)",
        "Idade do Módulo Conhecida?", "Degradação Anual Esperada (%)",
        "Voc Medido (V)", "Isc Medido (A)", "Pmáx Medido (W)", 
        "Fill Factor Medido (%)", "Potência (% da original)",
        "Fill Factor Original (%)",
        "Foi realizado Eletroluminescência?", "Rachaduras Detectadas?", ">50% Células Danificadas?"
    ]
    
    df = df[columns_order]
    
    # Nome do arquivo
    output_file = f"Planilha_Teste_Artigo_{num_modules}_Modulos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    # Salvar em Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Módulos', index=False)
        
        # Adicionar aba de instruções
        instructions = pd.DataFrame({
            'INSTRUÇÕES': [
                'PLANILHA DE TESTE - CONFORME ARTIGO CIENTÍFICO',
                '',
                'Referência: "Circular solar economy: PV modules decision-making framework for reuse"',
                'Journal of Cleaner Production, 2023',
                '',
                'CARACTERÍSTICAS DOS DADOS:',
                f'• Total de módulos: {num_modules}',
                f'• Idade média: ~{idade} anos (similar ao artigo)',
                '• Potência esperada para idade: ~78% (22 anos × 1%/ano)',
                '• Mínimo aceitável: ~68% (esperada - 10%)',
                '• Distribuição conforme artigo: ~68% reaproveitáveis',
                '',
                'CRITÉRIOS DO ARTIGO:',
                '1. Inspeção Visual: Vidro quebrado = reciclagem imediata',
                '2. Resistência Isolamento: Mínimo 40 MΩ·m²',
                '3. Potência:',
                '   • Idade conhecida: ≥ (Esperada - 10%)',
                '   • Idade desconhecida: ≥ 60% original',
                '4. Classificação: Classe A (≥ esperada), Classe B (≥ mínima)',
                '',
                'PARA USAR:',
                '1. Execute: python app.py',
                '2. Acesse: http://localhost:5000',
                '3. Faça upload deste arquivo',
                '4. Visualize os resultados com gráficos',
                '',
                'RESULTADO ESPERADO: ~68% taxa de reuso (conforme artigo)'
            ]
        })
        
        instructions.to_excel(writer, sheet_name='Instruções', index=False)
    
    print(f"✅ Arquivo gerado: {output_file}")
    print(f"📊 Total de módulos: {num_modules}")
    print(f"📈 Idade média: {idade} anos")
    
    # Calcular estatísticas esperadas
    potencia_media = df['Potência (% da original)'].mean() * 100
    resistencia_media = df['Resistência Ôhmica Fabricante (MΩ·m²)'].mean()
    resistencia_abaixo_40 = (df['Resistência Ôhmica Fabricante (MΩ·m²)'] < 40).sum()
    
    print(f"\n📋 ESTATÍSTICAS DOS DADOS:")
    print(f"   • Potência média: {potencia_media:.1f}%")
    print(f"   • Resistência média: {resistencia_media:.1f} MΩ·m²")
    print(f"   • Módulos com resistência < 40 MΩ·m²: {resistencia_abaixo_40}")
    print(f"   • Módulos com vidro quebrado: {(df['Vidro Quebrado/Rachado?'] == 'Sim').sum()}")
    print(f"   • Módulos com outros danos: {sum(1 for _, row in df.iterrows() if row['Backsheet Danificado?'] == 'Sim' or row['Junction Box Danificado?'] == 'Sim' or row['Cabos/Conectores Danificados?'] == 'Sim')}")
    
    print(f"\n🎯 RESULTADO ESPERADO (simulação):")
    print(f"   • Artigo reporta: 68% reuso para módulos de 22 anos")
    print(f"   • Nosso dataset: projetado para ~68% reuso")
    
    print(f"\n📁 INFORMAÇÕES DO ARQUIVO:")
    print(f"   • Localização: {os.path.abspath(output_file)}")
    print(f"   • Tamanho: {os.path.getsize(output_file) / 1024:.1f} KB")
    
    print(f"\n🚀 PRÓXIMOS PASSOS:")
    print(f"   1. Execute o servidor: python app.py")
    print(f"   2. Acesse: http://localhost:5000")
    print(f"   3. Faça upload do arquivo '{output_file}'")
    print(f"   4. Verifique se a taxa de reuso é ~68%")
    
    print(f"\n" + "=" * 80)
    print(f"✅ PLANILHA DE TESTE PRONTA PARA USO!")
    print("=" * 80)
    
    return output_file

# Função auxiliar
def converter_numero(valor):
    try:
        if pd.isna(valor):
            return 0.0
        if isinstance(valor, str):
            valor = valor.replace(',', '.').replace('%', '')
        return float(valor)
    except:
        return 0.0

# Executar a função
if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Gerar planilha de teste conforme artigo científico')
    parser.add_argument('-n', '--num-modules', type=int, default=20, help='Número de módulos (padrão: 20)')
    parser.add_argument('-o', '--output', help='Nome do arquivo de saída')
    
    args = parser.parse_args()
    
    try:
        print("🚀 Iniciando geração de arquivo de teste...")
        
        output_file = gerar_excel_testes_conforme_artigo(
            num_modules=args.num_modules
        )
        
        if args.output:
            # Renomear arquivo se especificado
            new_name = args.output
            if not new_name.endswith('.xlsx'):
                new_name += '.xlsx'
            os.rename(output_file, new_name)
            output_file = new_name
        
        print(f"\n🎉 Arquivo criado com sucesso: {output_file}")
        print("📤 Pronto para upload no sistema!")
        
    except ImportError as e:
        print(f"\n❌ Erro de importação: {e}")
        print("\n📦 Instale as dependências:")
        print("   pip install pandas numpy openpyxl")
    except Exception as e:
        print(f"\n❌ Erro inesperado: {e}")
        print(f"\n🔧 Detalhes: {sys.exc_info()}")