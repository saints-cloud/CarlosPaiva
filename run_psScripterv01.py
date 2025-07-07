# -*- coding: utf-8 -*-
import os
import sys
import time
import traceback
import clr
import io

# Configuração crítica para execução headless
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# -------------------------------------------------------------
# Configuração de caminhos - ESSENCIAL PARA HEADLESS
# -------------------------------------------------------------
dwsim_path = r"C:\Program Files\DWSIM"
sys.path.append(dwsim_path)
os.environ["PATH"] = dwsim_path + ";" + os.environ["PATH"]
flowsheet_path = r"C:\Users\Usuario\Documents\DWSIM-Paiva\psScripter\psScripterv01.dwxmz"

# -------------------------------------------------------------
# Carregar DLLs com verificação explícita
# -------------------------------------------------------------
dlls = [
    "CapeOpen.dll",
    "DWSIM.Automation.dll",
    "DWSIM.Interfaces.dll",
    "DWSIM.GlobalSettings.dll",
    "DWSIM.SharedClasses.dll",
    "DWSIM.Thermodynamics.dll",
    "DWSIM.UnitOperations.dll",
    "DWSIM.Inspector.dll",
    "System.Buffers.dll",
]

for dll in dlls:
    try:
        clr.AddReference(os.path.join(dwsim_path, dll))
        print(f"✅ DLL carregada: {dll}")
    except Exception as e:
        print(f"⚠️ Falha ao carregar {dll}: {str(e)}")
        # Tenta carregar sem caminho completo (pode funcionar no headless)
        try:
            clr.AddReference(dll)
            print(f"✅ DLL carregada (método alternativo): {dll}")
        except:
            print(f"❌ Falha crítica ao carregar {dll}")

# Importação ESSENCIAL para execução de scripts
from DWSIM.Automation import Automation3
from DWSIM.Interfaces.Enums import Enums

# -------------------------------------------------------------
# Função CRÍTICA para inicialização de blocos Python
# -------------------------------------------------------------
def enable_scripting_in_flowsheet(flowsheet):
    """Habilita explicitamente a execução de scripts nos blocos Python"""
    for obj in flowsheet.SimulationObjects.Values:
        if obj.GraphicObject.ObjectType == Enums.GraphicObjects.ObjectType.PythonScript:
            print(f"🔧 Habilitando script em: {obj.Name}")
            obj.Enabled = True  # Garante que o bloco está habilitado
            
            # Configuração especial para execução headless
            obj.AutomationMode = True  # Modo automação
            obj.ScriptingInstance = None  # Força recriação do contexto

# -------------------------------------------------------------
# Funções existentes modificadas
# -------------------------------------------------------------
def inicializar_flowsheet(caminho):
    gerenciador = Automation3()
    
    # CONFIGURAÇÃO CRÍTICA PARA HEADLESS
    gerenciador.InitializeScriptEnvironment = True  # Habilita ambiente de script
    gerenciador.ScriptPaths = []  # Limpa caminhos de script
    
    print(f"⏳ Carregando flowsheet: {caminho}")
    fs = gerenciador.LoadFlowsheet(caminho)
    
    # Habilita explicitamente os blocos Python
    enable_scripting_in_flowsheet(fs)
    
    return gerenciador, fs

def executar_auto(gerenciador, flowsheet, sufixo):
    """
    Execução automática incremental:
      - Usa CalculateFlowsheet3 para recalcular apenas blocos marcados como "dirty"
        (componentes que tiveram alterações em parâmetros, fluxos de entrada ou resultados dependentes).
      - Cada objeto do flowsheet possui flag boolean `Calculated` indicando se seus resultados estão válidos.
      - Mantém essas flags; não limpa resultados anteriores, de modo que apenas objetos com
        `Calculated=False` ou modificados são recalculados.
      - Respeita dependências internas e timeout de execução.

    Parâmetros:
      gerenciador: instância de Automation3
      flowsheet: objeto IFlowsheet carregado
      sufixo: string usada para diferenciar arquivos de saída e log

    Para manipular flags de cálculo:
      - flowsheet.ClearAllCalculatedFlags(): limpa todas as flags, marcando tudo como "dirty".
      - obj = flowsheet.GetFlowsheetSimulationObject(name); obj.Calculated = False: limpa flag de um objeto.
    """
    timeout_seconds = 300  # tempo máximo de cálculo antes de abortar
    print("⚙️ Iniciando cálculo automático (CalculateFlowsheet3)...")
    inicio = time.perf_counter()
    excecoes = gerenciador.CalculateFlowsheet3(flowsheet, timeout_seconds)
    duracao = time.perf_counter() - inicio

    _relatorio_excecoes(excecoes)
    print(f"⏱️ Duração cálculo automático: {duracao:.2f}s")

    _salvar_saida(gerenciador, flowsheet, sufixo, duracao)

def executar_ordenado(gerenciador, flowsheet, sufixo):
    """
    Execução incremental customizada:
      - Usa RequestCalculation3 para aplicar a ordem definida em <CalculationOrderList> no XML.
      - Respeita a CustomCalculationOrder gravada no flowsheet.
      - Cada objeto do flowsheet possui flag `Calculated`; métodos incrementais
        recalcultam apenas objetos com `Calculated=False` ou modificados.
      - Mantém flags existentes; não limpa resultados anteriores.

    Parâmetros:
      gerenciador: instância de Automation3 (para salvar saída)
      flowsheet: objeto IFlowsheet carregado
      sufixo: string usada para diferenciar arquivos de saída e log

    Para manipular flags de cálculo:
      - flowsheet.ClearAllCalculatedFlags(): limpa todas as flags, forçando recálculo completo.
      - obj.Calculated = False: limpa flag de um objeto específico.

    Para modificar a ordem de cálculo:
      - Edite o XML <CalculationOrderList> diretamente.
      - Ou use flowsheet.RequestCalculation3(None, True) com lista de GUIDs via API.
    """
    print("⚙️ Iniciando cálculo ordenado (RequestCalculation3)...")
    inicio = time.perf_counter()
    try:
        # sender=None, ChangeCalculationOrder=True
        flowsheet.RequestCalculation3(None, True)
        excecoes = None
    except Exception as e:
        excecoes = [e]
    duracao = time.perf_counter() - inicio

    _relatorio_excecoes(excecoes)
    print(f"⏱️ Duração cálculo ordenado: {duracao:.2f}s")

    _salvar_saida(gerenciador, flowsheet, sufixo, duracao)

def _relatorio_excecoes(excecoes):
    """
    Prints exception messages if any occurred during flowsheet calculation; otherwise, reports no exceptions.

    Args:
        excecoes (list or None): List of exception objects or None if no exceptions occurred.
    """
    if excecoes:
        for ex in excecoes:
            msg = ex.ToString() if hasattr(ex, 'ToString') else str(ex)
            print(f"⚠️ Exceção: {msg}")
    else:
        print("🚀 Sem exceções durante o cálculo.")

def _salvar_saida(gerenciador, flowsheet, sufixo, duracao):
    """
    Saves the calculated flowsheet and a log file with execution details.

    Args:
        gerenciador (Automation3): Instance of the DWSIM automation manager.
        flowsheet (IFlowsheet): The flowsheet object to be saved.
        sufixo (str): Suffix to differentiate output and log files.
        duracao (float): Duration of the calculation in seconds.
    """
    base, ext = os.path.splitext(flowsheet_path)
    arquivo_saida = f"{base}_{sufixo}{ext}"
    arquivo_log = f"{base}_{sufixo}_log.txt"

    print(f"💾 Salvando flowsheet: {arquivo_saida}")
    gerenciador.SaveFlowsheet(flowsheet, arquivo_saida, compressed=False)

    with open(arquivo_log, 'w', encoding='utf-8') as log:
        log.write(f"Entrada: {flowsheet_path}\n")
        log.write(f"Saída:   {arquivo_saida}\n")
        log.write(f"Duração: {duracao:.2f}s\n")

    print(f"📂 Flowsheet gerado: {arquivo_saida}")
    print(f"📄 Log gerado: {arquivo_log}\n")


# ... (mantenha as outras funções como executar_auto, executar_ordenado, etc) ...

def main():
    # Execução automática incremental
    mgr_auto, fs_auto = inicializar_flowsheet(flowsheet_path)
    
    # Força cálculo inicial para estabilizar
    print("⚙️ Executando cálculo inicial de estabilização...")
    fs_auto.RequestCalculation2(True, True)  # Força recálculo completo
    
    executar_auto(mgr_auto, fs_auto, sufixo='auto')

    # Execução ordenada baseada no XML
    mgr_ord, fs_ord = inicializar_flowsheet(flowsheet_path)
    
    # Força cálculo inicial para estabilização
    print("⚙️ Executando cálculo inicial de estabilização...")
    fs_ord.RequestCalculation2(True, True)
    
    executar_ordenado(mgr_ord, fs_ord, sufixo='ordenado')

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print("❌ Erro inesperado:")
        traceback.print_exc()
        
        # Log detalhado para diagnóstico
        with open("error_dump.txt", "w", encoding="utf-8") as f:
            f.write(f"Erro: {str(e)}\n\n")
            f.write(traceback.format_exc())
        
        sys.exit(1)