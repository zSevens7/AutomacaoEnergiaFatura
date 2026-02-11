import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import time
import glob
import shutil
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import threading

# =============================================================================
# 1. O CÉREBRO DO ROBÔ (SELENIUM) - VERSÃO FINAL OTIMIZADA
# =============================================================================
class EquatorialBot:
    def __init__(self, download_folder):
        self.driver = None
        self.wait = None
        self.download_folder = os.path.abspath(download_folder)
        
        if not os.path.exists(self.download_folder):
            os.makedirs(self.download_folder)

    def abrir_navegador(self):
        if self.driver is not None: 
            return 

        options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": self.download_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "profile.default_content_setting_values.automatic_downloads": 1,
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", prefs)
        options.add_argument("--start-maximized")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        
        # Remover assinaturas de automação
        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument('--disable-blink-features=AutomationControlled')
        
        # Adicionar headers para parecer mais humano
        options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 30)
        self.driver.get("https://ma.equatorialenergia.com.br/") 
        
        print(f"Navegador aberto. Pasta de download: {self.download_folder}")

    def verificar_e_trocar_uc(self, uc_alvo):
        """Verifica e troca a UC se necessário"""
        try:
            select_element = self.wait.until(
                EC.presence_of_element_located((By.ID, "conta_contrato"))
            )
            
            select = Select(select_element)
            uc_atual = select.first_selected_option.text.strip()
            uc_alvo_limpa = str(uc_alvo).strip()
            
            print(f"UC Atual: {uc_atual}")
            print(f"UC Alvo: {uc_alvo_limpa}")
            
            uc_atual_clean = uc_atual.replace('.', '')
            uc_alvo_clean = uc_alvo_limpa.replace('.', '')
            
            if uc_alvo_clean in uc_atual_clean:
                print("UC já está selecionada corretamente")
                return True
            
            print(f"Trocando UC: {uc_atual} -> {uc_alvo_limpa}")
            
            try:
                select.select_by_visible_text(uc_alvo_limpa)
            except:
                try:
                    for option in select.options:
                        option_text = option.text.strip().replace('.', '')
                        option_value = option.get_attribute('value')
                        
                        if uc_alvo_clean in option_text or uc_alvo_clean in option_value:
                            option.click()
                            print(f"UC selecionada via opção: {option.text}")
                            break
                except Exception as e:
                    print(f"Erro ao selecionar opção: {e}")
                    return f"Erro: UC {uc_alvo} não encontrada nas opções"
            
            # Aguarda mais tempo para a página atualizar
            time.sleep(8)
            
            try:
                select_element = self.driver.find_element(By.ID, "conta_contrato")
                select = Select(select_element)
                nova_uc = select.first_selected_option.text.strip().replace('.', '')
                
                if uc_alvo_clean in nova_uc:
                    print(f"UC trocada com sucesso para: {nova_uc}")
                    return True
                else:
                    return f"Erro: Não foi possível trocar para UC {uc_alvo}"
                    
            except Exception as e:
                return f"Erro ao verificar UC após troca: {str(e)}"
                
        except Exception as e:
            print(f"Erro no verificar_e_trocar_uc: {e}")
            return f"Erro ao acessar seletor de UC: {str(e)}"

    def clicar_ver_fatura_direto(self):
        """Clica diretamente no botão 'Ver Fatura' no modal"""
        print("Tentando clicar diretamente no botão 'Ver Fatura'...")
        
        try:
            # Localiza o modal
            modal = self.driver.find_element(By.CLASS_NAME, "lista-debitos-modal")
            print("Modal encontrado")
            
            # Verifica se o modal está visível
            estilo = modal.get_attribute("style")
            print(f"Estilo do modal: {estilo}")
            
            if "display: none" in estilo:
                print("Modal está oculto, tentando abrir...")
                # Se o modal estiver oculto, clica na primeira fatura para abrir
                primeira_fatura = self.driver.find_element(By.CLASS_NAME, "bill-reference")
                linha_fatura = primeira_fatura.find_element(By.XPATH, "./ancestor::tr")
                linha_fatura.click()
                time.sleep(4)
            
            # Aguarda um pouco mais para o modal carregar completamente
            time.sleep(3)
            
            # Agora tenta encontrar o botão no modal
            try:
                # Procura pelo botão dentro do modal
                btn_ver_fatura = modal.find_element(By.CSS_SELECTOR, "a.download-pdf")
                print(f"Botão encontrado no modal: {btn_ver_fatura.text}")
            except:
                # Tenta por texto
                btn_ver_fatura = modal.find_element(By.XPATH, ".//a[contains(text(), 'Ver Fatura')]")
                print("Botão encontrado por texto")
            
            # Verifica se o botão está visível
            print(f"Botão visível: {btn_ver_fatura.is_displayed()}")
            print(f"Botão habilitado: {btn_ver_fatura.is_enabled()}")
            
            # Rola até o botão
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_ver_fatura)
            time.sleep(1)
            
            # Tenta diferentes métodos de clique
            print("Tentando clique com ActionChains...")
            actions = ActionChains(self.driver)
            actions.move_to_element(btn_ver_fatura).pause(0.5).click().pause(0.5).perform()
            print("✅ Clique ActionChains realizado")
            
            # Aguarda um pouco
            time.sleep(3)
            
            return True
            
        except Exception as e:
            print(f"Erro ao clicar diretamente: {e}")
            
            # Tenta método alternativo: JavaScript direto
            try:
                print("Tentando clique via JavaScript...")
                script = """
                var modal = document.querySelector('.lista-debitos-modal');
                var btn = modal.querySelector('a.download-pdf');
                if (btn) {
                    // Simula um clique humano
                    var event = new MouseEvent('click', {
                        view: window,
                        bubbles: true,
                        cancelable: true
                    });
                    btn.dispatchEvent(event);
                    return true;
                }
                return false;
                """
                resultado = self.driver.execute_script(script)
                if resultado:
                    print("✅ Clique JavaScript realizado")
                    time.sleep(3)
                    return True
            except Exception as js_e:
                print(f"JavaScript também falhou: {js_e}")
            
            return False

    def verificar_e_evitar_clara(self):
        """Verifica se está na página da assistente Clara e clica para emitir segunda via."""
        try:
            # Verifica se estamos na página da Clara por algum elemento único nela
            # Ex: Texto "Olá, tudo bem?", nome "Clara", ou o título da página.
            # O 'timeout' baixo (5 seg) é intencional, só para checar rápido.
            elemento_clara = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Clara') or contains(text(), 'Olá, tudo bem?')]"))
            )
            print("⚠️  Detectada página da assistente virtual 'Clara'. Redirecionando...")

            # Procura e clica no link específico para "Emitir segunda via"
            link_segunda_via = self.driver.find_element(By.XPATH, "//a[contains(@href, 'emitir-segunda-via')]")
            self.driver.execute_script("arguments[0].click();", link_segunda_via)
            print("✅ Clicou em 'Emitir segunda via'. Aguardando carregamento...")

            # Aguarda um tempo para a nova página carregar ou algum elemento dela aparecer
            time.sleep(5)
            return True

        except TimeoutException:
            # Se os elementos da Clara não forem encontrados em 5 segundos,
            # assumimos que NÃO estamos naquela página e o fluxo continua normal.
            return False
        except Exception as e:
            print(f"⚠️  Erro ao tentar sair da página da Clara: {e}")
            return False

    def baixar_ultima_fatura(self, uc_cliente):
        """Baixa a última fatura disponível (a mais recente)"""
        if not self.driver:
            return "Erro: Navegador não inicializado"
            
        try:
            print(f"\n{'='*50}")
            print(f"PROCESSANDO: UC {uc_cliente}")
            print(f"{'='*50}")

            # --- NOVA ETAPA ADICIONADA AQUI ---
            print("Verificando se foi redirecionado para a página da Clara...")
            self.verificar_e_evitar_clara()
            # ----------------------------------
            
            # 1. Verifica e troca UC se necessário
            print("Verificando UC atual...")
            resultado_uc = self.verificar_e_trocar_uc(uc_cliente)
            
            if resultado_uc != True:
                return resultado_uc
            
            # 2. Limpa a pasta de downloads temporários
            self.limpar_downloads_temporarios()
            
            # 3. Registra arquivos existentes antes do download
            arquivos_antes = set(glob.glob(os.path.join(self.download_folder, "*.pdf")))
            print(f"Arquivos PDF antes do download: {len(arquivos_antes)}")
            
            # 4. Localiza a PRIMEIRA fatura (a mais recente)
            print("Procurando a última fatura (mais recente)...")
            
            # Aguarda a tabela de faturas carregar
            try:
                self.wait.until(
                    EC.presence_of_element_located((By.CLASS_NAME, "bill-reference"))
                )
                print("Tabela de faturas carregada")
            except TimeoutException:
                return "Erro: Tabela de faturas não encontrada"
            
            # Procura a PRIMEIRA fatura da lista (a mais recente)
            try:
                todas_faturas = self.driver.find_elements(By.CLASS_NAME, "bill-reference")
                
                if not todas_faturas:
                    return "Erro: Nenhuma fatura encontrada"
                
                primeira_fatura = todas_faturas[0]
                
                # Encontra o mês de referência
                try:
                    elemento_mes = primeira_fatura.find_element(By.CLASS_NAME, "referencia_legada")
                    mes_referencia = elemento_mes.text.strip()
                    print(f"Última fatura encontrada: {mes_referencia}")
                except:
                    mes_referencia = "Ultima_Fatura"
                    print("Mês não encontrado, usando nome padrão")
                
                # Encontra a linha da tabela
                linha_fatura = primeira_fatura.find_element(By.XPATH, "./ancestor::tr")
                
                # 5. Clica no valor da fatura para abrir o modal
                try:
                    celula_valor = linha_fatura.find_element(By.CLASS_NAME, "bill-value")
                    print("Clicando no valor da fatura para abrir modal...")
                    
                    # Clique suave com JavaScript
                    self.driver.execute_script("arguments[0].click();", celula_valor)
                    print("✅ Clique no valor realizado")
                    
                    # Aguarda o modal carregar completamente
                    time.sleep(5)
                    
                except NoSuchElementException:
                    print("Célula de valor não encontrada, tentando clicar na linha...")
                    linha_fatura.click()
                    time.sleep(5)
                
                # 6. Tenta clicar no botão "Ver Fatura" de várias formas
                print("Tentando clicar em 'Ver Fatura'...")
                
                # Método 1: Clique direto no botão
                clique_sucesso = self.clicar_ver_fatura_direto()
                
                if not clique_sucesso:
                    # Método 2: Procura alternativas
                    print("Tentando método alternativo...")
                    try:
                        # Procura por qualquer link de download
                        links = self.driver.find_elements(By.TAG_NAME, "a")
                        for link in links:
                            href = link.get_attribute("href") or ""
                            onclick = link.get_attribute("onclick") or ""
                            texto = link.text or ""
                            
                            if "ver fatura" in texto.lower() or ".pdf" in href.lower() or "download" in onclick.lower():
                                print(f"Encontrado link alternativo: {texto}")
                                self.driver.execute_script("arguments[0].scrollIntoView(true);", link)
                                time.sleep(1)
                                link.click()
                                clique_sucesso = True
                                break
                    except Exception as e:
                        print(f"Erro no método alternativo: {e}")
                
                if not clique_sucesso:
                    return "Erro: Não foi possível acionar o download"
                
                print("Download iniciado. Aguardando...")
                
                # 7. Aguarda o download completar com timeout maior
                novo_arquivo = self.esperar_download_completar(arquivos_antes, timeout=90)
                
                if novo_arquivo:
                    # Define o nome do arquivo final
                    nome_final = os.path.join(
                        self.download_folder, 
                        f"Fatura_{uc_cliente}_{mes_referencia.replace('/', '-')}.pdf"
                    )
                    
                    # Evita sobrescrever arquivos existentes
                    contador = 1
                    nome_base, extensao = os.path.splitext(nome_final)
                    while os.path.exists(nome_final):
                        nome_final = f"{nome_base}_{contador}{extensao}"
                        contador += 1
                    
                    # Move/Renomeia o arquivo
                    try:
                        shutil.move(novo_arquivo, nome_final)
                        print(f"✅ Download realizado: {os.path.basename(nome_final)}")
                        print(f"📍 Salvo em: {nome_final}")
                        return f"Sucesso: {mes_referencia}"
                    except Exception as e:
                        print(f"Erro ao renomear arquivo: {e}")
                        # Se não conseguir mover, verifica se o arquivo já está com nome correto
                        return f"Download realizado mas não renomeado: {novo_arquivo}"
                else:
                    # Verifica se o arquivo já foi baixado anteriormente
                    nome_potencial = os.path.join(
                        self.download_folder, 
                        f"Fatura_{uc_cliente}_{mes_referencia.replace('/', '-')}.pdf"
                    )
                    
                    if os.path.exists(nome_potencial):
                        print(f"Arquivo já existe: {os.path.basename(nome_potencial)}")
                        return f"Sucesso: {mes_referencia} (já existia)"
                    else:
                        return "Erro: Download não finalizado ou arquivo não encontrado"
                    
            except Exception as e:
                print(f"Erro ao processar primeira fatura: {e}")
                return f"Erro: Não foi possível acessar a última fatura - {str(e)}"
                
        except Exception as e:
            print(f"Erro crítico durante o download: {e}")
            return f"Falha: {str(e)}"

    def limpar_downloads_temporarios(self):
        """Limpa arquivos temporários de downloads anteriores"""
        try:
            for temp_file in glob.glob(os.path.join(self.download_folder, "*.crdownload")):
                try:
                    os.remove(temp_file)
                    print(f"Removido arquivo temporário: {os.path.basename(temp_file)}")
                except:
                    pass
            
            for temp_file in glob.glob(os.path.join(self.download_folder, "*.part")):
                try:
                    os.remove(temp_file)
                    print(f"Removido arquivo temporário: {os.path.basename(temp_file)}")
                except:
                    pass
                    
        except Exception as e:
            print(f"Erro ao limpar downloads temporários: {e}")

    def esperar_download_completar(self, arquivos_antes, timeout=90):
        """Aguarda o download de um novo arquivo PDF com timeout maior"""
        print("Aguardando download completar...")
        
        inicio = time.time()
        ultimo_status = time.time()
        arquivos_temporarios_anteriores = 0
        
        while time.time() - inicio < timeout:
            # Lista todos os arquivos PDF atuais
            arquivos_atual = set(glob.glob(os.path.join(self.download_folder, "*.pdf")))
            
            # Lista arquivos temporários para monitoramento
            temp_files = glob.glob(os.path.join(self.download_folder, "*.crdownload"))
            
            # Se houver arquivos temporários, mostra quantos
            if temp_files:
                if len(temp_files) != arquivos_temporarios_anteriores:
                    print(f"  📥 Download em progresso... {len(temp_files)} arquivo(s) temporário(s)")
                    arquivos_temporarios_anteriores = len(temp_files)
            
            # Log a cada 10 segundos se não houver arquivos temporários
            elif time.time() - ultimo_status > 10:
                print(f"  ⏳ Aguardando... {int(time.time() - inicio)} segundos")
                ultimo_status = time.time()
            
            # Encontra novos arquivos
            novos_arquivos = arquivos_atual - arquivos_antes
            
            for arquivo in novos_arquivos:
                # Verifica se o arquivo não está sendo baixado
                if not (arquivo.endswith('.crdownload') or arquivo.endswith('.part')):
                    # Tenta abrir o arquivo para verificar se está completo
                    try:
                        tamanho = os.path.getsize(arquivo)
                        if tamanho > 1024:  # Mais que 1KB
                            print(f"✅ Download completado: {os.path.basename(arquivo)} ({tamanho} bytes)")
                            return arquivo
                    except (IOError, OSError) as e:
                        # Arquivo ainda está sendo escrito
                        continue
            
            # Se não tem mais arquivos temporários, verifica se algum PDF novo apareceu
            if not temp_files and novos_arquivos:
                time.sleep(2)
                arquivos_atual = set(glob.glob(os.path.join(self.download_folder, "*.pdf")))
                novos_arquivos = arquivos_atual - arquivos_antes
                
                for arquivo in novos_arquivos:
                    if os.path.exists(arquivo):
                        print(f"✅ Download finalizado sem arquivos temporários: {os.path.basename(arquivo)}")
                        return arquivo
            
            time.sleep(2)  # Verifica a cada 2 segundos
        
        print(f"⏰ Timeout ({timeout}s): Download não completou no tempo esperado")
        
        # Verifica se algum arquivo foi baixado mesmo com timeout
        arquivos_atual = set(glob.glob(os.path.join(self.download_folder, "*.pdf")))
        novos_arquivos = arquivos_atual - arquivos_antes
        
        if novos_arquivos:
            arquivo = list(novos_arquivos)[0]
            print(f"⚠️ Usando arquivo disponível (possivelmente incompleto): {os.path.basename(arquivo)}")
            return arquivo
        
        return None

    def fazer_logout(self):
        """Realiza logout do sistema - versão corrigida sem duplicação"""
        print("Realizando logout...")
        
        try:
            # Tenta encontrar e clicar no botão de sair
            btn_sair = self.driver.find_element(By.ID, "clear-section")
            self.driver.execute_script("arguments[0].click();", btn_sair)
            print("✅ Logout realizado")
            time.sleep(3)  # Aguarda o logout completar
            
            # Verifica se ainda está na mesma página, se sim, tenta recarregar
            if "sua-conta" in self.driver.current_url:
                print("Ainda na página de conta, recarregando...")
                self.driver.delete_all_cookies()
                self.driver.get("https://ma.equatorialenergia.com.br/")
                time.sleep(3)
                
        except Exception as e:
            print(f"⚠️ Não foi possível fazer logout automático: {e}")
            print("Tentando método alternativo...")
            
            try:
                # Método alternativo: limpar cookies e ir para página inicial
                self.driver.delete_all_cookies()
                self.driver.get("https://ma.equatorialenergia.com.br/")
                print("✅ Logout forçado realizado")
                time.sleep(3)
            except Exception as e2:
                print(f"⚠️ Erro no logout forçado: {e2}")

# =============================================================================
# 2. INTERFACE TKINTER - VERSÃO FINAL
# =============================================================================
class PainelControle:
    def __init__(self, root, excel_path):
        self.root = root
        self.root.title("🤖 Equatorial Cyborg Controller v1.0")
        self.root.geometry("500x650")
        self.root.configure(bg="#2C3E50")
        
        # Configurações
        base_dir = os.getcwd()
        self.download_path = os.path.join(base_dir, "output", "faturas")
        self.excel_path = excel_path
        
        # Inicializa Dados e Robô
        self.dados = []
        self.index_atual = 0
        self.bot = EquatorialBot(self.download_path)
        self.processo_em_andamento = False  # Flag para evitar múltiplos cliques
        
        # Variável para status
        self.status_var = tk.StringVar()
        self.status_var.set("Aguardando início...")
        
        self.carregar_excel()
        self.montar_layout()
        
        # Inicia navegador automaticamente
        self.iniciar_navegador()
        
        self.atualizar_tela()

    def carregar_excel(self):
        try:
            df = pd.read_excel(self.excel_path, dtype=str)
            self.dados = df.to_dict('records')
            print(f"✅ Excel carregado: {len(self.dados)} clientes encontrados")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler o Excel:\n{e}")

    def montar_layout(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # HEADER
        frame_header = tk.Frame(self.root, bg="#34495E", pady=10)
        frame_header.pack(fill="x")
        
        tk.Label(frame_header, text="EQUATORIAL CYBORG CONTROLLER", 
                font=("Segoe UI", 14, "bold"), fg="white", bg="#34495E").pack()
        tk.Label(frame_header, text="Sistema Automático de Download de Faturas", 
                font=("Segoe UI", 9), fg="#BDC3C7", bg="#34495E").pack()
        
        # CONTADOR
        frame_contador = tk.Frame(self.root, bg="#2C3E50", pady=5)
        frame_contador.pack()
        self.lbl_contador = tk.Label(frame_contador, 
                                    text="Cliente 0/0", 
                                    font=("Arial", 10, "bold"), 
                                    fg="#F39C12", bg="#2C3E50")
        self.lbl_contador.pack()
        
        # INFORMAÇÕES DO CLIENTE
        frame_cliente = tk.Frame(self.root, bg="#2C3E50", pady=10, padx=20)
        frame_cliente.pack(fill="x")
        tk.Label(frame_cliente, text="CLIENTE ATUAL", 
                font=("Segoe UI", 10), fg="#BDC3C7", bg="#2C3E50").pack(anchor="w")
        self.lbl_nome_cliente = tk.Label(frame_cliente, 
                                        text="---", 
                                        font=("Segoe UI", 12, "bold"), 
                                        fg="white", bg="#2C3E50")
        self.lbl_nome_cliente.pack(anchor="w")
        
        # DADOS DE ACESSO
        frame_dados = tk.Frame(self.root, bg="#2C3E50", padx=20, pady=10)
        frame_dados.pack(fill="x")
        
        # UC
        tk.Label(frame_dados, text="Conta Contrato (UC)", 
                fg="#BDC3C7", bg="#2C3E50", font=("Arial", 9)).pack(anchor="w")
        f_uc = tk.Frame(frame_dados, bg="#2C3E50")
        f_uc.pack(fill="x", pady=(2, 10))
        self.entry_uc = tk.Entry(f_uc, font=("Consolas", 12), bg="#ECF0F1")
        self.entry_uc.pack(side="left", fill="x", expand=True)
        tk.Button(f_uc, text="📋", command=lambda: self.copiar(self.entry_uc.get()),
                 bg="#3498DB", fg="white", relief="flat", width=3).pack(side="right", padx=5)
        
        # LOGIN
        tk.Label(frame_dados, text="Login / CPF", 
                fg="#BDC3C7", bg="#2C3E50", font=("Arial", 9)).pack(anchor="w")
        f_login = tk.Frame(frame_dados, bg="#2C3E50")
        f_login.pack(fill="x", pady=(2, 10))
        self.entry_login = tk.Entry(f_login, font=("Consolas", 12), bg="#ECF0F1")
        self.entry_login.pack(side="left", fill="x", expand=True)
        tk.Button(f_login, text="📋", command=lambda: self.copiar(self.entry_login.get()),
                 bg="#3498DB", fg="white", relief="flat", width=3).pack(side="right", padx=5)
        
        # SENHA
        tk.Label(frame_dados, text="Senha", 
                fg="#BDC3C7", bg="#2C3E50", font=("Arial", 9)).pack(anchor="w")
        f_senha = tk.Frame(frame_dados, bg="#2C3E50")
        f_senha.pack(fill="x", pady=(2, 0))
        self.entry_senha = tk.Entry(f_senha, font=("Consolas", 12), bg="#ECF0F1", show="•")
        self.entry_senha.pack(side="left", fill="x", expand=True)
        tk.Button(f_senha, text="📋", command=lambda: self.copiar(self.entry_senha.get()),
                 bg="#3498DB", fg="white", relief="flat", width=3).pack(side="right", padx=5)
        
        # ÁREA DE STATUS
        frame_status = tk.Frame(self.root, bg="#2C3E50", padx=20, pady=10)
        frame_status.pack(fill="x")
        self.lbl_status = tk.Label(frame_status, 
                                  textvariable=self.status_var,
                                  font=("Arial", 10), 
                                  bg="#2C3E50", fg="#F1C40F",
                                  wraplength=450, justify="left")
        self.lbl_status.pack()
        
        # BOTÕES DE AÇÃO
        frame_botoes = tk.Frame(self.root, bg="#2C3E50", padx=20, pady=10)
        frame_botoes.pack(fill="x")
        
        self.btn_baixar = tk.Button(frame_botoes, 
                                   text="🤖 BAIXAR ÚLTIMA FATURA", 
                                   font=("Segoe UI", 11, "bold"), 
                                   bg="#27AE60", fg="white",
                                   command=self.executar_robo,
                                   height=2, cursor="hand2",
                                   state="disabled")
        self.btn_baixar.pack(fill="x", pady=(0, 5))
        
        self.btn_pular = tk.Button(frame_botoes, 
                                   text="⏭️ PULAR CLIENTE (Manual)", 
                                   font=("Segoe UI", 9), 
                                   bg="#E74C3C", fg="white",
                                   command=self.pular_cliente)
        self.btn_pular.pack(fill="x", pady=(0, 5))
        
        # NAVEGAÇÃO
        frame_nav = tk.Frame(frame_botoes, bg="#2C3E50")
        frame_nav.pack(fill="x", pady=5)
        tk.Button(frame_nav, text="◀ Anterior", 
                 command=self.voltar, 
                 bg="#95A5A6", fg="white",
                 font=("Arial", 9)).pack(side="left", fill="x", expand=True, padx=2)
        tk.Button(frame_nav, text="Próximo ▶", 
                 command=self.avancar, 
                 bg="#95A5A6", fg="white",
                 font=("Arial", 9)).pack(side="right", fill="x", expand=True, padx=2)
        
        # RODAPÉ
        frame_rodape = tk.Frame(self.root, bg="#34495E", pady=10)
        frame_rodape.pack(side="bottom", fill="x")
        tk.Label(frame_rodape, 
                text=f"Faturas salvas em: {self.download_path}",
                font=("Arial", 8), 
                fg="#BDC3C7", bg="#34495E").pack()
        
        # ATALHOS
        self.root.bind('<Return>', lambda e: self.executar_robo())
        self.root.bind('<Right>', lambda e: self.avancar())
        self.root.bind('<Left>', lambda e: self.voltar())
        self.root.bind('<Escape>', lambda e: self.pular_cliente())

    def iniciar_navegador(self):
        def iniciar():
            try:
                self.bot.abrir_navegador()
                self.status_var.set("✅ Navegador iniciado!\n1. Faça login manualmente\n2. Clique no botão verde quando estiver logado")
                self.btn_baixar.config(state="normal", bg="#27AE60")
                self.btn_pular.config(state="normal")
            except Exception as e:
                self.status_var.set(f"❌ Erro ao iniciar navegador: {str(e)}")
        
        thread = threading.Thread(target=iniciar)
        thread.daemon = True
        thread.start()

    def atualizar_tela(self):
        if not self.dados or self.index_atual >= len(self.dados):
            return
            
        item = self.dados[self.index_atual]
        
        self.lbl_contador.config(
            text=f"Cliente {self.index_atual + 1} de {len(self.dados)}"
        )
        
        nome = str(item.get('Nome', 'Cliente Sem Nome'))
        self.lbl_nome_cliente.config(text=nome[:30] + "..." if len(nome) > 30 else nome)
        
        self.entry_uc.delete(0, tk.END)
        self.entry_uc.insert(0, str(item.get('Conta Contrato', '')).replace('.0', ''))
        
        self.entry_login.delete(0, tk.END)
        self.entry_login.insert(0, str(item.get('CNPJ/CPF', '')).replace('.0', ''))
        
        self.entry_senha.delete(0, tk.END)
        self.entry_senha.insert(0, str(item.get('Acesso equatorial', '')).strip())
        
        self.status_var.set(f"✅ Dados carregados para {nome}\nPronto para baixar a última fatura")
        self.processo_em_andamento = False

    def copiar(self, texto):
        if texto:
            pyperclip.copy(texto)
            self.status_var.set(f"📋 Texto copiado: {texto[:20]}...")

    def executar_robo(self):
        if self.processo_em_andamento:
            return
            
        if not self.bot.driver:
            self.status_var.set("❌ Navegador não inicializado")
            return
            
        self.processo_em_andamento = True
        self.btn_baixar.config(state="disabled", bg="#7F8C8D", 
                              text="⏳ PROCESSANDO...")
        self.btn_pular.config(state="disabled")
        self.root.update()
        
        uc = self.entry_uc.get().strip()
        
        if not uc:
            self.status_var.set("❌ UC não encontrada")
            self.btn_baixar.config(state="normal", bg="#27AE60", 
                                  text="🤖 BAIXAR ÚLTIMA FATURA")
            self.btn_pular.config(state="normal")
            self.processo_em_andamento = False
            return
        
        self.status_var.set(f"⏳ Baixando a última fatura para UC {uc}...")
        self.root.update()
        
        # Executa em thread separada para não travar a interface
        def executar_download():
            resultado = self.bot.baixar_ultima_fatura(uc)
            
            # Atualiza a interface na thread principal
            self.root.after(0, lambda: self.processar_resultado(resultado))
        
        thread = threading.Thread(target=executar_download)
        thread.daemon = True
        thread.start()

    def processar_resultado(self, resultado):
        if "Sucesso" in resultado:
            self.status_var.set(f"✅ {resultado} baixada com sucesso!\nRealizando logout...")
            self.root.update()
            
            # Faz logout e avança para próximo cliente
            def finalizar():
                time.sleep(2)
                self.bot.fazer_logout()
                time.sleep(2)
                self.root.after(0, self.avancar)
            
            threading.Thread(target=finalizar, daemon=True).start()
        else:
            self.status_var.set(f"❌ Falha no download: {resultado}")
            messagebox.showwarning("Atenção", 
                                 f"O robô encontrou um problema:\n\n{resultado}\n\n"
                                 "1. Verifique se está logado\n"
                                 "2. Verifique se existem faturas disponíveis\n"
                                 "3. Tente fazer manualmente e clique em 'PULAR'")
            self.btn_baixar.config(state="normal", bg="#E67E22", 
                                  text="🤖 TENTAR NOVAMENTE")
            self.btn_pular.config(state="normal")
            self.processo_em_andamento = False

    def pular_cliente(self):
        if self.processo_em_andamento:
            return
            
        self.status_var.set("⏭️ Pulando cliente atual...")
        self.root.update()
        
        def fazer_logout_e_avancar():
            self.bot.fazer_logout()
            time.sleep(2)
            self.root.after(0, self.avancar)
        
        threading.Thread(target=fazer_logout_e_avancar, daemon=True).start()

    def avancar(self):
        if self.index_atual < len(self.dados) - 1:
            self.index_atual += 1
            self.atualizar_tela()
            self.btn_baixar.config(state="normal", bg="#27AE60", 
                                  text="🤖 BAIXAR ÚLTIMA FATURA")
            self.btn_pular.config(state="normal")
            self.processo_em_andamento = False
        else:
            messagebox.showinfo("Fim da Lista", 
                              "✅ Todos os clientes foram processados!\n\n"
                              f"Faturas salvas em: {self.download_path}")
            self.status_var.set("✅ Processo finalizado!")
            self.btn_baixar.config(state="disabled", bg="#7F8C8D")
            self.btn_pular.config(state="disabled")

    def voltar(self):
        if self.processo_em_andamento:
            return
            
        if self.index_atual > 0:
            self.index_atual -= 1
            self.atualizar_tela()

# =============================================================================
# 3. EXECUÇÃO PRINCIPAL
# =============================================================================
if __name__ == "__main__":
    base_dir = os.getcwd()
    caminho_excel = os.path.join(base_dir, "output", "Cad_RateioConsumo_Final.xlsx") 
    
    if not os.path.exists(caminho_excel):
        print(f"❌ Arquivo Excel não encontrado: {caminho_excel}")
        print("Certifique-se de que o arquivo está na pasta 'output'")
        input("Pressione Enter para sair...")
    else:
        print(f"✅ Excel encontrado: {caminho_excel}")
        
        root = tk.Tk()
        app = PainelControle(root, caminho_excel)
        
        # Centraliza a janela
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'{width}x{height}+{x}+{y}')
        
        root.mainloop()