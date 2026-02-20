import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog

import core
from core import *



ctk.set_default_color_theme("green")
ctk.set_appearance_mode("dark")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Excel Manager")
        self.geometry("700x450")
        self.iconbitmap("logo.ico")

        # Sidebar (Coluna 0)
        self.sidebar = ctk.CTkFrame(self, width=180, height=9000)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        # Frame de Conteúdo (COLUNA 1)
        self.frame_conteudo = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_conteudo.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.grid_columnconfigure(1, weight=1)

        self.tela_home()


        # Botão que chama a interface de Busca
        self.btn_busca = ctk.CTkButton(self.sidebar, text="Buscar", command=self.tela_busca)
        self.btn_busca.grid(row=1, column=0, padx=20, pady=10)

        self.btn_criar = ctk.CTkButton(self.sidebar, text="Criar Planilha", command=self.tela_criar)
        self.btn_criar.grid(row=2, column=0, padx=20, pady=10)

        self.btn_preencher = ctk.CTkButton(self.sidebar, text="Preencher Planilha", command=self.tela_preencher)
        self.btn_preencher.grid(row=3, column=0, padx=20, pady=10)

        self.btn_view = ctk.CTkButton(self.sidebar, text="Visualizar Planilha", command=self.tela_visualizar)
        self.btn_view.grid(row=4, column=0, padx=20, pady=10)

        self.log_resultado = ctk.CTkLabel(self.frame_conteudo, text="")

    def tela_home(self):
        self.limpar_tela()

        # Espaçador superior para centralizar verticalmente
        ctk.CTkLabel(self.frame_conteudo, text="").pack(pady=20)

        # Logo ou Ícone Grande (Pode usar aquele que você vai criar)
        lbl_logo = ctk.CTkLabel(self.frame_conteudo, text="📊", font=("Arial", 80))
        lbl_logo.pack(pady=10)

        lbl_boas_vindas = ctk.CTkLabel(
            self.frame_conteudo,
            text="Excel Manager",
            font=("Arial", 24, "bold")
        )
        lbl_boas_vindas.pack()

        lbl_status = ctk.CTkLabel(
            self.frame_conteudo,
            text="Selecione uma opção no menu lateral para começar",
            font=("Arial", 14, "italic"),
            text_color="gray"
        )
        lbl_status.pack(pady=10)

        # Pequeno card de ajuda rápida
        card_ajuda = ctk.CTkFrame(self.frame_conteudo, fg_color="#2b2b2b", corner_radius=10)
        card_ajuda.pack(pady=30, padx=40, fill="x")

        ajuda_texto = (
            "💡 Dicas Rápidas:\n\n"
            "• Criar: Gera um arquivo .xlsx novo do zero.\n"
            "• Injetar: Adiciona dados em uma planilha que já existe.\n"
            "• Inspeção: Veja cabeçalhos e contagem de linhas sem abrir o Excel."
        )
        ctk.CTkLabel(card_ajuda, text=ajuda_texto, justify="left", padx=20, pady=20).pack()

    def limpar_tela(self):
        # Destrói todos os widgets que estão dentro do frame de conteúdo
        for widget in self.frame_conteudo.winfo_children():
            widget.destroy()

    def tela_busca(self):
        self.limpar_tela()

        # Agora desenha só o que a busca precisa
        lbl = ctk.CTkLabel(self.frame_conteudo, text="Digite o termo para busca:")
        lbl.pack(pady=10)

        self.entry_filtro = ctk.CTkEntry(self.frame_conteudo, width=300)
        self.entry_filtro.pack(pady=10)

        btn_confirmar = ctk.CTkButton(self.frame_conteudo, text="Procurar", command=self.executar_logica_busca)
        btn_confirmar.pack(pady=10)

    def tela_criar(self):
        self.limpar_tela()

        # Título
        ctk.CTkLabel(self.frame_conteudo, text="GERAR NOVA PLANILHA", font=("Arial", 18, "bold")).pack(pady=(0, 20))

        # 1. Nome do Arquivo
        ctk.CTkLabel(self.frame_conteudo, text="Nome do Arquivo (sem .xlsx):").pack(pady=5)
        self.entry_nome = ctk.CTkEntry(self.frame_conteudo, width=300, placeholder_text="ex: relatorio_vendas")
        self.entry_nome.pack(pady=5)

        # 2. Quantidade de Colunas (Crucial para a sua lógica de quebra de linha)
        ctk.CTkLabel(self.frame_conteudo, text="Quantidade de Colunas:").pack(pady=5)
        self.entry_num_cols = ctk.CTkEntry(self.frame_conteudo, width=100, placeholder_text="Ex: 3")
        self.entry_num_cols.pack(pady=5)

        # 3. Dados (Lista separada por vírgula)
        ctk.CTkLabel(self.frame_conteudo, text="Dados (Cabeçalhos e Conteúdo separados por vírgula):").pack(pady=5)
        self.txt_dados = ctk.CTkTextbox(self.frame_conteudo, width=450, height=150)
        self.txt_dados.pack(pady=10)
        ctk.CTkLabel(self.frame_conteudo, text="Ex: Nome, Idade, Cargo, João, 25, Dev, Maria, 30, Designer",
                     font=("Arial", 10, "italic")).pack()

        # Botão Gerar
        btn_gerar = ctk.CTkButton(
            self.frame_conteudo,
            text="CRIAR E PREENCHER",
            fg_color="green",
            command=self.executar_criacao_completa
        )
        btn_gerar.pack(pady=20)

    def tela_preencher(self):
        self.limpar_tela()

        ctk.CTkLabel(self.frame_conteudo, text="INJETAR DADOS EM PLANILHA EXISTENTE", font=("Arial", 18, "bold")).pack(
            pady=(0, 20))

        # 1. Selecionar o Arquivo
        self.path_alvo = ""  # Variável para guardar o caminho
        self.btn_file = ctk.CTkButton(self.frame_conteudo, text="📁 Selecionar Planilha", command=self.abrir_explorador)
        self.btn_file.pack(pady=5)
        self.lbl_arquivo = ctk.CTkLabel(self.frame_conteudo, text="Nenhum arquivo selecionado",
                                        font=("Arial", 10, "italic"))
        self.lbl_arquivo.pack(pady=(0, 10))

        # 2. Quantidade de Colunas
        ctk.CTkLabel(self.frame_conteudo, text="Quantidade de Colunas (Largura):").pack(pady=5)
        self.entry_num_cols = ctk.CTkEntry(self.frame_conteudo, width=100)
        self.entry_num_cols.pack(pady=5)

        # 3. Dados para Injetar
        ctk.CTkLabel(self.frame_conteudo, text="Dados a serem inseridos (separados por vírgula):").pack(pady=5)
        self.txt_dados = ctk.CTkTextbox(self.frame_conteudo, width=450, height=150)
        self.txt_dados.pack(pady=10)

        # Botão Injetar
        self.btn_executar = ctk.CTkButton(self.frame_conteudo, text="🚀 INJETAR DADOS", fg_color="green",
                                          command=self.executar_preenchimento)
        self.btn_executar.pack(pady=20)

    def abrir_explorador(self):
        self.path_alvo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
        if self.path_alvo:
            # Mostra só o nome do arquivo na label pra não ficar um texto gigante
            nome_limpo = self.path_alvo.split("/")[-1]
            self.lbl_arquivo.configure(text=f"Selecionado: {nome_limpo}", text_color="cyan")

    def executar_preenchimento(self):

        if not hasattr(self, "log_resultado") or not self.log_resultado.winfo_exists():
            self.log_resultado = ctk.CTkLabel(self.frame_conteudo, text="")
            self.log_resultado.pack(pady=10)

        try:
            if not self.path_alvo:
                self.log_resultado.configure(text="❌ Erro: Selecione um arquivo primeiro!", text_color="red")
                return
            cols = int(self.entry_num_cols.get())
            raw_text = self.txt_dados.get("1.0", "end").strip()
            lista_dados = [item.strip() for item in raw_text.split(",") if item.strip()]

            resultado_path = core.preencherPlanilha(self.path_alvo, lista_dados, cols)
            self.log_resultado.configure(text=f"✅ Dados injetados em: {resultado_path.split('/')[-1]}", text_color="green")

        except ValueError:
            self.log_resultado.configure(text="❌ ERRO: Colunas precisa ser um número!", text_color="red")

        except Exception as e:
            self.log_resultado.configure(text=f"⚠️ Erro: {e}", text_color="yellow")

    def executar_criacao_completa(self):
        # GARANTIA: Se a label de log não existir no self, a gente cria ela aqui
        if not hasattr(self, "log_resultado") or not self.log_resultado.winfo_exists():
            self.log_resultado = ctk.CTkLabel(self.frame_conteudo, text="")
            self.log_resultado.pack(pady=10)

        try:
            nome = self.entry_nome.get()
            # CONVERSÃO ESSENCIAL: Tem que ser int para a lógica do core funcionar
            qtd_cols = int(self.entry_num_cols.get())

            # Pegando o texto (0.0 ou 1.0 funcionam no CTkTextbox)
            conteudo_raw = self.txt_dados.get("1.0", "end").strip()

            lista_final = [item.strip() for item in conteudo_raw.split(",") if item.strip()]

            # CHAMADA DO CORE
            arquivo_gerado = core.criarPlanilha(lista_final, qtd_cols, nome)

            # FEEDBACK
            self.log_resultado.configure(text=f"✅ Sucesso! '{arquivo_gerado}' criado.\nItens: {len(lista_final)}",
                                         text_color="green")

        except ValueError:
            self.log_resultado.configure(text="❌ ERRO: Digite um NÚMERO em colunas!", text_color="red")
        except Exception as e:
            self.log_resultado.configure(text=f"⚠️ Erro: {e}", text_color="yellow")

    def executar_logica_busca(self):
        # Aqui você pega o valor e manda pro core.py
        path = filedialog.askopenfilename(
            title="Buscar arquivos excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if not path:
            return

        valor = str(self.entry_filtro.get())
        try:
            resultado = buscarRegistro(path, valor)
            self.label = ctk.CTkLabel(
            self.frame_conteudo,
            text=resultado,
            font=("Consolas", 13),
            fg_color="#1a1a1a",
            text_color="#d1d1d1",   # Texto cinza claro para não agredir os olhos
            corner_radius=10
        )
            self.label.pack(padx=20, pady=10)

        except Exception as e:
            print(f"{e}")

    def tela_visualizar(self):
        self.limpar_tela()

        ctk.CTkLabel(self.frame_conteudo, text="INSPEÇÃO DE PLANILHA", font=("Arial", 18, "bold")).pack(pady=(0, 20))

        # 1. Seletor de Arquivo (Reaproveitando a lógica)
        self.path_inspecao = ""
        self.btn_file = ctk.CTkButton(self.frame_conteudo, text="📂 Escolher Arquivo para Ler",
                                      command=self.abrir_explorador_inspecao)
        self.btn_file.pack(pady=5)
        self.lbl_arq_inspecao = ctk.CTkLabel(self.frame_conteudo, text="Nenhum arquivo selecionado",
                                             font=("Arial", 10, "italic"))
        self.lbl_arq_inspecao.pack(pady=(0, 20))

        # Frame para os botões ficarem lado a lado (estilo Dashboard)
        frame_botoes = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        frame_botoes.pack(pady=10)

        # Botão 1: Contar Linhas
        self.btn_linhas = ctk.CTkButton(frame_botoes, text="Contar Linhas", width=120, command=self.exec_contar_linhas)
        self.btn_linhas.grid(row=0, column=0, padx=5)

        # Botão 2: Ler Cabeçalhos
        self.btn_headers = ctk.CTkButton(frame_botoes, text="Ver Cabeçalhos", width=120, command=self.exec_ler_headers)
        self.btn_headers.grid(row=0, column=1, padx=5)

        # Botão 3: Listar Tudo
        self.btn_listar = ctk.CTkButton(frame_botoes, text="Listar Planilha", width=120, command=self.exec_listar_tudo)
        self.btn_listar.grid(row=0, column=2, padx=5)

        # Output para mostrar o resultado das funções
        self.txt_output = ctk.CTkTextbox(self.frame_conteudo, width=500, height=200, font=("Consolas", 12))
        self.txt_output.pack(pady=20)

    def abrir_explorador_inspecao(self):
        self.path_inspecao = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if self.path_inspecao:
            self.lbl_arq_inspecao.configure(text=self.path_inspecao.split("/")[-1], text_color="cyan")

    def exec_contar_linhas(self):
        try:
            # Supondo que sua função seja core.contarLinhas(path)
            total = core.contarLinhas(self.path_inspecao)
            self.atualizar_output(f"📊 Total de linhas preenchidas: {total}")
        except Exception as e:
            self.atualizar_output(f"❌ Erro ao contar: {e}")

    def exec_ler_headers(self):
        try:
            # Supondo que retorne uma lista: ['ID', 'Nome', 'Cargo']
            headers = core.lerCabecalhos(self.path_inspecao)
            texto = "📋 Cabeçalhos encontrados:\n" + headers
            self.atualizar_output(texto)
        except Exception as e:
            self.atualizar_output(f"❌ Erro nos cabeçalhos: {e}")

    def exec_listar_tudo(self):
        try:
            # Supondo que retorne uma string formatada ou lista de listas
            dados = core.listarPlanilha(self.path_inspecao)
            self.atualizar_output(f"📄 Conteúdo da Planilha:\n\n{dados}")
        except Exception as e:
            self.atualizar_output(f"❌ Erro ao listar: {e}")

    def atualizar_output(self, texto):
        self.txt_output.delete("1.0", "end")
        self.txt_output.insert("1.0", texto)

if __name__ == "__main__":
    app = App()
    app.mainloop()