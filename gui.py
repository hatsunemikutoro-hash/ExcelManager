# =============================================================================
# EXCEL MANAGER — gui.py
# Responsabilidade: interface gráfica. Toda lógica de dados fica no core.py.
# =============================================================================

import threading
import customtkinter as ctk
from tkinter import filedialog

import core


# =============================================================================
# SEÇÃO 1 — CONFIGURAÇÃO GLOBAL DO TEMA
# =============================================================================

ctk.set_default_color_theme("green")
ctk.set_appearance_mode("dark")

FONTE_TITULO  = ("Arial", 18, "bold")
FONTE_LABEL   = ("Arial", 13)
FONTE_PEQUENA = ("Arial", 10, "italic")
FONTE_MONO    = ("Consolas", 12)
COR_SUCESSO   = "#2ecc71"
COR_ERRO      = "#e74c3c"
COR_AVISO     = "#f39c12"
COR_INFO      = "#3498db"


# =============================================================================
# SEÇÃO 2 — CLASSE PRINCIPAL
# =============================================================================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Excel Manager")
        self.geometry("820x560")
        self.minsize(700, 480)
        self._tema_atual = "dark"
        self._construir_layout()
        self.tela_home()

    # -------------------------------------------------------------------------
    # LAYOUT BASE
    # -------------------------------------------------------------------------

    def _construir_layout(self):
        """Constrói a sidebar e o frame de conteúdo que persistem entre telas."""
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Sidebar fixa à esquerda
        self.sidebar = ctk.CTkFrame(self, width=190)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        self.sidebar.grid_rowconfigure(20, weight=1)  # empurra toggle para baixo

        # Frame de conteúdo dinâmico à direita
        self.frame_conteudo = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_conteudo.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.frame_conteudo.grid_columnconfigure(0, weight=1)
        self.frame_conteudo.grid_rowconfigure(0, weight=1)

        # Botões de navegação
        nav_itens = [
            ("🏠  Início",            self.tela_home),
            ("🔍  Buscar",            self.tela_busca),
            ("✨  Criar Planilha",    self.tela_criar),
            ("💉  Preencher",         self.tela_preencher),
            ("🔗  Mesclar Planilhas", self.tela_mesclar),
            ("🔎  Inspecionar",       self.tela_visualizar),
            ("📋  Histórico",         self.tela_historico),
        ]
        for i, (texto, comando) in enumerate(nav_itens):
            btn = ctk.CTkButton(
                self.sidebar, text=texto, anchor="w",
                command=comando, width=170, height=36,
                fg_color="transparent", hover_color="#2d6a4f",
                text_color=("gray10", "gray90"),
            )
            btn.grid(row=i, column=0, padx=10, pady=(10 if i == 0 else 3, 0))

        # Toggle de tema (fica no rodapé da sidebar)
        self.btn_tema = ctk.CTkButton(
            self.sidebar, text="☀️  Tema Claro", width=170,
            fg_color="transparent", hover_color="#2d6a4f",
            text_color=("gray10", "gray90"),
            command=self._alternar_tema, anchor="w"
        )
        self.btn_tema.grid(row=21, column=0, padx=10, pady=15, sticky="s")

    def _limpar_tela(self):
        for widget in self.frame_conteudo.winfo_children():
            widget.destroy()

    def _alternar_tema(self):
        self._tema_atual = "light" if self._tema_atual == "dark" else "dark"
        ctk.set_appearance_mode(self._tema_atual)
        icone = "🌙  Tema Escuro" if self._tema_atual == "light" else "☀️  Tema Claro"
        self.btn_tema.configure(text=icone)

    # -------------------------------------------------------------------------
    # UTILITÁRIOS DE UI REUTILIZÁVEIS
    # -------------------------------------------------------------------------

    def _criar_titulo(self, texto: str):
        lbl = ctk.CTkLabel(self.frame_conteudo, text=texto, font=FONTE_TITULO)
        lbl.pack(pady=(0, 12))
        return lbl

    def _criar_log(self) -> ctk.CTkLabel:
        log = ctk.CTkLabel(self.frame_conteudo, text="", font=FONTE_LABEL, wraplength=500)
        log.pack(pady=4)
        return log

    def _criar_progressbar(self) -> ctk.CTkProgressBar:
        pb = ctk.CTkProgressBar(self.frame_conteudo, mode="indeterminate", width=400)
        pb.pack(pady=4)
        pb.pack_forget()
        return pb

    def _set_log(self, log: ctk.CTkLabel, texto: str, cor: str = "white"):
        self.after(0, lambda: log.configure(text=texto, text_color=cor))

    def _iniciar_progresso(self, pb: ctk.CTkProgressBar):
        self.after(0, lambda: (pb.pack(pady=4), pb.start()))

    def _parar_progresso(self, pb: ctk.CTkProgressBar):
        self.after(0, lambda: (pb.stop(), pb.pack_forget()))

    def _selecionar_arquivo(self, label: ctk.CTkLabel, atributo: str,
                            titulo: str = "Selecionar Planilha"):
        path = filedialog.askopenfilename(
            title=titulo,
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if path:
            setattr(self, atributo, path)
            label.configure(text=f"📄 {path.split('/')[-1]}", text_color=COR_INFO)

    def _selecionar_destino(self, nome_sugerido: str = "resultado") -> str | None:
        return filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            initialfile=nome_sugerido,
            title="Salvar planilha como..."
        )

    def _rodar_em_thread(self, fn, *args):
        threading.Thread(target=fn, args=args, daemon=True).start()

    def _criar_seletor_arquivo(self, atributo: str, titulo_btn: str = "📁 Selecionar Planilha"):
        """Cria botão + label de seleção de arquivo e retorna a label."""
        setattr(self, atributo, "")
        lbl = ctk.CTkLabel(self.frame_conteudo, text="Nenhum arquivo selecionado.", font=FONTE_PEQUENA)
        ctk.CTkButton(
            self.frame_conteudo, text=titulo_btn,
            command=lambda: self._selecionar_arquivo(lbl, atributo)
        ).pack(pady=(5, 2))
        lbl.pack(pady=(0, 8))
        return lbl

    def _criar_entry(self, placeholder: str, width: int = 340) -> ctk.CTkEntry:
        e = ctk.CTkEntry(self.frame_conteudo, width=width, placeholder_text=placeholder)
        e.pack(pady=4)
        return e

    def _criar_textbox(self, height: int = 110) -> ctk.CTkTextbox:
        tb = ctk.CTkTextbox(self.frame_conteudo, width=480, height=height, font=FONTE_MONO)
        tb.pack(pady=4)
        return tb

    def _get_lista(self, textbox: ctk.CTkTextbox) -> list[str]:
        raw = textbox.get("1.0", "end").strip()
        return [item.strip() for item in raw.split(",") if item.strip()]

    def _get_colunas(self, entry: ctk.CTkEntry) -> int:
        val = entry.get().strip()
        if not val.isdigit() or int(val) <= 0:
            raise ValueError("'Número de colunas' deve ser um número inteiro maior que zero.")
        return int(val)


# =============================================================================
# SEÇÃO 3 — TELA HOME
# =============================================================================

    def tela_home(self):
        self._limpar_tela()
        ctk.CTkLabel(self.frame_conteudo, text="📊", font=("Arial", 72)).pack(pady=(20, 4))
        ctk.CTkLabel(self.frame_conteudo, text="Excel Manager", font=("Arial", 26, "bold")).pack()
        ctk.CTkLabel(
            self.frame_conteudo,
            text="Selecione uma opção no menu lateral para começar.",
            font=FONTE_PEQUENA, text_color="gray"
        ).pack(pady=6)

        card = ctk.CTkFrame(self.frame_conteudo, fg_color="#1e3a2f", corner_radius=12)
        card.pack(pady=16, padx=30, fill="x")
        ctk.CTkLabel(
            card,
            text=(
                "💡  Dicas rápidas\n\n"
                "• Criar → gera um .xlsx novo do zero com seus dados.\n"
                "• Preencher → injeta dados em uma planilha existente.\n"
                "• Mesclar → une múltiplas planilhas em um só arquivo.\n"
                "• Inspecionar → lê cabeçalhos, linhas e conteúdo sem abrir o Excel.\n"
                "• Buscar → localiza qualquer valor em qualquer célula.\n"
                "• Histórico → registro de todas as operações realizadas."
            ),
            justify="left", font=FONTE_LABEL, padx=20, pady=18
        ).pack()


# =============================================================================
# SEÇÃO 4 — TELA BUSCA
# =============================================================================

    def tela_busca(self):
        self._limpar_tela()
        self._criar_titulo("BUSCAR REGISTRO")
        self._criar_seletor_arquivo("path_busca")

        ctk.CTkLabel(self.frame_conteudo, text="Termo de busca:", font=FONTE_LABEL).pack(pady=(6, 2))
        self.entry_filtro = self._criar_entry("Ex: João Silva")

        ctk.CTkButton(
            self.frame_conteudo, text="🔍 Procurar",
            command=self._executar_busca
        ).pack(pady=10)

        self._pb_busca  = self._criar_progressbar()
        self._log_busca = self._criar_log()

    def _executar_busca(self):
        if not self.path_busca:
            self._set_log(self._log_busca, "❌ Selecione um arquivo primeiro.", COR_ERRO)
            return
        filtro = self.entry_filtro.get().strip()
        if not filtro:
            self._set_log(self._log_busca, "❌ Digite um termo para buscar.", COR_ERRO)
            return

        self._iniciar_progresso(self._pb_busca)
        self._set_log(self._log_busca, "⏳ Buscando...", COR_AVISO)

        def tarefa():
            try:
                resultado = core.buscarRegistro(self.path_busca, filtro)
                if resultado:
                    msg = (f"✅ Encontrado: \"{resultado['valor']}\"  —  "
                           f"Coluna {resultado['coluna']}, Linha {resultado['linha']}.")
                    self._set_log(self._log_busca, msg, COR_SUCESSO)
                    core.registrarHistorico("Busca", f"'{filtro}' em {self.path_busca.split('/')[-1]}")
                else:
                    self._set_log(self._log_busca, f"🔎 '{filtro}' não encontrado.", COR_AVISO)
            except Exception as e:
                self._set_log(self._log_busca, f"⚠️ {e}", COR_ERRO)
            finally:
                self._parar_progresso(self._pb_busca)

        self._rodar_em_thread(tarefa)


# =============================================================================
# SEÇÃO 5 — TELA CRIAR PLANILHA
# =============================================================================

    def tela_criar(self):
        self._limpar_tela()
        self._criar_titulo("CRIAR NOVA PLANILHA")

        ctk.CTkLabel(self.frame_conteudo, text="Nome do arquivo:", font=FONTE_LABEL).pack(pady=(4, 2))
        self.entry_nome_criar = self._criar_entry("ex: relatorio_vendas")

        ctk.CTkLabel(self.frame_conteudo, text="Número de colunas:", font=FONTE_LABEL).pack(pady=(4, 2))
        self.entry_cols_criar = self._criar_entry("ex: 3", width=120)

        ctk.CTkLabel(self.frame_conteudo, text="Dados separados por vírgula:", font=FONTE_LABEL).pack(pady=(4, 2))
        self.txt_dados_criar = self._criar_textbox()
        ctk.CTkLabel(
            self.frame_conteudo,
            text="Ex: Nome, Idade, Cargo, João, 25, Dev, Maria, 30, Designer",
            font=FONTE_PEQUENA
        ).pack()

        frame_btns = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        frame_btns.pack(pady=10)
        ctk.CTkButton(frame_btns, text="👁️ Preview", width=140,
                      command=self._preview_criar).grid(row=0, column=0, padx=6)
        ctk.CTkButton(frame_btns, text="💾 Salvar Planilha", width=180,
                      fg_color=COR_SUCESSO, command=self._executar_criacao).grid(row=0, column=1, padx=6)

        self._pb_criar  = self._criar_progressbar()
        self._log_criar = self._criar_log()

    def _preview_criar(self):
        """Abre janela de preview antes de salvar."""
        try:
            cols  = self._get_colunas(self.entry_cols_criar)
            lista = self._get_lista(self.txt_dados_criar)
            if not lista:
                self._set_log(self._log_criar, "❌ Insira ao menos um dado.", COR_ERRO)
                return
            grade = core.previewDados(lista, cols)
            self._mostrar_preview(grade)
        except ValueError as e:
            self._set_log(self._log_criar, f"❌ {e}", COR_ERRO)

    def _executar_criacao(self):
        try:
            nome  = self.entry_nome_criar.get().strip() or "resultado"
            cols  = self._get_colunas(self.entry_cols_criar)
            lista = self._get_lista(self.txt_dados_criar)
            if not lista:
                self._set_log(self._log_criar, "❌ Insira ao menos um dado.", COR_ERRO)
                return
        except ValueError as e:
            self._set_log(self._log_criar, f"❌ {e}", COR_ERRO)
            return

        destino = self._selecionar_destino(nome)
        if not destino:
            self._set_log(self._log_criar, "⚠️ Operação cancelada.", COR_AVISO)
            return

        self._iniciar_progresso(self._pb_criar)
        self._set_log(self._log_criar, "⏳ Criando planilha...", COR_AVISO)

        def tarefa():
            try:
                caminho = core.criarPlanilha(lista, cols, destino)
                nome_arq = caminho.split("/")[-1]
                self._set_log(self._log_criar,
                              f"✅ '{nome_arq}' criado com {len(lista)} itens.", COR_SUCESSO)
                core.registrarHistorico("Criar", f"{nome_arq} — {len(lista)} itens")
            except Exception as e:
                self._set_log(self._log_criar, f"⚠️ {e}", COR_ERRO)
            finally:
                self._parar_progresso(self._pb_criar)

        self._rodar_em_thread(tarefa)


# =============================================================================
# SEÇÃO 6 — TELA PREENCHER PLANILHA
# =============================================================================

    def tela_preencher(self):
        self._limpar_tela()
        self._criar_titulo("INJETAR DADOS EM PLANILHA")
        self._criar_seletor_arquivo("path_preencher")

        ctk.CTkLabel(self.frame_conteudo, text="Número de colunas:", font=FONTE_LABEL).pack(pady=(4, 2))
        self.entry_cols_preencher = self._criar_entry("ex: 3", width=120)

        ctk.CTkLabel(self.frame_conteudo, text="Dados a inserir (separados por vírgula):", font=FONTE_LABEL).pack(pady=(4, 2))
        self.txt_dados_preencher = self._criar_textbox()

        frame_btns = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        frame_btns.pack(pady=10)
        ctk.CTkButton(frame_btns, text="👁️ Preview", width=140,
                      command=self._preview_preencher).grid(row=0, column=0, padx=6)
        ctk.CTkButton(frame_btns, text="🚀 Injetar Dados", width=180,
                      fg_color=COR_SUCESSO, command=self._executar_preenchimento).grid(row=0, column=1, padx=6)

        self._pb_preencher  = self._criar_progressbar()
        self._log_preencher = self._criar_log()

    def _preview_preencher(self):
        try:
            cols  = self._get_colunas(self.entry_cols_preencher)
            lista = self._get_lista(self.txt_dados_preencher)
            if not lista:
                self._set_log(self._log_preencher, "❌ Insira ao menos um dado.", COR_ERRO)
                return
            grade = core.previewDados(lista, cols)
            self._mostrar_preview(grade)
        except ValueError as e:
            self._set_log(self._log_preencher, f"❌ {e}", COR_ERRO)

    def _executar_preenchimento(self):
        if not self.path_preencher:
            self._set_log(self._log_preencher, "❌ Selecione um arquivo primeiro.", COR_ERRO)
            return
        try:
            cols  = self._get_colunas(self.entry_cols_preencher)
            lista = self._get_lista(self.txt_dados_preencher)
            if not lista:
                self._set_log(self._log_preencher, "❌ Insira ao menos um dado.", COR_ERRO)
                return
        except ValueError as e:
            self._set_log(self._log_preencher, f"❌ {e}", COR_ERRO)
            return

        self._iniciar_progresso(self._pb_preencher)
        self._set_log(self._log_preencher, "⏳ Injetando dados...", COR_AVISO)

        def tarefa():
            try:
                caminho = core.preencherPlanilha(self.path_preencher, lista, cols)
                nome_arq = caminho.split("/")[-1]
                self._set_log(self._log_preencher,
                              f"✅ {len(lista)} itens injetados em '{nome_arq}'.", COR_SUCESSO)
                core.registrarHistorico("Preencher", f"{nome_arq} — {len(lista)} itens")
            except Exception as e:
                self._set_log(self._log_preencher, f"⚠️ {e}", COR_ERRO)
            finally:
                self._parar_progresso(self._pb_preencher)

        self._rodar_em_thread(tarefa)


# =============================================================================
# SEÇÃO 7 — TELA MESCLAR PLANILHAS
# =============================================================================

    def tela_mesclar(self):
        self._limpar_tela()
        self._criar_titulo("MESCLAR PLANILHAS")

        ctk.CTkLabel(
            self.frame_conteudo,
            text="Selecione 2 ou mais arquivos Excel para unir em um só.\nOs cabeçalhos do primeiro arquivo serão mantidos.",
            font=FONTE_LABEL, justify="center"
        ).pack(pady=(0, 10))

        self._arquivos_mesclar = []

        ctk.CTkButton(
            self.frame_conteudo, text="📁 Adicionar Arquivos",
            command=self._adicionar_arquivos_mesclar
        ).pack(pady=5)

        self.lbl_arquivos_mesclar = ctk.CTkLabel(
            self.frame_conteudo, text="Nenhum arquivo adicionado.", font=FONTE_PEQUENA
        )
        self.lbl_arquivos_mesclar.pack(pady=(0, 10))

        frame_btns = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        frame_btns.pack(pady=8)
        ctk.CTkButton(frame_btns, text="🗑️ Limpar Lista", width=140,
                      fg_color="#555", command=self._limpar_mesclar).grid(row=0, column=0, padx=6)
        ctk.CTkButton(frame_btns, text="🔗 Mesclar e Salvar", width=180,
                      fg_color=COR_SUCESSO, command=self._executar_mesclar).grid(row=0, column=1, padx=6)

        self._pb_mesclar  = self._criar_progressbar()
        self._log_mesclar = self._criar_log()

    def _adicionar_arquivos_mesclar(self):
        paths = filedialog.askopenfilenames(
            title="Selecionar arquivos para mesclar",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if paths:
            self._arquivos_mesclar.extend(paths)
            nomes = "\n".join(f"  • {p.split('/')[-1]}" for p in self._arquivos_mesclar)
            self.lbl_arquivos_mesclar.configure(
                text=f"{len(self._arquivos_mesclar)} arquivo(s):\n{nomes}",
                text_color=COR_INFO
            )

    def _limpar_mesclar(self):
        self._arquivos_mesclar = []
        self.lbl_arquivos_mesclar.configure(text="Nenhum arquivo adicionado.", text_color="gray")
        self._set_log(self._log_mesclar, "", "white")

    def _executar_mesclar(self):
        if len(self._arquivos_mesclar) < 2:
            self._set_log(self._log_mesclar, "❌ Adicione ao menos 2 arquivos.", COR_ERRO)
            return

        destino = self._selecionar_destino("planilha_mesclada")
        if not destino:
            self._set_log(self._log_mesclar, "⚠️ Operação cancelada.", COR_AVISO)
            return

        self._iniciar_progresso(self._pb_mesclar)
        self._set_log(self._log_mesclar, "⏳ Mesclando arquivos...", COR_AVISO)
        arquivos = list(self._arquivos_mesclar)

        def tarefa():
            try:
                caminho, total = core.mesclarPlanilhas(arquivos, destino)
                nome_arq = caminho.split("/")[-1]
                self._set_log(self._log_mesclar,
                              f"✅ Mesclagem concluída! {total} linhas → '{nome_arq}'.", COR_SUCESSO)
                core.registrarHistorico("Mesclar", f"{len(arquivos)} arquivos → {nome_arq}")
            except Exception as e:
                self._set_log(self._log_mesclar, f"⚠️ {e}", COR_ERRO)
            finally:
                self._parar_progresso(self._pb_mesclar)

        self._rodar_em_thread(tarefa)


# =============================================================================
# SEÇÃO 8 — TELA INSPECIONAR PLANILHA
# =============================================================================

    def tela_visualizar(self):
        self._limpar_tela()
        self._criar_titulo("INSPEÇÃO DE PLANILHA")
        self._criar_seletor_arquivo("path_inspecao", "📂 Escolher Arquivo")

        frame_botoes = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        frame_botoes.pack(pady=8)

        acoes = [
            ("📊 Contar Linhas",   self._exec_contar_linhas),
            ("📋 Cabeçalhos",      self._exec_ler_headers),
            ("📄 Conteúdo",        self._exec_listar_conteudo),
            ("🗂️ Abas",            self._exec_listar_abas),
            ("💾 Exportar CSV",    self._exec_exportar_csv),
        ]
        for i, (texto, cmd) in enumerate(acoes):
            ctk.CTkButton(frame_botoes, text=texto, width=118, command=cmd).grid(
                row=0, column=i, padx=3
            )

        self._pb_inspecao = self._criar_progressbar()

        self.txt_output = ctk.CTkTextbox(
            self.frame_conteudo, width=560, height=220, font=FONTE_MONO
        )
        self.txt_output.pack(pady=10, fill="both", expand=True)

    def _verificar_inspecao(self) -> bool:
        if not self.path_inspecao:
            self._atualizar_output("❌ Selecione um arquivo primeiro.")
            return False
        return True

    def _atualizar_output(self, texto: str):
        def _u():
            self.txt_output.delete("1.0", "end")
            self.txt_output.insert("1.0", texto)
        self.after(0, _u)

    def _exec_contar_linhas(self):
        if not self._verificar_inspecao():
            return
        self._iniciar_progresso(self._pb_inspecao)

        def tarefa():
            try:
                total = core.contarLinhas(self.path_inspecao)
                self._atualizar_output(f"📊 Total de linhas preenchidas: {total}")
                core.registrarHistorico("Inspecionar", f"Contagem: {total} linhas")
            except Exception as e:
                self._atualizar_output(f"❌ {e}")
            finally:
                self._parar_progresso(self._pb_inspecao)

        self._rodar_em_thread(tarefa)

    def _exec_ler_headers(self):
        if not self._verificar_inspecao():
            return
        self._iniciar_progresso(self._pb_inspecao)

        def tarefa():
            try:
                headers = core.lerCabecalhos(self.path_inspecao)
                texto = "📋 Cabeçalhos encontrados:\n\n" + " | ".join(headers)
                self._atualizar_output(texto)
            except Exception as e:
                self._atualizar_output(f"❌ {e}")
            finally:
                self._parar_progresso(self._pb_inspecao)

        self._rodar_em_thread(tarefa)

    def _exec_listar_conteudo(self):
        if not self._verificar_inspecao():
            return
        self._iniciar_progresso(self._pb_inspecao)

        def tarefa():
            try:
                linhas = core.listarConteudo(self.path_inspecao)
                if not linhas:
                    self._atualizar_output("A planilha está vazia.")
                    return
                texto = f"📄 {len(linhas)} linha(s):\n\n"
                for i, linha in enumerate(linhas, 1):
                    valores = [str(v) if v is not None else "—" for v in linha]
                    texto += f"  {i:>4}: {' | '.join(valores)}\n"
                self._atualizar_output(texto)
            except Exception as e:
                self._atualizar_output(f"❌ {e}")
            finally:
                self._parar_progresso(self._pb_inspecao)

        self._rodar_em_thread(tarefa)

    def _exec_listar_abas(self):
        if not self._verificar_inspecao():
            return

        def tarefa():
            try:
                abas = core.listarAbas(self.path_inspecao)
                texto = f"🗂️ {len(abas)} aba(s):\n\n" + "\n".join(f"  • {a}" for a in abas)
                self._atualizar_output(texto)
            except Exception as e:
                self._atualizar_output(f"❌ {e}")

        self._rodar_em_thread(tarefa)

    def _exec_exportar_csv(self):
        if not self._verificar_inspecao():
            return

        destino = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Arquivo CSV", "*.csv")],
            initialfile="exportado",
            title="Salvar CSV como..."
        )
        if not destino:
            return

        self._iniciar_progresso(self._pb_inspecao)
        self._atualizar_output("⏳ Exportando CSV...")

        def tarefa():
            try:
                caminho = core.exportarCSV(self.path_inspecao, destino)
                nome_arq = caminho.split("/")[-1]
                self._atualizar_output(f"✅ CSV exportado: '{nome_arq}'")
                core.registrarHistorico("Exportar CSV", nome_arq)
            except Exception as e:
                self._atualizar_output(f"❌ {e}")
            finally:
                self._parar_progresso(self._pb_inspecao)

        self._rodar_em_thread(tarefa)


# =============================================================================
# SEÇÃO 9 — TELA HISTÓRICO
# =============================================================================

    def tela_historico(self):
        self._limpar_tela()
        self._criar_titulo("HISTÓRICO DE OPERAÇÕES")

        historico = core.obterHistorico()

        frame_topo = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        frame_topo.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(
            frame_topo,
            text=f"{len(historico)} operação(ões) registrada(s).",
            font=FONTE_LABEL
        ).pack(side="left")
        ctk.CTkButton(
            frame_topo, text="🗑️ Limpar Histórico", width=160,
            fg_color="#555", command=self._limpar_historico
        ).pack(side="right")

        # Área de conteúdo do histórico
        scroll = ctk.CTkScrollableFrame(self.frame_conteudo, height=350)
        scroll.pack(fill="both", expand=True)
        scroll.grid_columnconfigure(0, weight=1)

        if not historico:
            ctk.CTkLabel(
                scroll, text="Nenhuma operação registrada ainda.",
                font=FONTE_PEQUENA, text_color="gray"
            ).pack(pady=20)
        else:
            for item in historico:
                card = ctk.CTkFrame(scroll, fg_color="#1e2d27", corner_radius=8)
                card.pack(fill="x", pady=3, padx=4)

                ctk.CTkLabel(
                    card, text=f"  {item['data']}",
                    font=FONTE_PEQUENA, text_color="gray"
                ).grid(row=0, column=0, sticky="w", padx=12, pady=(6, 0))

                ctk.CTkLabel(
                    card,
                    text=f"  [{item['operacao']}]  {item['detalhe']}",
                    font=FONTE_LABEL, anchor="w"
                ).grid(row=1, column=0, sticky="w", padx=12, pady=(0, 6))

    def _limpar_historico(self):
        core.limparHistorico()
        self.tela_historico()  # Recarrega a tela


# =============================================================================
# SEÇÃO 10 — JANELA DE PREVIEW
# =============================================================================

    def _mostrar_preview(self, grade: list[list]):
        """Abre uma janela modal exibindo a grade de dados como tabela."""
        win = ctk.CTkToplevel(self)
        win.title("Preview dos Dados")
        win.geometry("600x380")
        win.grab_set()  # Modal

        ctk.CTkLabel(win, text="📋 Preview — como os dados ficarão na planilha",
                     font=FONTE_TITULO).pack(pady=(14, 8))

        scroll = ctk.CTkScrollableFrame(win)
        scroll.pack(fill="both", expand=True, padx=16, pady=8)

        for r, linha in enumerate(grade):
            for c, valor in enumerate(linha):
                bg = "#1e3a2f" if r == 0 else "#2b2b2b"
                ctk.CTkLabel(
                    scroll, text=str(valor),
                    fg_color=bg, corner_radius=4,
                    padx=10, pady=6, width=120, anchor="w"
                ).grid(row=r, column=c, padx=2, pady=2, sticky="ew")

        ctk.CTkButton(win, text="Fechar", command=win.destroy).pack(pady=10)


# =============================================================================
# SEÇÃO 11 — PONTO DE ENTRADA
# =============================================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()