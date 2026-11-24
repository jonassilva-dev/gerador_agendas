from datetime import datetime, timedelta
import os
from pathlib import Path
import re
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import traceback
from docx import Document


CLIENTES = [
    {
        "arquivo": "01_FUP.txt",
        "titulo": "sindicatos filiados à FUP",
        "titulo_pauta": "PAUTAS DA SEMANA",
        "tags": {"agenda fup", "todas as agendas"},
    },
    {
        "arquivo": "02_Urbanitarios.txt",
        "titulo": "sindicatos de Urbanitários",
        "titulo_pauta": "PAUTAS DA SEMANA",
        "tags": {"agenda urbanitários", "agenda urbanitarios", "todas as agendas"},
    },
    {
        "arquivo": "03_Bancarios_SP.txt",
        "titulo": "sindicatos de Bancários SP",
        "titulo_pauta": "PAUTA DA SEMANA",
        "tags": {"agenda bancários", "agenda bancarios", "todas as agendas"},
    },
    {
        "arquivo": "04_Bancarios_DF.txt",
        "titulo": "sindicatos de Bancários DF",
        "titulo_pauta": "PAUTA DA SEMANA",
        "tags": {"agenda bancários", "agenda bancarios", "todas as agendas"},
    },
    {
        "arquivo": "05_Moedeiros.txt",
        "titulo": "Sind. Moedeiros",
        "titulo_pauta": "PAUTA DA SEMANA",
        "tags": {"agenda bancários", "agenda bancarios", "todas as agendas"},
    },
    {
        "arquivo": "06_Comerciarios.txt",
        "titulo": "sindicatos de Comerciários",
        "titulo_pauta": "PAUTA DA SEMANA",
        "tags": {"todas as agendas"},
    },
]

DIAS_SEMANA = {
    0: "Segunda-feira",
    1: "Terça-feira",
    2: "Quarta-feira",
    3: "Quinta-feira",
    4: "Sexta-feira",
    5: "Sábado",
    6: "Domingo",
}


def formatar_data_curta(dt: datetime) -> str:
    return f"{dt.day}.{dt.strftime('%m.%y')}"


def extrair_periodo_semana(
    nome_arquivo: str, datas_encontradas: list[datetime]
) -> str:
    match = re.search(r"(\d{2})[-./](\d{2})[-./](\d{4})", nome_arquivo)
    if match:
        dia, mes, ano = map(int, match.groups())
        data_ref = datetime(ano, mes, dia)
    elif datas_encontradas:
        data_ref = min(datas_encontradas)
    else:
        return "XX.XX.26 a XX.XX.26"

    segunda = data_ref - timedelta(days=data_ref.weekday())
    sexta = segunda + timedelta(days=4)
    return f"{formatar_data_curta(segunda)} a {formatar_data_curta(sexta)}"


def normalizar_horario(texto_hora: str) -> tuple[int, str]:
    texto = texto_hora.strip()
    texto = re.sub(
        r"^(HORÁRIO|Início previsto|Início)\s*:\s*", "", texto, flags=re.IGNORECASE
    ).strip()
    texto_limpo = texto.split("|")[0].strip()

    match = re.search(r"(\d{1,2})[h:](\d{2})?", texto_limpo, flags=re.IGNORECASE)
    if match:
        hora = int(match.group(1))
        minuto = int(match.group(2)) if match.group(2) else 0
        minutos_totais = hora * 60 + minuto
        hora_fmt = f"{hora}h" if minuto == 0 else f"{hora}h{minuto:02d}"
        return (minutos_totais, hora_fmt)

    texto_cap = (
        texto_limpo[0].upper() + texto_limpo[1:]
        if texto_limpo
        else "Horário a definir"
    )
    return (9999, texto_cap)


def extrair_data(texto: str) -> datetime | None:
    match = re.search(r"(\d{2}/\d{2}/\d{4})", texto)
    if match:
        return datetime.strptime(match.group(1), "%d/%m/%Y")
    return None


def eh_tag_roteamento(linha: str) -> bool:
    l = linha.strip().lower()
    return l.startswith("agenda ") or l == "todas as agendas"


def eh_inicio_item(linha: str) -> bool:
    return bool(re.match(r"^\d+\s*-\s*", linha.strip()))


def limpar_linha_cabecalho(linha: str) -> list[str]:
    s = linha.strip()

    if re.match(r"^(DIA\s*:?\s*)?\d{2}/\d{2}/\d{4}$", s, flags=re.IGNORECASE):
        return []
    if re.match(
        r"^\d{2}/\d{2}/\d{4}\s*(?:às\s*\d{1,2}h(\d{2})?|,\s*.+)$",
        s,
        flags=re.IGNORECASE,
    ):
        return []
    if re.match(r"^(HORÁRIO\s*:?\s*)?\d{1,2}h(\d{2})?$", s, flags=re.IGNORECASE):
        return []
    if re.match(
        r"^Início previsto\s*:\s*\d{1,2}:\d{2}$", s, flags=re.IGNORECASE
    ):
        return []

    if s.lower().startswith(("destinada à", "destinada a")):
        return []

    if s.lower().startswith("início previsto:"):
        return [s]

    s = re.sub(r"^\d{2}/\d{2}/\d{4}\s*[-–]\s*", "", s)

    if s.endswith(" Agendada"):
        s = s[: -len(" Agendada")].strip()

    if s.startswith("Anexo ") and not s.upper().startswith("LOCAL:"):
        return [f"Local: {s}"]

    return [s] if s else []


def parse_docx(caminho_docx: Path) -> list[dict]:
    doc = Document(caminho_docx)
    linhas = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    eventos = []
    casa_atual = "CÂMARA DOS DEPUTADOS"
    evento_atual = None
    item_atual = None

    def salvar_item():
        nonlocal item_atual
        if item_atual and evento_atual is not None:
            evento_atual["itens"].append(item_atual)
            item_atual = None

    def salvar_evento():
        nonlocal evento_atual
        salvar_item()
        if evento_atual is not None:
            eventos.append(evento_atual)
            evento_atual = None

    for linha in linhas:
        if linha.startswith("#"):
            conteudo_hash = linha.lstrip("#").strip()

            if conteudo_hash.lower() in ("câmara", "camara"):
                salvar_evento()
                casa_atual = "CÂMARA DOS DEPUTADOS"
                continue
            elif conteudo_hash.lower() == "senado":
                salvar_evento()
                casa_atual = "SENADO"
                continue

            salvar_evento()
            url = conteudo_hash if conteudo_hash.startswith("http") else ""
            evento_atual = {
                "casa": casa_atual,
                "url": url,
                "cabecalho": [],
                "data_dt": None,
                "hora_ord": 9999,
                "hora_str": "",
                "itens": [],
            }
            continue

        if evento_atual is None:
            evento_atual = {
                "casa": casa_atual,
                "url": "",
                "cabecalho": [],
                "data_dt": None,
                "hora_ord": 9999,
                "hora_str": "",
                "itens": [],
            }

        if eh_tag_roteamento(linha):
            tag_norm = linha.strip().lower()
            if item_atual is not None:
                item_atual["tags"].add(tag_norm)
                salvar_item()
            else:
                evento_atual["itens"].append(
                    {"linhas": [], "tags": {tag_norm}, "sem_numero": True}
                )
            continue

        if eh_inicio_item(linha):
            salvar_item()
            linha_limpa = re.sub(r"^(\d+\s*-)\s+", r"\1 ", linha.strip())
            item_atual = {
                "linhas": [linha_limpa],
                "tags": set(),
                "sem_numero": False,
            }
            continue

        if item_atual is not None:
            item_atual["linhas"].append(linha)
            continue

        if "congressonacional.leg.br" in evento_atual["url"]:
            evento_atual["casa"] = "CONGRESSO NACIONAL"

        data_encontrada = extrair_data(linha)
        if data_encontrada and not evento_atual["data_dt"]:
            evento_atual["data_dt"] = data_encontrada

        match_data_virgula = re.match(
            r"^\d{2}/\d{2}/\d{4}\s*,\s*(.+)$", linha.strip()
        )
        if match_data_virgula:
            evento_atual["hora_ord"], evento_atual["hora_str"] = normalizar_horario(
                match_data_virgula.group(1)
            )
        elif any(
            k in linha.upper()
            for k in ["HORÁRIO:", "INÍCIO PREVISTO:", "INÍCIO ", " ÀS "]
        ) or re.match(r"^\d{1,2}h", linha, flags=re.IGNORECASE):
            if "às" in linha.lower():
                parte_hora = linha.split("às", 1)[1].strip()
                evento_atual["hora_ord"], evento_atual["hora_str"] = (
                    normalizar_horario(parte_hora)
                )
            elif any(
                k in linha.upper()
                for k in ["HORÁRIO:", "INÍCIO PREVISTO:", "INÍCIO"]
            ):
                evento_atual["hora_ord"], evento_atual["hora_str"] = (
                    normalizar_horario(linha)
                )
            elif re.match(r"^\d{1,2}h(\d{2})?$", linha.strip(), flags=re.IGNORECASE):
                evento_atual["hora_ord"], evento_atual["hora_str"] = (
                    normalizar_horario(linha)
                )

        if linha.lower() in (
            "comunicado",
            "sessão anterioroutra sessão próxima sessão",
        ):
            continue

        if linha.startswith("http"):
            continue

        evento_atual["cabecalho"].extend(limpar_linha_cabecalho(linha))

    salvar_evento()
    return eventos


def formatar_linhas_item(linhas: list[str], eh_urbanitarios: bool) -> list[str]:
    blocos_item = []
    for idx, l_item in enumerate(linhas):
        if idx == 0:
            blocos_item.append(l_item)
        elif l_item.upper().startswith(
            ("RELATOR", "PARECER", "EXPLICAÇÃO", "RETIRADO", "(AVULSO", "CONVIDADOS")
        ) or l_item.startswith(("Relator", "Relatório")):
            blocos_item.append(l_item)
        elif (
            idx == 1
            and re.match(r"^\d+\s*-\s*PL\s+\d+/\d{4}$", blocos_item[-1].strip())
            and eh_urbanitarios
        ):
            blocos_item[-1] = f"{blocos_item[-1].strip()} - {l_item.strip()}"
        else:
            blocos_item[-1] = f"{blocos_item[-1]}\n{l_item}"
    return blocos_item


def gerar_mensagem_cliente(
    eventos: list[dict], cliente: dict, periodo_str: str
) -> str:
    tags_cliente = cliente["tags"]
    eh_urbanitarios = "urbanitários" in cliente["titulo"].lower()
    blocos_filtrados = []

    for ev in eventos:
        itens_validos = [
            it for it in ev["itens"] if it["tags"].intersection(tags_cliente)
        ]
        if itens_validos:
            blocos_filtrados.append((ev, itens_validos))

    blocos_filtrados.sort(
        key=lambda x: (x[0]["data_dt"] or datetime.max, x[0]["hora_ord"])
    )

    saida = [
        f"*Advocacia Garcez - Agenda Legislativa da semana de {periodo_str} para {cliente['titulo']}*",
        "",
        cliente["titulo_pauta"],
    ]

    if not blocos_filtrados:
        saida.append(
            "\nNenhuma pauta específica registrada para esta categoria nesta semana."
        )
        return "\n".join(saida)

    for ev, itens in blocos_filtrados:
        if ev["data_dt"]:
            dia_sem = DIAS_SEMANA[ev["data_dt"].weekday()]
            data_fmt = ev["data_dt"].strftime("%d/%m")
            saida.extend(["", f"*{dia_sem} | {data_fmt}*"])

        if ev["hora_str"]:
            saida.extend(["", f"*{ev['hora_str']}*"])

        saida.extend(["", ev["casa"]])

        for linha_cab in ev["cabecalho"]:
            saida.extend(["", linha_cab])

        for it in itens:
            if it.get("sem_numero"):
                continue
            for bloco_texto in formatar_linhas_item(it["linhas"], eh_urbanitarios):
                saida.extend(["", bloco_texto])

        if ev["url"]:
            saida.extend(["", f"Mais informações: {ev['url']}"])

    return "\n".join(saida)


class AppAgendasLegislativas:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Advocacia Garcez - Gerador de Agendas Legislativas")
        self.root.geometry("680x480")
        self.root.resizable(False, False)
        self.root.configure(bg="#F4F6F9")

        self.caminho_docx = tk.StringVar()
        self.pasta_destino = tk.StringVar(value=str(Path.home() / "Downloads"))
        self.ultima_pasta_gerada = None

        if len(sys.argv) > 1 and Path(sys.argv[1]).suffix.lower() == ".docx":
            arq_inicial = Path(sys.argv[1]).resolve()
            if arq_inicial.exists():
                self.caminho_docx.set(str(arq_inicial))
                self.pasta_destino.set(str(arq_inicial.parent))

        self._construir_interface()

    def _construir_interface(self):
        header = tk.Frame(self.root, bg="#1B365D", height=64)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header,
            text="Advocacia Garcez  |  Gerador de Agendas Legislativas",
            font=("Segoe UI", 13, "bold"),
            fg="#FFFFFF",
            bg="#1B365D",
        ).pack(side="left", padx=20)

        body = tk.Frame(self.root, bg="#F4F6F9", padx=20, pady=18)
        body.pack(fill="both", expand=True)

        tk.Label(
            body,
            text="1. Selecione o documento Word (.docx) da semana:",
            font=("Segoe UI", 10, "bold"),
            fg="#1E293B",
            bg="#F4F6F9",
        ).pack(anchor="w")

        frame_doc = tk.Frame(body, bg="#F4F6F9", pady=5)
        frame_doc.pack(fill="x")

        entry_doc = tk.Entry(
            frame_doc,
            textvariable=self.caminho_docx,
            font=("Segoe UI", 9),
            state="readonly",
            readonlybackground="#FFFFFF",
            relief="solid",
            bd=1,
        )
        entry_doc.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 8))

        tk.Button(
            frame_doc,
            text="Selecionar Arquivo...",
            font=("Segoe UI", 9, "bold"),
            bg="#E2E8F0",
            fg="#1E293B",
            activebackground="#CBD5E1",
            relief="flat",
            cursor="hand2",
            padx=14,
            pady=5,
            command=self.selecionar_arquivo,
        ).pack(side="right")

        tk.Label(
            body,
            text="2. Escolha onde salvar a pasta com os arquivos prontos:",
            font=("Segoe UI", 10, "bold"),
            fg="#1E293B",
            bg="#F4F6F9",
        ).pack(anchor="w", pady=(12, 0))

        frame_pasta = tk.Frame(body, bg="#F4F6F9", pady=5)
        frame_pasta.pack(fill="x")

        entry_pasta = tk.Entry(
            frame_pasta,
            textvariable=self.pasta_destino,
            font=("Segoe UI", 9),
            state="readonly",
            readonlybackground="#FFFFFF",
            relief="solid",
            bd=1,
        )
        entry_pasta.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 8))

        tk.Button(
            frame_pasta,
            text="Escolher Pasta...",
            font=("Segoe UI", 9, "bold"),
            bg="#E2E8F0",
            fg="#1E293B",
            activebackground="#CBD5E1",
            relief="flat",
            cursor="hand2",
            padx=14,
            pady=5,
            command=self.selecionar_pasta,
        ).pack(side="right")

        frame_acoes = tk.Frame(body, bg="#F4F6F9", pady=14)
        frame_acoes.pack(fill="x")

        self.btn_gerar = tk.Button(
            frame_acoes,
            text="Gerar Agendas Formatadas",
            font=("Segoe UI", 10, "bold"),
            bg="#1B365D",
            fg="#FFFFFF",
            activebackground="#122540",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=20,
            pady=8,
            command=self.processar,
        )
        self.btn_gerar.pack(side="left")

        self.btn_abrir_pasta = tk.Button(
            frame_acoes,
            text="Abrir Pasta de Saída",
            font=("Segoe UI", 9, "bold"),
            bg="#10B981",
            fg="#FFFFFF",
            activebackground="#059669",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=16,
            pady=8,
            state="disabled",
            command=self.abrir_pasta_saida,
        )
        self.btn_abrir_pasta.pack(side="right")

        tk.Label(
            body,
            text="Status do processamento:",
            font=("Segoe UI", 9, "bold"),
            fg="#475569",
            bg="#F4F6F9",
        ).pack(anchor="w")

        self.log_box = scrolledtext.ScrolledText(
            body,
            height=9,
            font=("Consolas", 9),
            bg="#FFFFFF",
            fg="#1E293B",
            relief="solid",
            bd=1,
            state="disabled",
        )
        self.log_box.pack(fill="both", expand=True, pady=(4, 0))
        self._escrever_log(
            "Pronto. Selecione o documento Word e a pasta de destino para iniciar."
        )

    def _escrever_log(self, texto: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", texto + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.root.update_idletasks()

    def _limpar_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione o documento Word da Agenda Legislativa",
            filetypes=[("Documentos Word", "*.docx"), ("Todos os arquivos", "*.*")],
        )
        if arquivo:
            self.caminho_docx.set(arquivo)
            if self.pasta_destino.get() == str(Path.home() / "Downloads"):
                self.pasta_destino.set(str(Path(arquivo).parent))

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(
            title="Escolha a pasta onde deseja salvar as agendas"
        )
        if pasta:
            self.pasta_destino.set(pasta)

    def abrir_pasta_saida(self):
        if self.ultima_pasta_gerada and self.ultima_pasta_gerada.exists():
            if sys.platform == "win32":
                os.startfile(self.ultima_pasta_gerada)
            elif sys.platform == "darwin":
                subprocess.run(["open", str(self.ultima_pasta_gerada)], check=False)
            else:
                subprocess.run(["xdg-open", str(self.ultima_pasta_gerada)], check=False)

    def processar(self):
        doc_str = self.caminho_docx.get().strip()
        dest_str = self.pasta_destino.get().strip()

        if not doc_str:
            messagebox.showwarning(
                "Aviso", "Por favor, selecione o documento Word (.docx) primeiro."
            )
            return

        arquivo_alvo = Path(doc_str)
        if not arquivo_alvo.exists():
            messagebox.showerror(
                "Erro", f"O arquivo selecionado não foi encontrado:\n{arquivo_alvo}"
            )
            return

        if not dest_str:
            messagebox.showwarning(
                "Aviso", "Por favor, escolha uma pasta de destino para salvar."
            )
            return

        pasta_base = Path(dest_str)
        if not pasta_base.exists():
            messagebox.showerror(
                "Erro", f"A pasta de destino não existe:\n{pasta_base}"
            )
            return

        self._limpar_log()
        self.btn_gerar.configure(state="disabled")
        self.btn_abrir_pasta.configure(state="disabled")

        try:
            self._escrever_log(f"Lendo arquivo: {arquivo_alvo.name}...")
            eventos = parse_docx(arquivo_alvo)
            datas = [ev["data_dt"] for ev in eventos if ev["data_dt"]]
            periodo_str = extrair_periodo_semana(arquivo_alvo.name, datas)

            pasta_saida = (
                pasta_base / f"Agendas_Prontas_{periodo_str.replace('.', '-')}"
            )
            pasta_saida.mkdir(parents=True, exist_ok=True)
            self.ultima_pasta_gerada = pasta_saida

            self._escrever_log(f"Período identificado: {periodo_str}")
            self._escrever_log(f"Eventos encontrados no Word: {len(eventos)}\n")

            for cliente in CLIENTES:
                texto_wpp = gerar_mensagem_cliente(eventos, cliente, periodo_str)
                caminho_txt = pasta_saida / cliente["arquivo"]
                caminho_txt.write_text(texto_wpp, encoding="utf-8")
                self._escrever_log(f"[OK] {cliente['arquivo']} -> {cliente['titulo']}")

            self._escrever_log("\nConcluído com sucesso!")
            self._escrever_log(f"Arquivos salvos em: {pasta_saida}")
            self.btn_abrir_pasta.configure(state="normal")
            messagebox.showinfo(
                "Concluído",
                f"Agendas geradas com sucesso!\n\nPasta: {pasta_saida.name}",
            )
        except Exception as e:
            self._escrever_log("\n[ERRO] Falha ao processar o documento:")
            self._escrever_log(traceback.format_exc())
            messagebox.showerror(
                "Erro no Processamento",
                f"Ocorreu um erro ao processar o arquivo:\n{e}",
            )
        finally:
            self.btn_gerar.configure(state="normal")


def main():
    root = tk.Tk()
    AppAgendasLegislativas(root)
    root.mainloop()


if __name__ == "__main__":
    main()