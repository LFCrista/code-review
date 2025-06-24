import os
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk

import config
import auth
from job_queue import Job
from pdf_processor import process_pdfs


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Automatizador de PDFs – Fila Dinâmica")
        self.geometry("800x640")

        # ---- Estruturas de dados ----
        self.jobs: list[Job] = []
        self.jobs_lock = threading.Lock()        # protege self.jobs
        self.worker_thread: threading.Thread | None = None
        self.current_pdfs: list[str] = []

        # ---- GUI ----
        self._build_gui()

    # ----------------------------------------------------------------
    # GUI builders
    # ----------------------------------------------------------------
    def _build_gui(self):
        self._build_topbar()
        self._build_inputs()
        self._build_job_queue()
        self._build_log()

    def _build_topbar(self):
        top = tk.Frame(self)
        top.pack(pady=8, padx=10, anchor="w")

        def abrir_debug():
            subprocess.Popen(
                [
                    config.CHROME_PATH,
                    f"--remote-debugging-port={config.CHROME_REMOTE_DEBUGGING_PORT}",
                    f"--user-data-dir={config.CHROME_USER_DATA_DIR}",
                ]
            )

        tk.Button(top, text="Abrir Chrome Debug", command=abrir_debug).pack(
            side=tk.LEFT, padx=5
        )

        tk.Button(
            top,
            text="Trocar credenciais",
            command=lambda: [
                os.remove(f)
                for f in ("token.json", "credentials.json")
                if os.path.exists(f)
            ],
        ).pack(side=tk.LEFT, padx=5)

    def _build_inputs(self):
        frm = tk.LabelFrame(self, text="Novo Job")
        frm.pack(fill="x", padx=10, pady=5, ipadx=5, ipady=5)

        tk.Label(frm, text="Link GPT:").grid(row=0, column=0, sticky="e")
        self.entry_gpt = tk.Entry(frm, width=80)
        self.entry_gpt.grid(row=0, column=1, sticky="w", padx=4, pady=2)

        tk.Label(frm, text="Link Docs:").grid(row=1, column=0, sticky="e")
        self.entry_docs = tk.Entry(frm, width=80)
        self.entry_docs.grid(row=1, column=1, sticky="w", padx=4, pady=2)

        tk.Label(frm, text="Prompt:").grid(row=2, column=0, sticky="e")
        self.entry_prompt = tk.Entry(frm, width=80)
        self.entry_prompt.grid(row=2, column=1, sticky="w", padx=4, pady=2)

        tk.Label(frm, text="Máx. PDFs por aba:").grid(row=3, column=0, sticky="e")
        self.var_max = tk.StringVar(value="30")
        tk.Spinbox(
            frm, from_=1, to=100, width=5, textvariable=self.var_max, increment=1
        ).grid(row=3, column=1, sticky="w", padx=4, pady=2)

        self.lbl_pdfs = tk.Label(frm, text="Nenhum PDF selecionado")
        self.lbl_pdfs.grid(row=4, column=1, sticky="w", padx=4, pady=2)

        tk.Button(frm, text="Selecionar PDFs", command=self._select_pdfs).grid(
            row=4, column=0, pady=2
        )

        tk.Button(frm, text="Adicionar à Fila", command=self._add_job).grid(
            row=5, column=1, sticky="e", pady=4, padx=4
        )

    def _build_job_queue(self):
        frm = tk.LabelFrame(self, text="Fila de Jobs (dinâmica)")
        frm.pack(fill="both", expand=False, padx=10, pady=5, ipadx=5, ipady=5)

        self.tree = ttk.Treeview(
            frm,
            columns=("gpt", "docs", "pdfs", "status"),
            show="headings",
            height=6,
            selectmode="browse",
        )
        self.tree.heading("gpt", text="GPT")
        self.tree.heading("docs", text="Docs")
        self.tree.heading("pdfs", text="PDFs")
        self.tree.heading("status", text="Status")
        self.tree.column("gpt", width=180)
        self.tree.column("docs", width=180)
        self.tree.column("pdfs", width=60, anchor="center")
        self.tree.column("status", width=80, anchor="center")
        self.tree.pack(side=tk.LEFT, fill="both", expand=True)

        scrollbar = ttk.Scrollbar(frm, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill="y")

        btns = tk.Frame(self)
        btns.pack(pady=5)

        self.btn_start_queue = tk.Button(
            btns, text="Iniciar Fila", command=self._start_queue
        )
        self.btn_start_queue.pack(side=tk.LEFT, padx=10)

        tk.Button(
            btns, text="Remover Selecionado", command=self._remove_selected
        ).pack(side=tk.LEFT, padx=10)

    def _build_log(self):
        self.texto_log = scrolledtext.ScrolledText(self, width=100, height=15)
        self.texto_log.pack(padx=10, pady=10, fill="both", expand=True)

    # ----------------------------------------------------------------
    # Handlers
    # ----------------------------------------------------------------
    def _select_pdfs(self):
        pdfs = filedialog.askopenfilenames(
            title="Selecione os PDFs", filetypes=[("PDF", "*.pdf")]
        )
        if pdfs:
            self.current_pdfs = list(pdfs)
            self.lbl_pdfs.config(text=f"{len(self.current_pdfs)} PDF(s) selecionado(s)")
        else:
            self.current_pdfs = []
            self.lbl_pdfs.config(text="Nenhum PDF selecionado")

    def _add_job(self):
        if not self.current_pdfs:
            messagebox.showwarning("Aviso", "Selecione pelo menos um PDF.")
            return

        link_gpt = self.entry_gpt.get().strip()
        link_doc = self.entry_docs.get().strip()
        prompt = self.entry_prompt.get().strip()
        doc_id = auth.extrair_document_id(link_doc)

        if not all((link_gpt, doc_id, prompt)):
            messagebox.showerror(
                "Erro",
                "Preencha Link GPT, Link Docs (URL válida) e Prompt antes de adicionar.",
            )
            return

        try:
            max_per_tab = int(self.var_max.get())
            if not 1 <= max_per_tab <= 100:
                max_per_tab = 30
        except Exception:
            max_per_tab = 30

        job = Job(
            pdfs=tuple(self.current_pdfs),
            link_gpt=link_gpt,
            doc_id=doc_id,
            prompt=prompt,
            max_per_tab=max_per_tab,
        )

        with self.jobs_lock:
            self.jobs.append(job)
            idx = len(self.jobs) - 1

        # adiciona/atualiza lista visual
        self.tree.insert(
            "",
            "end",
            iid=idx,
            values=(link_gpt[:40], link_doc[:40], len(job.pdfs), "Pendente"),
        )

        # limpa seleção atual
        self.current_pdfs = []
        self.lbl_pdfs.config(text="Nenhum PDF selecionado")

    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])

        with self.jobs_lock:
            if idx < len(self.jobs):
                del self.jobs[idx]

        self.tree.delete(sel)

        # reajusta iids na tree
        for new_i, item in enumerate(self.tree.get_children()):
            self.tree.item(item, iid=new_i)

    # ----------------------------------------------------------------
    # Execução da fila – dinâmica
    # ----------------------------------------------------------------
    def _start_queue(self):
        # evita múltiplos workers
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Fila", "A fila já está em execução.")
            return

        def worker():
            idx = 0
            while True:
                with self.jobs_lock:
                    if idx >= len(self.jobs):
                        break
                    job = self.jobs[idx]
                    # marca como "Em execução" na tree
                    self.tree.set(idx, column="status", value="Rodando")

                self._log(f"=== Job {idx + 1} ===\n")
                try:
                    process_pdfs(
                        job.pdfs,
                        job.link_gpt,
                        job.doc_id,
                        job.prompt,
                        job.max_per_tab,
                        self.texto_log,
                        self,
                    )
                    with self.jobs_lock:
                        self.tree.set(idx, column="status", value="Concluído")
                except Exception as e:
                    with self.jobs_lock:
                        self.tree.set(idx, column="status", value="Erro")
                    messagebox.showerror("Erro no Job", str(e))

                idx += 1

            self._log("Fila finalizada.\n")

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def _log(self, txt: str):
        self.texto_log.insert("end", txt)
        self.texto_log.see("end")


if __name__ == "__main__":
    App().mainloop()
