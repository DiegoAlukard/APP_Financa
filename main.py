import ssl
ssl._create_default_https_context = ssl._create_unverified_context

import flet as ft
import sqlite3
import shutil
import json
import openpyxl
from pathlib import Path
from datetime import datetime

# Garante os caminhos locais na mesma pasta
BASE_DIR = Path(__file__).resolve().parent
DB_NAME = str(BASE_DIR / "financas_local.db")
BACKUP_DIR = BASE_DIR / "backups"
BACKUP_DIR.mkdir(exist_ok=True)
CONFIG_FILE = BASE_DIR / "config_app.json"


class BancoDados:
    def __init__(self, db_path=DB_NAME):
        self.db_path = db_path
        self.conectar()
        self.criar_tabelas()

    def conectar(self):
        self.conn = sqlite3.connect(self.db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row

    def criar_tabelas(self):
        cur = self.conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS lancamentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ano_mes TEXT NOT NULL,
                tipo TEXT NOT NULL,
                descricao TEXT NOT NULL,
                valor REAL NOT NULL,
                fixo INTEGER DEFAULT 0
            )
        """)
        self.conn.commit()

        colunas_existentes = [col["name"] for col in cur.execute("PRAGMA table_info(lancamentos)").fetchall()]
        if "forma_pagamento" not in colunas_existentes:
            cur.execute("ALTER TABLE lancamentos ADD COLUMN forma_pagamento TEXT DEFAULT ''")
        if "pago" not in colunas_existentes:
            cur.execute("ALTER TABLE lancamentos ADD COLUMN pago INTEGER DEFAULT 0")
        self.conn.commit()

    def adicionar_lancamento(self, ano_mes, tipo, descricao, valor, fixo=0, forma_pagamento="", pago=0):
        cur = self.conn.cursor()
        cur.execute("""
            INSERT INTO lancamentos (ano_mes, tipo, descricao, valor, fixo, forma_pagamento, pago)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (ano_mes, tipo, descricao, float(valor), int(fixo), forma_pagamento, int(pago)))
        self.conn.commit()

    def atualizar_lancamento(self, item_id, tipo, descricao, valor, forma_pagamento, pago):
        cur = self.conn.cursor()
        cur.execute("""
            UPDATE lancamentos
            SET tipo = ?, descricao = ?, valor = ?, forma_pagamento = ?, pago = ?
            WHERE id = ?
        """, (tipo, descricao, float(valor), forma_pagamento, int(pago), item_id))
        self.conn.commit()

    def excluir_lancamento(self, item_id):
        cur = self.conn.cursor()
        cur.execute("DELETE FROM lancamentos WHERE id = ?", (item_id,))
        self.conn.commit()

    def buscar_por_mes(self, ano_mes):
        cur = self.conn.cursor()
        cur.execute("""
            SELECT * FROM lancamentos
            WHERE ano_mes = ?
            ORDER BY tipo DESC, id DESC
        """, (ano_mes,))
        return [dict(row) for row in cur.fetchall()]

    def listar_todos_meses_com_dados(self):
        cur = self.conn.cursor()
        cur.execute("SELECT DISTINCT ano_mes FROM lancamentos ORDER BY ano_mes DESC")
        return [row["ano_mes"] for row in cur.fetchall()]

    def resumo_mes(self, ano_mes):
        itens = self.buscar_por_mes(ano_mes)
        receitas = sum(i["valor"] for i in itens if i["tipo"] == "Receita")
        despesas = sum(i["valor"] for i in itens if i["tipo"] == "Despesa")
        saldo = receitas - despesas
        return receitas, despesas, saldo

    def exportar_excel(self, caminho_xlsx):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM lancamentos ORDER BY ano_mes, id")
        rows = [dict(r) for r in cur.fetchall()]
        if not rows:
            return False, "Nenhum dado para exportar."

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Lançamentos"

        headers = ["ID", "Ano-Mês", "Tipo", "Descrição", "Valor", "Fixo", "Forma de Pagamento", "Pago"]
        ws.append(headers)

        for r in rows:
            linha = [
                r["id"],
                r["ano_mes"],
                r["tipo"],
                r["descricao"],
                r["valor"],
                r["fixo"],
                r.get("forma_pagamento", ""),
                "Sim" if r.get("pago", 0) else "Não"
            ]
            ws.append(linha)

        wb.save(caminho_xlsx)
        return True, f"Exportado para {caminho_xlsx}"

    def importar_excel(self, caminho_xlsx):
        wb = openpyxl.load_workbook(caminho_xlsx, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            return False, "Planilha vazia ou sem cabeçalho."

        headers = [str(h).strip().lower() for h in rows[0]]
        importados = 0

        for row in rows[1:]:
            if not row or not row[1]:
                continue
            dados_linha = {headers[i]: row[i] for i in range(min(len(headers), len(row))) if headers[i]}
            
            ano_mes = str(dados_linha.get("ano-mês", dados_linha.get("ano_mes", datetime.now().strftime("%Y-%m"))))
            tipo = str(dados_linha.get("tipo", "Receita"))
            desc = str(dados_linha.get("descrição", dados_linha.get("descricao", "Importado")))
            try:
                val = float(str(dados_linha.get("valor", 0)).replace(",", "."))
            except ValueError:
                val = 0.0
            fixo = int(dados_linha.get("fixo", 0) or 0)
            forma_pgto = str(dados_linha.get("forma de pagamento", dados_linha.get("forma_pagamento", "")) or "")
            pago_val = dados_linha.get("pago", 0)
            pago = 1 if str(pago_val).lower() in ["1", "true", "sim", "s", "pago"] else 0

            self.adicionar_lancamento(ano_mes, tipo, desc, val, fixo, forma_pgto, pago)
            importados += 1

        return True, f"{importados} registros importados do Excel."

    def exportar_db(self, destino):
        self.conn.commit()
        shutil.copy2(self.db_path, destino)
        return True, f"Banco salvo em: {destino}"

    def importar_db(self, origem):
        self.conn.close()
        shutil.copy2(origem, self.db_path)
        self.conectar()
        return True, "Banco de dados substituído com sucesso."


def carregar_config_zoom():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("escala_zoom", 1.5)  # Padrão agora é 1.5 (150%)
        except:
            pass
    return 1.5  # Padrão inicial 150%


def salvar_config_zoom(escala):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"escala_zoom": escala}, f)
    except:
        pass


def gerar_lista_meses():
    ano_atual = datetime.now().year
    meses = []
    for ano in [ano_atual - 1, ano_atual, ano_atual + 1]:
        for mes in range(1, 13):
            meses.append(f"{ano}-{mes:02d}")
    return meses


def formatar_mes_br(ano_mes):
    try:
        a, m = ano_mes.split("-")
        return f"{m}/{a}"
    except:
        return ano_mes


def somar_meses(ano_mes_inicial, qtd_meses):
    ano, mes = map(int, ano_mes_inicial.split("-"))
    lista = []
    for i in range(qtd_meses):
        m = mes + i
        a = ano + (m - 1) // 12
        m_final = ((m - 1) % 12) + 1
        lista.append(f"{a}-{m_final:02d}")
    return lista


def rotacionar_backups():
    backups = sorted(BACKUP_DIR.glob("backup_auto_*.db"), key=lambda p: p.stat().st_mtime, reverse=True)
    for velho in backups[9:]:
        try:
            velho.unlink()
        except:
            pass


def fazer_backup_automatico(db_path):
    try:
        rotacionar_backups()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        destino = BACKUP_DIR / f"backup_auto_{timestamp}.db"
        shutil.copy2(db_path, destino)
    except Exception as e:
        print(f"Erro no backup automático: {e}")


def main(page: ft.Page):
    page.title = "Controle Financeiro"
    page.theme_mode = ft.ThemeMode.DARK
    page.padding = 6
    page.window.width = 430
    page.window.height = 800

    db = BancoDados()
    zoom_atual = carregar_config_zoom()

    meses_com_dados = db.listar_todos_meses_com_dados()
    mes_atual = meses_com_dados[0] if meses_com_dados else datetime.now().strftime("%Y-%m")
    aba_ativa = "principal"
    item_selecionado = None

    estado_ordenacao = {"coluna": None, "ascendente": True}

    def f_size(base):
        return int(base * zoom_atual)

    def popular_dropdown_meses():
        meses_disponiveis = set(gerar_lista_meses() + db.listar_todos_meses_com_dados())
        sorted_meses = sorted(list(meses_disponiveis), reverse=True)
        dd_mes.options = [ft.dropdown.Option(key=m, text=formatar_mes_br(m)) for m in sorted_meses]

    dd_mes = ft.Dropdown(
        value=mes_atual,
        width=145,
        dense=True,
        text_size=f_size(13),
        content_padding=6,
    )
    popular_dropdown_meses()

    txt_receitas = ft.Text("Receitas: R$ 0,00", color=ft.Colors.GREEN_400, size=f_size(12))
    txt_despesas = ft.Text("Despesas: R$ 0,00", color=ft.Colors.RED_400, size=f_size(12))
    txt_saldo = ft.Text("Saldo: R$ 0,00", size=f_size(13), weight=ft.FontWeight.BOLD)

    # Painel de detalhes do item selecionado com mesmo alinhamento
    txt_detalhes_info = ft.Text("Selecione um item abaixo para ver detalhes.", size=f_size(11), color=ft.Colors.GREY_400)
    
    container_detalhes_direita = ft.Container(
        content=ft.Column([
            ft.Text("Detalhes do Item:", size=f_size(12), weight=ft.FontWeight.BOLD, color=ft.Colors.CYAN_300),
            txt_detalhes_info
        ], spacing=4, alignment=ft.MainAxisAlignment.CENTER),
        padding=8,
        border_radius=6,
        bgcolor=ft.Colors.with_opacity(0.06, ft.Colors.WHITE),
        expand=True,
    )

    txt_aviso = ft.Text(f"Sistema pronto. Zoom padrão: {int(zoom_atual * 100)}%.", size=f_size(11), color=ft.Colors.AMBER_300)

    def avisar(msg, erro=False):
        txt_aviso.value = msg
        txt_aviso.color = ft.Colors.RED_300 if erro else ft.Colors.AMBER_300
        page.update()

    async def exportar_excel_click(e):
        caminho = await ft.FilePicker().save_file(file_name="financas_export.xlsx", allowed_extensions=["xlsx"])
        if caminho:
            ok, msg = db.exportar_excel(caminho)
            avisar(msg, not ok)

    async def importar_excel_click(e):
        arquivos = await ft.FilePicker().pick_files(allowed_extensions=["xlsx"])
        if arquivos:
            try:
                ok, msg = db.importar_excel(arquivos[0].path)
                popular_dropdown_meses()
                atualizar_tela()
                avisar(msg, not ok)
            except Exception as ex:
                avisar(f"Erro ao importar Excel: {ex}", True)

    async def exportar_db_click(e):
        caminho = await ft.FilePicker().save_file(file_name="financas_backup.db", allowed_extensions=["db"])
        if caminho:
            ok, msg = db.exportar_db(caminho)
            avisar(msg, not ok)

    async def importar_db_click(e):
        arquivos = await ft.FilePicker().pick_files(allowed_extensions=["db"])
        if arquivos:
            try:
                ok, msg = db.importar_db(arquivos[0].path)
                popular_dropdown_meses()
                atualizar_tela()
                avisar(msg, not ok)
            except Exception as ex:
                avisar(f"Erro ao importar DB: {ex}", True)

    def atualizar_resumo():
        mes = dd_mes.value or mes_atual
        rec, desp, sal = db.resumo_mes(mes)
        txt_receitas.value = f"Receitas: R$ {rec:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        txt_despesas.value = f"Despesas: R$ {desp:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        txt_saldo.value = f"Saldo: R$ {sal:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        txt_saldo.color = ft.Colors.CYAN_300 if sal >= 0 else ft.Colors.RED_400

    corpo_central = ft.Container(
        expand=True,
        border=ft.Border.all(1, ft.Colors.OUTLINE),
        border_radius=6,
        padding=6,
    )

    def confirmar_exclusao(item_id):
        nonlocal item_selecionado
        def apagar(e):
            nonlocal item_selecionado
            db.excluir_lancamento(item_id)
            if item_selecionado and item_selecionado.get("id") == item_id:
                item_selecionado = None
                txt_detalhes_info.value = "Selecione um item abaixo para ver detalhes."
                txt_detalhes_info.color = ft.Colors.GREY_400
            page.dialog.open = False
            popular_dropdown_meses()
            atualizar_tela()
            avisar("Registro excluído com sucesso.")

        dlg = ft.AlertDialog(
            title=ft.Text("Confirmar Exclusão"),
            content=ft.Text("Deseja realmente excluir este lançamento?"),
            actions=[
                ft.TextButton("Cancelar", on_click=lambda e: setattr(page.dialog, 'open', False) or page.update()),
                ft.TextButton("Excluir", on_click=apagar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        page.dialog = dlg
        dlg.open = True
        page.update()

    def construir_tela_editar(item):
        dd_tipo = ft.Dropdown(
            label="Tipo",
            options=[ft.dropdown.Option("Receita"), ft.dropdown.Option("Despesa")],
            value=item["tipo"],
            dense=True,
        )
        tf_desc = ft.TextField(label="Descrição", value=item["descricao"], dense=True)
        tf_valor = ft.TextField(
            label="Valor",
            value=str(item["valor"]),
            keyboard_type=ft.KeyboardType.NUMBER,
            dense=True,
        )

        dd_forma_pagamento = ft.Dropdown(
            label="Forma de Pagamento",
            options=[
                ft.dropdown.Option("Crédito"),
                ft.dropdown.Option("Débito"),
                ft.dropdown.Option("Vale"),
                ft.dropdown.Option("Dinheiro"),
            ],
            value=item["forma_pagamento"] if item["forma_pagamento"] else "Crédito",
            dense=True,
        )
        cb_pago = ft.Checkbox(label="Pago", value=bool(item["pago"]))

        container_despesa_campos = ft.Column([
            dd_forma_pagamento,
            cb_pago,
        ], spacing=4)
        container_despesa_campos.visible = (item["tipo"] == "Despesa")

        def on_tipo_change(e):
            container_despesa_campos.visible = (dd_tipo.value == "Despesa")
            page.update()

        dd_tipo.on_change = on_tipo_change

        def salvar_edicao(e):
            try:
                desc = (tf_desc.value or "").strip()
                valor_str = (tf_valor.value or "").strip().replace(",", ".")
                if not desc or not valor_str:
                    avisar("Preencha descrição e valor.", True)
                    return
                valor = float(valor_str)

                tipo_val = dd_tipo.value
                forma_pgto = dd_forma_pagamento.value if tipo_val == "Despesa" else ""
                pago_val = 1 if (tipo_val == "Despesa" and cb_pago.value) else 0

                db.atualizar_lancamento(item["id"], tipo_val, desc, valor, forma_pgto, pago_val)
                popular_dropdown_meses()
                mudar_aba("principal")
                avisar("Lançamento atualizado com sucesso!")
            except Exception as ex:
                avisar(f"Erro ao atualizar: {ex}", True)

        corpo_central.content = ft.ListView(
            controls=[
                ft.Row([
                    ft.IconButton(icon=ft.Icons.ARROW_BACK, on_click=lambda _: mudar_aba("principal")),
                    ft.Text("Editar Lançamento", size=f_size(15), weight=ft.FontWeight.BOLD),
                ], spacing=4),
                dd_tipo,
                tf_desc,
                tf_valor,
                container_despesa_campos,
                ft.Button("Salvar Alterações", icon=ft.Icons.SAVE, on_click=salvar_edicao),
            ],
            expand=True,
            spacing=6,
        )
        page.update()

    def construir_tela_principal():
        nonlocal item_selecionado
        mes = dd_mes.value or mes_atual
        itens = db.buscar_por_mes(mes)

        if not itens:
            return ft.Column(
                [
                    ft.Container(
                        content=ft.Text(f"Nenhum lançamento em {formatar_mes_br(mes)}.", color=ft.Colors.GREY_400, size=f_size(12)),
                        alignment=ft.Alignment.CENTER,
                        expand=True,
                    )
                ],
                expand=True,
                alignment=ft.MainAxisAlignment.CENTER,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            )

        coluna_atual = estado_ordenacao["coluna"]
        asc = estado_ordenacao["ascendente"]

        if coluna_atual == 0:
            itens.sort(key=lambda x: x["tipo"], reverse=not asc)
        elif coluna_atual == 1:
            itens.sort(key=lambda x: x["descricao"].lower(), reverse=not asc)
        elif coluna_atual == 2:
            itens.sort(key=lambda x: x["valor"], reverse=not asc)
        elif coluna_atual == 3:
            itens.sort(key=lambda x: x.get("pago", 0), reverse=not asc)

        def ao_ordenar(e):
            estado_ordenacao["coluna"] = e.column_index
            estado_ordenacao["ascendente"] = e.ascending
            atualizar_tela()

        colunas_tabela = [
            ft.DataColumn(ft.Text("Tipo", size=f_size(12)), on_sort=ao_ordenar),
            ft.DataColumn(ft.Text("Descrição", size=f_size(12)), on_sort=ao_ordenar),
            ft.DataColumn(ft.Text("Valor", size=f_size(12)), numeric=True, on_sort=ao_ordenar),
            ft.DataColumn(ft.Text("Pago", size=f_size(12)), on_sort=ao_ordenar),
            ft.DataColumn(ft.Text("⚙️", size=f_size(12))),
        ]

        linhas = []
        for item in itens:
            cor_tipo = ft.Colors.GREEN_400 if item["tipo"] == "Receita" else ft.Colors.RED_400
            prefixo = "[Fixo] " if item["fixo"] else ""
            pago_txt = "Sim" if (item["tipo"] == "Despesa" and item.get("pago")) else ("-" if item["tipo"] == "Receita" else "Não")

            menu_opcoes = ft.PopupMenuButton(
                icon=ft.Icons.MORE_VERT,
                icon_size=f_size(18),
                items=[
                    ft.PopupMenuItem(
                        content=ft.Text("Editar", size=f_size(12)),
                        icon=ft.Icons.EDIT,
                        on_click=lambda e, it=item: construir_tela_editar(it),
                    ),
                    ft.PopupMenuItem(
                        content=ft.Text("Excluir", size=f_size(12)),
                        icon=ft.Icons.DELETE,
                        on_click=lambda e, it_id=item["id"]: confirmar_exclusao(it_id),
                    ),
                ]
            )

            def selecionar_item(it=item):
                nonlocal item_selecionado
                item_selecionado = it
                tipo_str = it['tipo']
                desc_str = f"{'[Fixo] ' if it['fixo'] else ''}{it['descricao']}"
                val_str = f"R$ {it['valor']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                forma_str = it.get('forma_pagamento', '') if tipo_str == 'Despesa' else 'N/A'
                pago_val = "Sim" if (tipo_str == 'Despesa' and it.get('pago')) else ("Não aplicável" if tipo_str == 'Receita' else "Não")
                
                txt_detalhes_info.value = f"[{tipo_str.upper()}] {desc_str}\nValor: {val_str} | Pag: {forma_str}\nPago: {pago_val}"
                txt_detalhes_info.color = ft.Colors.CYAN_300
                page.update()

            celulas = [
                ft.DataCell(ft.Text(item["tipo"][:3].upper(), color=cor_tipo, weight=ft.FontWeight.BOLD, size=f_size(12)), on_tap=lambda e, it=item: selecionar_item(it)),
                ft.DataCell(ft.Text(f"{prefixo}{item['descricao']}", size=f_size(12)), on_tap=lambda e, it=item: selecionar_item(it)),
                ft.DataCell(ft.Text(f"{item['valor']:.2f}", size=f_size(12)), on_tap=lambda e, it=item: selecionar_item(it)),
                ft.DataCell(ft.Text(pago_txt, size=f_size(12)), on_tap=lambda e, it=item: selecionar_item(it)),
                ft.DataCell(menu_opcoes),
            ]
            
            linhas.append(ft.DataRow(cells=celulas))

        tabela = ft.DataTable(
            columns=colunas_tabela,
            rows=linhas,
            column_spacing=12,
            horizontal_margin=8,
            heading_row_height=36,
            data_row_min_height=38,
            data_row_max_height=42,
            sort_column_index=estado_ordenacao["coluna"],
            sort_ascending=estado_ordenacao["ascendente"],
        )

        return ft.ListView(
            controls=[
                ft.Row([tabela], scroll=ft.ScrollMode.AUTO),
            ],
            expand=True,
            spacing=0,
        )

    def construir_tela_adicionar():
        dd_tipo = ft.Dropdown(
            label="Tipo",
            options=[ft.dropdown.Option("Receita"), ft.dropdown.Option("Despesa")],
            value="Receita",
            dense=True,
        )
        tf_desc = ft.TextField(label="Descrição (ex: Salário, Luz)", dense=True)
        tf_valor = ft.TextField(
            label="Valor (ex: 1500.00)",
            keyboard_type=ft.KeyboardType.NUMBER,
            dense=True,
        )

        dd_forma_pagamento = ft.Dropdown(
            label="Forma de Pagamento",
            options=[
                ft.dropdown.Option("Crédito"),
                ft.dropdown.Option("Débito"),
                ft.dropdown.Option("Vale"),
                ft.dropdown.Option("Dinheiro"),
            ],
            value="Crédito",
            dense=True,
        )
        cb_pago = ft.Checkbox(label="Pago", value=False)

        container_despesa_campos = ft.Column([
            dd_forma_pagamento,
            cb_pago,
        ], spacing=4)

        def on_tipo_change(e):
            eh_despesa = (dd_tipo.value == "Despesa")
            container_despesa_campos.visible = eh_despesa
            page.update()

        dd_tipo.on_change = on_tipo_change
        container_despesa_campos.visible = False

        modo_selecionado = {"valor": "avulso"}

        def on_modo_change(e):
            modo_selecionado["valor"] = e.control.value or "avulso"

        rg_modo = ft.RadioGroup(
            value="avulso",
            on_change=on_modo_change,
            content=ft.Column([
                ft.Radio(value="avulso", label="Somente neste mês"),
                ft.Radio(value="parcelado", label="Repetir por X meses"),
                ft.Radio(value="fixo", label="Fixo (todos os meses do ano)"),
            ], spacing=0),
        )
        tf_qtd_meses = ft.TextField(
            label="Qtd. meses (se repetir)",
            value="2",
            keyboard_type=ft.KeyboardType.NUMBER,
            dense=True,
        )

        def salvar_click(e):
            try:
                desc = (tf_desc.value or "").strip()
                valor_str = (tf_valor.value or "").strip().replace(",", ".")
                if not desc or not valor_str:
                    avisar("Preencha descrição e valor.", True)
                    return

                try:
                    valor = float(valor_str)
                except ValueError:
                    avisar("Valor numérico inválido.", True)
                    return

                tipo_val = dd_tipo.value or "Receita"
                forma_pgto = dd_forma_pagamento.value if tipo_val == "Despesa" else ""
                pago_val = 1 if (tipo_val == "Despesa" and cb_pago.value) else 0

                modo = rg_modo.value or modo_selecionado["valor"]
                mes_base = dd_mes.value or mes_atual

                if modo == "avulso":
                    db.adicionar_lancamento(mes_base, tipo_val, desc, valor, fixo=0, forma_pagamento=forma_pgto, pago=pago_val)
                    msg_ok = f"{tipo_val} adicionada em {formatar_mes_br(mes_base)}."
                elif modo == "parcelado":
                    try:
                        qtd = max(1, int(tf_qtd_meses.value or "1"))
                    except ValueError:
                        avisar("Quantidade de meses inválida.", True)
                        return
                    for m in somar_meses(mes_base, qtd):
                        db.adicionar_lancamento(m, tipo_val, desc, valor, fixo=0, forma_pagamento=forma_pgto, pago=pago_val)
                    msg_ok = f"Lançado em {qtd} meses a partir de {formatar_mes_br(mes_base)}."
                else:
                    ano = mes_base.split("-")[0]
                    mes_inicio = int(mes_base.split("-")[1])
                    for m in range(mes_inicio, 13):
                        ym = f"{ano}-{m:02d}"
                        db.adicionar_lancamento(ym, tipo_val, desc, valor, fixo=1, forma_pagamento=forma_pgto, pago=pago_val)
                    msg_ok = f"Lançamento fixo aplicado de {formatar_mes_br(mes_base)} até {m}/12."

                tf_desc.value = ""
                tf_valor.value = ""
                cb_pago.value = False

                popular_dropdown_meses()
                mudar_aba("principal")
                avisar(msg_ok)
            except Exception as ex:
                avisar(f"Erro ao salvar: {ex}", True)

        return ft.ListView(
            controls=[
                ft.Text("Adicionar Lançamento", size=f_size(15), weight=ft.FontWeight.BOLD),
                dd_tipo,
                tf_desc,
                tf_valor,
                container_despesa_campos,
                rg_modo,
                tf_qtd_meses,
                ft.Button("Salvar no Banco Local", icon=ft.Icons.SAVE, on_click=salvar_click),
            ],
            expand=True,
            spacing=6,
        )

    def construir_tela_configurar():
        def alterar_zoom(delta):
            nonlocal zoom_atual
            zoom_atual = round(max(0.8, min(1.8, zoom_atual + delta)), 2)
            salvar_config_zoom(zoom_atual)
            avisar(f"Zoom ajustado para {int(zoom_atual * 100)}%. Reinicie a aba para aplicar totalmente.")
            atualizar_tela()

        return ft.ListView(
            controls=[
                ft.Text("Configurar Acessibilidade & Dados:", size=f_size(15), weight=ft.FontWeight.BOLD),
                ft.Divider(),
                ft.Text(f"Zoom da Tela (Acessibilidade): {int(zoom_atual * 100)}%", weight=ft.FontWeight.W_600, size=f_size(13)),
                ft.Row([
                    ft.OutlinedButton("Aumentar Zoom (+)", icon=ft.Icons.ZOOM_IN, on_click=lambda _: alterar_zoom(0.1)),
                    ft.OutlinedButton("Reduzir Zoom (-)", icon=ft.Icons.ZOOM_OUT, on_click=lambda _: alterar_zoom(-0.1)),
                ], spacing=6),
                ft.Divider(),
                ft.Text("Importar / Exportar Excel (.xlsx)", weight=ft.FontWeight.W_600, size=f_size(13)),
                ft.Row([
                    ft.OutlinedButton("Exportar Excel", icon=ft.Icons.UPLOAD_FILE, on_click=exportar_excel_click),
                    ft.OutlinedButton("Importar Excel", icon=ft.Icons.DOWNLOAD, on_click=importar_excel_click),
                ], spacing=6),
                ft.Divider(),
                ft.Text("Banco de Dados (.db)", weight=ft.FontWeight.W_600, size=f_size(13)),
                ft.Row([
                    ft.OutlinedButton("Salvar .db", icon=ft.Icons.BACKUP, on_click=exportar_db_click),
                    ft.OutlinedButton("Restaurar .db", icon=ft.Icons.RESTORE, on_click=importar_db_click),
                ], spacing=6),
            ],
            expand=True,
            spacing=8,
        )

    def mudar_aba(nova_aba):
        nonlocal aba_ativa
        aba_ativa = nova_aba
        atualizar_tela()

    def atualizar_tela():
        atualizar_resumo()
        dd_mes.visible = (aba_ativa == "principal")

        if aba_ativa == "principal":
            corpo_central.content = construir_tela_principal()
        elif aba_ativa == "adicionar":
            corpo_central.content = construir_tela_adicionar()
        elif aba_ativa == "configurar":
            corpo_central.content = construir_tela_configurar()

        page.update()

    dd_mes.on_select = lambda _: atualizar_tela()
    dd_mes.on_change = lambda _: atualizar_tela()

    def ao_fechar(e):
        fazer_backup_automatico(DB_NAME)
        page.window.destroy()

    page.window.on_close = ao_fechar

    # Topo organizado lado a lado sem texto redundante de Mês
    bloco_topo = ft.Container(
        border=ft.Border.all(1, ft.Colors.OUTLINE),
        border_radius=6,
        padding=8,
        content=ft.Row([
            # Esquerda: Dropdown de Mês e Resumo Financeiro
            ft.Column([
                dd_mes,
                txt_receitas,
                txt_despesas,
                txt_saldo,
            ], spacing=3, alignment=ft.MainAxisAlignment.CENTER),
            # Direita: Detalhes do item selecionado
            container_detalhes_direita
        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN, spacing=8),
    )

    bloco_avisos = ft.Container(
        border=ft.Border.all(1, ft.Colors.OUTLINE),
        border_radius=6,
        padding=6,
        content=ft.Row([
            ft.Icon(ft.Icons.INFO_OUTLINE, size=f_size(16), color=ft.Colors.AMBER_300),
            ft.Container(content=txt_aviso, expand=True),
        ], spacing=6),
    )

    bloco_menu = ft.Row(
        controls=[
            ft.Button(
                "Principal",
                expand=True,
                on_click=lambda e: mudar_aba("principal"),
                style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=6)),
            ),
            ft.Button(
                "Adicionar",
                expand=True,
                on_click=lambda e: mudar_aba("adicionar"),
                style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=6)),
            ),
            ft.Button(
                "Configurar",
                expand=True,
                on_click=lambda e: mudar_aba("configurar"),
                style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=6)),
            ),
        ],
        spacing=4,
    )

    page.add(
        ft.Column(
            [
                bloco_topo,
                corpo_central,
                bloco_avisos,
                bloco_menu,
            ],
            expand=True,
            spacing=6,
        )
    )

    atualizar_tela()


if __name__ == "__main__":
    ft.run(main)