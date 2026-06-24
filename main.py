import flet as ft
import pandas as pd
import sqlite3
import threading
import http.server
import socketserver
import os
import sys


class GamesApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.start_local_server()
        self.page.window.width = 1150
        self.page.window.height = 900
        self.page.bgcolor = "#12141C"
        
        if getattr(sys, 'frozen', False):
            data_folder = os.path.dirname(sys.executable)
        else:
            data_folder = os.path.dirname(os.path.abspath(__file__))

        self.db_file = os.path.join(data_folder, "games_store.db")
        self.report_txt = os.path.join(data_folder, "games_report.txt")
        self.report_excel = os.path.join(data_folder, "games_report.xlsx")

        self.df = pd.DataFrame()
        self.current_df = pd.DataFrame()

        self.columns_map = {
            "ID": "id",
            "Название": "name",
            "Жанр": "category",
            "Цена": "price",
            "Разработчик": "developer",
            "Дата выхода": "release_date",
            "Рейтинг": "rating",
            "Цена со скидкой": "discount_price"
        }
        self.display_cols = list(self.columns_map.keys())
        self.reverse_map = {v: k for k, v in self.columns_map.items()}

        self.combo_filter_col = ft.Dropdown(
            label="Фильтр по:",
            options=[ft.dropdown.Option("(Нет)")] + [ft.dropdown.Option(c) for c in self.display_cols],
            value="(Нет)", expand=1, border_color="#2A2F42", focused_border_color="#6366F1",
        )
        self.entry_filter_val = ft.TextField(
            label="Значение фильтра", expand=2, border_color="#2A2F42", focused_border_color="#6366F1"
        )

        self.combo_sort_col = ft.Dropdown(
            label="Сортировка по:",
            options=[ft.dropdown.Option("(Нет)")] + [ft.dropdown.Option(c) for c in self.display_cols],
            value="(Нет)", expand=1, border_color="#2A2F42", focused_border_color="#6366F1"
        )

        self.switch_asc = ft.Switch(
            label="По возрастанию",
            value=True,
            active_color="#6366F1",
            label_position=ft.LabelPosition.RIGHT
        )

        self.entry_cols = ft.TextField(
            label="Столбцы (id, name, price, category...)", expand=1, border_color="#2A2F42",
            focused_border_color="#6366F1"
        )

        self.btn_apply = ft.FilledButton(
            content="Применить", on_click=self.apply_ops, icon=ft.Icons.FILTER_ALT_ROUNDED,
            style=ft.ButtonStyle(bgcolor="#6366F1", color="#FFFFFF", shape=ft.RoundedRectangleBorder(radius=8))
        )

        self.entry_add_name = ft.TextField(label="Название игры", expand=2, border_color="#2A2F42",
                                           focused_border_color="#6366F1")
        self.entry_add_description = ft.TextField(
            label="Описание игры",
            expand=True,
            multiline=True,
            min_lines=1,
            max_lines=3,
            border_color="#2A2F42",
            focused_border_color="#6366F1"
        )
        self.entry_add_category = ft.TextField(label="Жанр", expand=1, border_color="#2A2F42",
                                               focused_border_color="#6366F1")
        self.entry_add_price = ft.TextField(label="Цена", expand=1, border_color="#2A2F42",
                                            focused_border_color="#6366F1")
        self.entry_add_developer = ft.TextField(label="Разработчик", expand=1, border_color="#2A2F42",
                                                focused_border_color="#6366F1")
        self.entry_add_release_date = ft.TextField(label="Дата выхода", expand=1, border_color="#2A2F42",
                                                   focused_border_color="#6366F1")
        self.entry_add_rating = ft.TextField(label="Рейтинг", expand=1, border_color="#2A2F42",
                                             focused_border_color="#6366F1")
        self.entry_add_discount = ft.TextField(label="Цена со скидкой", expand=1, border_color="#2A2F42",
                                               focused_border_color="#6366F1")

        self.btn_add_game = ft.FilledButton(
            content="Добавить игру", on_click=self.add_game, icon=ft.Icons.ADD_ROUNDED,
            style=ft.ButtonStyle(bgcolor="#4F46E5", color="#FFFFFF", shape=ft.RoundedRectangleBorder(radius=8))
        )

        self.entry_delete_id = ft.TextField(label="ID для удаления", expand=1, border_color="#2A2F42",
                                            focused_border_color="#EF4444")
        self.btn_delete_game = ft.FilledButton(
            content="Удалить по ID", on_click=self.delete_game, icon=ft.Icons.DELETE_ROUNDED,
            style=ft.ButtonStyle(bgcolor="#EF4444", color="#FFFFFF", shape=ft.RoundedRectangleBorder(radius=8))
        )

        self.reports_list = [
            "1. KPI: Общая выручка и продажи",
            "2. ABC-анализ (Доля выручки по играм)",
            "3. Топ-5 самых продаваемых игр (шт.)",
            "4. Топ-5 свежих релизов (Новинки)",
            "5. Распределение игр по категориям",
            "6. Топ-5 игр по рейтингу",
            "7. Топ-5 самых дорогих игр",
            "8. Топ-5 игр с максимальной скидкой (%)",
            "9. Игры со скидкой",
            "10. Статистика по разработчикам"
        ]

        self.combo_report = ft.Dropdown(
            label="Выберите аналитический отчет:",
            options=[ft.dropdown.Option(r) for r in self.reports_list],
            expand=True, border_color="#2A2F42", focused_border_color="#10B981",
        )

        self.btn_run_report = ft.FilledButton(
            content="Отчет", on_click=self.generate_report, icon=ft.Icons.ANALYTICS_ROUNDED,
            style=ft.ButtonStyle(bgcolor="#10B981", color="#FFFFFF", shape=ft.RoundedRectangleBorder(radius=8))
        )

        self.data_table = ft.DataTable(
            columns=[ft.DataColumn(ft.Text(""))],
            rows=[],
            border_radius=8,
            heading_row_color="#1E2235",
            divider_thickness=1,
            horizontal_lines=ft.border.BorderSide(1, "#2A2F42"),
            vertical_lines=ft.border.BorderSide(1, "#1E2235"),
            column_spacing=20
        )

        self.status_text = ft.Text("Готово", color="#94A3B8", size=13)

    def start_local_server(self):
        try:
            if getattr(sys, 'frozen', False):
                script_dir = sys._MEIPASS
            else:
                script_dir = os.path.dirname(os.path.abspath(__file__))
            
            class CustomHandler(http.server.SimpleHTTPRequestHandler):
                def __init__(self, *args, **kwargs):
                    super().__init__(*args, directory=script_dir, **kwargs)

            self.httpd = socketserver.TCPServer(("localhost", 8000), CustomHandler)
            thread = threading.Thread(target=self.httpd.serve_forever, daemon=True)
            thread.start()
        except Exception as ex:
            print(f"Ошибка запуска сервера: {str(ex)}")

    def build(self):
        self.page.appbar = ft.AppBar(
            title=ft.Text("Магазин Игр ZPay", size=18, weight=ft.FontWeight.W_600, color="#F8FAFC"),
            bgcolor="#161926",
            center_title=False,
            actions=[
                ft.IconButton(icon=ft.Icons.INFO, tooltip="О разработчике", url="http://localhost:8000/site.html"),
                ft.IconButton(icon=ft.Icons.TEXT_SNIPPET_ROUNDED, icon_color="#38BDF8", tooltip="Экспорт TXT",
                              on_click=self.export_txt),
                ft.IconButton(icon=ft.Icons.GRID_ON_ROUNDED, icon_color="#34D399", tooltip="Экспорт Excel",
                              on_click=self.export_excel),
                ft.IconButton(icon=ft.Icons.REFRESH_ROUNDED, icon_color="#FB7185", tooltip="Сбросить фильтры",
                              on_click=self.reset_data),
            ]
        )

        filter_card = ft.Container(
            content=ft.Column([
                ft.Text("Фильтры и сортировка", color="#6366F1", weight=ft.FontWeight.BOLD, size=14),
                ft.Row([self.combo_filter_col, self.entry_filter_val], spacing=10),
                ft.Row([self.combo_sort_col, self.entry_cols], spacing=10),
                ft.Row([self.switch_asc, self.btn_apply], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                ft.Divider(color="#2A2F42", thickness=1),
                ft.Text("Бизнес-аналитика (БД)", color="#10B981", weight=ft.FontWeight.BOLD, size=14),
                ft.Row([self.combo_report, self.btn_run_report], spacing=10),
            ], spacing=12),
            padding=15,
            bgcolor="#161926",
            border_radius=12,
        )

        management_card = ft.Container(
            content=ft.Column([
                ft.Text("Управление базой данных (Добавление / Удаление)", color="#F43F5E", weight=ft.FontWeight.BOLD,
                        size=14),
                ft.Row([self.entry_add_name, self.entry_add_category, self.entry_add_price], spacing=10),
                ft.Row([self.entry_add_developer, self.entry_add_release_date, self.entry_add_rating,
                        self.entry_add_discount], spacing=10),
                ft.Row([self.entry_add_description], spacing=10),
                ft.Row([self.btn_add_game], alignment=ft.MainAxisAlignment.END),
                ft.Divider(color="#2A2F42", thickness=1),
                ft.Row([self.entry_delete_id, self.btn_delete_game], spacing=10),
            ], spacing=12),
            padding=15,
            bgcolor="#161926",
            border_radius=12,
        )

        table_container = ft.Container(
            content=ft.Row(
                controls=[self.data_table],
                scroll=ft.ScrollMode.ADAPTIVE,
            ),
            bgcolor="#161926",
            border_radius=12,
            padding=10,
        )

        status_bar = ft.Container(content=self.status_text, padding=ft.Padding.only(left=5, top=5, bottom=5))

        return ft.Container(
            content=ft.Column([filter_card, management_card, table_container, status_bar], spacing=15,
                              scroll=ft.ScrollMode.ADAPTIVE),
            padding=10, expand=True
        )

    def load_db(self):
        if not os.path.exists(self.db_file):
            self._init_db_schema()

        try:
            conn = sqlite3.connect(self.db_file)
            self.df = pd.read_sql_query("SELECT * FROM games", conn)
            conn.close()
            self.reset_data()
        except Exception as e:
            self._show_snackbar(f"Ошибка БД: {str(e)}", error=True)

    def _init_db_schema(self):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS games 
                          (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                           name TEXT, 
                           category TEXT, 
                           price REAL, 
                           developer TEXT, 
                           release_date TEXT, 
                           rating REAL, 
                           discount_price REAL,
                           description TEXT,
                           image_url TEXT)''')
        
        cursor.execute("SELECT COUNT(*) FROM games")
        if cursor.fetchone()[0] == 0:
            initial_games = [
                ("The Witcher 3: Wild Hunt", "RPG", 1500.0, "CD Projekt RED", "2015-05-19", 4.9, 450.0, "Культовая RPG", ""),
                ("Cyberpunk 2077", "RPG", 2000.0, "CD Projekt RED", "2020-12-10", 4.5, 1000.0, "Экшен в будущем", ""),
                ("GTA V", "Action", 1200.0, "Rockstar Games", "2013-09-17", 4.8, None, "Криминальный экшен", ""),
                ("Minecraft", "Sandbox", 1900.0, "Mojang Studios", "2011-11-18", 4.7, None, "Песочница", ""),
                ("Elden Ring", "Action-RPG", 3999.0, "FromSoftware", "2022-02-25", 4.9, 2799.0, "Сложная RPG", "")
            ]
            cursor.executemany(
                "INSERT INTO games (name, category, price, developer, release_date, rating, discount_price, description, image_url) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                initial_games
            )
            conn.commit()
        conn.close()

    def refresh_data_from_db(self):
        try:
            conn = sqlite3.connect(self.db_file)
            self.df = pd.read_sql_query("SELECT * FROM games", conn)
            conn.close()
            self.apply_ops()
        except Exception as e:
            self._show_snackbar(f"Ошибка обновления данных: {str(e)}", error=True)

    def add_game(self, e):
        name = (self.entry_add_name.value or "").strip()
        category = (self.entry_add_category.value or "").strip()
        price = (self.entry_add_price.value or "").strip()
        developer = (self.entry_add_developer.value or "").strip()
        release_date = (self.entry_add_release_date.value or "").strip()
        rating = (self.entry_add_rating.value or "").strip()
        discount = (self.entry_add_discount.value or "").strip()
        description = (self.entry_add_description.value or "").strip()

        if not name:
            self._show_snackbar("Поле 'Название игры' обязательно!", error=True)
            return

        try:
            price_val = float(price) if price else 0.0
            rating_val = float(rating) if rating else 0.0
            discount_val = float(discount) if discount and discount != "" else None

            conn = sqlite3.connect(self.db_file, timeout=10)
            cursor = conn.cursor()

            cursor.execute(
                "INSERT INTO games (name, category, price, developer, release_date, rating, discount_price, description, image_url) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (name, category, price_val, developer, release_date, rating_val, discount_val, description, "")
            )
            conn.commit()
            conn.close()

            self.entry_add_name.value = ""
            self.entry_add_category.value = ""
            self.entry_add_price.value = ""
            self.entry_add_developer.value = ""
            self.entry_add_release_date.value = ""
            self.entry_add_rating.value = ""
            self.entry_add_discount.value = ""
            self.entry_add_description.value = ""

            self._show_snackbar(f"Игра '{name}' успешно добавлена!")
            self.refresh_data_from_db()
            self.page.update()

        except ValueError:
            self._show_snackbar("Цена и Рейтинг должны быть числами!", error=True)
        except Exception as ex:
            self._show_snackbar(f"Ошибка БД: {str(ex)}", error=True)

    def delete_game(self, e):
        game_id = self.entry_delete_id.value.strip()
        if not game_id:
            self._show_snackbar("Введите корректный ID для удаления!", error=True)
            return

        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()

            cursor.execute("SELECT id, name FROM games WHERE id = ?", (game_id,))
            row = cursor.fetchone()

            if not row:
                self._show_snackbar(f"Игра с ID {game_id} не найдена!", error=True)
                conn.close()
                return

            cursor.execute("DELETE FROM games WHERE id = ?", (game_id,))
            conn.commit()
            conn.close()

            self.entry_delete_id.value = ""
            self._show_snackbar(f"Игра '{row[1]}' (ID: {game_id}) успешно удалена!")
            self.refresh_data_from_db()

        except Exception as ex:
            self._show_snackbar(f"Ошибка удаления: {str(ex)}", error=True)

    def generate_report(self, e=None):
        report_name = self.combo_report.value
        if not report_name:
            self._show_snackbar("Выберите тип отчета из списка", error=True)
            return

        try:
            conn = sqlite3.connect(self.db_file)
            self.reverse_map = {}

            if "1." in report_name:
                query = "SELECT COUNT(*) as 'Количество игр в базе (шт)', SUM(price) as 'Общая стоимость пула (руб)', AVG(price) as 'Средняя цена (руб)' FROM games"
                self.current_df = pd.read_sql_query(query, conn)

            elif "2." in report_name:
                df_abc = pd.read_sql_query("SELECT name as 'Название', price as 'Выручка' FROM games", conn)
                total_rev = df_abc['Выручка'].sum()
                df_abc = df_abc.sort_values(by='Выручка', ascending=False)
                df_abc['Доля (%)'] = ((df_abc['Выручка'] / (total_rev if total_rev > 0 else 1)) * 100).round(2)
                df_abc['Нарастающая доля (%)'] = df_abc['Доля (%)'].cumsum()
                df_abc['Класс'] = df_abc['Нарастающая доля (%)'].apply(
                    lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C'))
                self.current_df = df_abc

            elif "3." in report_name:
                self.current_df = pd.read_sql_query(
                    "SELECT id as 'ID', name as 'Название', category as 'Жанр', rating as 'Рейтинг популярности' FROM games ORDER BY rating DESC LIMIT 5",
                    conn)

            elif "4." in report_name:
                self.current_df = pd.read_sql_query(
                    "SELECT id as 'ID', name as 'Название', category as 'Жанр', release_date as 'Дата выхода', developer as 'Разработчик' FROM games ORDER BY release_date DESC LIMIT 5",
                    conn)

            elif "5." in report_name:
                self.current_df = pd.read_sql_query(
                    "SELECT category as 'Жанр', COUNT(*) as 'Кол-во игр (шт)', AVG(price) as 'Средняя цена жанра' FROM games GROUP BY category",
                    conn)

            elif "6." in report_name:
                self.current_df = pd.read_sql_query(
                    "SELECT id as 'ID', name as 'Название', developer as 'Разработчик', rating as 'Рейтинг' FROM games ORDER BY rating DESC LIMIT 5",
                    conn)

            elif "7." in report_name:
                self.current_df = pd.read_sql_query(
                    "SELECT id as 'ID', name as 'Название', category as 'Жанр', price as 'Цена (руб)' FROM games ORDER BY price DESC LIMIT 5",
                    conn)

            elif "8." in report_name:
                self.current_df = pd.read_sql_query(
                    "SELECT id as 'ID', name as 'Название', price as 'Старая цена', discount_price as 'Цена со скидкой', ROUND(((price - discount_price) / price) * 100, 1) as 'Скидка (%)' FROM games WHERE discount_price IS NOT NULL AND price > 0 ORDER BY ((price - discount_price) / price) DESC LIMIT 5",
                    conn)

            elif "9." in report_name:
                self.current_df = pd.read_sql_query(
                    "SELECT id as 'ID', name as 'Название', price as 'Старая цена', discount_price as 'Цена со скидкой', (price - discount_price) as 'Выгода' FROM games WHERE discount_price IS NOT NULL AND discount_price > 0",
                    conn)

            elif "10." in report_name:
                self.current_df = pd.read_sql_query(
                    "SELECT developer as 'Разработчик', COUNT(*) as 'Всего игр выпущено', AVG(rating) as 'Средний рейтинг игр' FROM games GROUP BY developer",
                    conn)

            conn.close()
            self.update_table()
            self._show_snackbar(f"Отчет успешно построен!")
        except Exception as ex:
            self._show_snackbar(f"Ошибка при генерации отчета: {str(ex)}", error=True)

    def update_table(self):
        self.data_table.columns.clear()
        self.data_table.rows.clear()

        if self.current_df.empty:
            self.data_table.columns.append(ft.DataColumn(ft.Text("Нет данных для отображения", color="#94A3B8")))
            self.status_text.value = "Строк: 0"
            self.page.update()
            return

        for col in self.current_df.columns:
            header_text = self.reverse_map.get(col, col)
            self.data_table.columns.append(
                ft.DataColumn(ft.Text(str(header_text), weight=ft.FontWeight.W_600, color="#F8FAFC")))

        for row in self.current_df.itertuples(index=False, name=None):
            cells = [ft.DataCell(ft.Text(str(val) if pd.notna(val) else "-", color="#E2E8F0")) for val in row]
            self.data_table.rows.append(ft.DataRow(cells=cells))

        self.status_text.value = f"Отображено строк: {len(self.current_df)}"
        self.page.update()

    def apply_ops(self, e=None):
        self.reverse_map = {v: k for k, v in self.columns_map.items()}
        df = self.df.copy()

        f_col_display = self.combo_filter_col.value
        f_val = self.entry_filter_val.value.strip() if self.entry_filter_val.value else ""

        if f_col_display != "(Нет)" and f_val:
            actual_col = self.columns_map[f_col_display]
            if actual_col in df.columns:
                df = df[df[actual_col].astype(str).str.contains(f_val, case=False, na=False)]

        s_col_display = self.combo_sort_col.value
        if s_col_display != "(Нет)":
            actual_sort_col = self.columns_map[s_col_display]
            if actual_sort_col in df.columns:
                df = df.sort_values(by=actual_sort_col, ascending=self.switch_asc.value)

        c_input = self.entry_cols.value.strip() if self.entry_cols.value else ""
        if c_input:
            req_cols = [c.strip() for c in c_input.split(",")]
            valid_cols = [c for c in req_cols if c in df.columns]
            if valid_cols:
                df = df[valid_cols]

        self.current_df = df
        self.update_table()

    def reset_data(self, e=None):
        self.reverse_map = {v: k for k, v in self.columns_map.items()}
        self.current_df = self.df.copy()

        self.combo_filter_col.value = "(Нет)"
        self.entry_filter_val.value = ""
        self.combo_sort_col.value = "(Нет)"
        self.entry_cols.value = ""
        self.combo_report.value = None
        self.switch_asc.value = True

        self.update_table()

    def export_txt(self, e=None):
        if self.current_df.empty:
            self._show_snackbar("Нет данных для экспорта", error=True)
            return
        try:
            df_export = self.current_df.rename(columns=self.reverse_map)
            with open(self.report_txt, "w", encoding="utf-8") as f:
                f.write(df_export.to_string(index=False))
            self._show_snackbar(f"Сохранено в файл {os.path.basename(self.report_txt)}!")
        except Exception as ex:
            self._show_snackbar(f"Ошибка экспорта: {str(ex)}", error=True)

    def export_excel(self, e=None):
        if self.current_df.empty:
            self._show_snackbar("Нет данных для экспорта", error=True)
            return
        try:
            df_export = self.current_df.rename(columns=self.reverse_map)
            df_export.to_excel(self.report_excel, index=False)
            self._show_snackbar(f"Сохранено в файл {os.path.basename(self.report_excel)}!")
        except Exception as ex:
            self._show_snackbar(f"Ошибка экспорта: {str(ex)}", error=True)

    def _show_snackbar(self, message: str, error: bool = False):
        self.page.snack_bar = ft.SnackBar(
            content=ft.Text(message, color="#FFFFFF"),
            bgcolor="#EF4444" if error else "#10B981",
        )
        self.page.snack_bar.open = True
        self.page.update()


def main(page: ft.Page):
    page.title = "ZPay Games Store"
    page.theme_mode = ft.ThemeMode.DARK
    page.padding = 0

    app = GamesApp(page)
    page.add(app.build())
    app.load_db()


if __name__ == "__main__":
    ft.app(target=main)
