"""Библиотека предоставляет класс и набор полезных функция для
формирования технического отчета в формате Microsoft Word
с помощью библиотеки python-docx. Таких как:

    * Report - класс для создания отчетов в файле Microsoft Word
    * get_xls_sheet_quantity - функция возвращающая количество листов в файле Excel
    * set_table_columns_width - установка ширины колонок в таблице docx.Table
    * set_table_style - установка стиля текста в ячейках всей таблицы docx.Table
    * set_table_rows_style - установка стиля текста в указанных строках таблицы docx.Table
"""

import os
import random
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Cm


class Report:
    """Класс для создания отчетов в файле Microsoft Word.
    Позволяет создавать отчеты с текстовой информацией, таблицами и графиками
    и вставлять их в существующий документ или создавать новый.

    Args:
        file_path (str, optional): путь к файлу, в который добавляется отчет.
    Если не указан или файл не существует, отчет создается на основе шаблона по умолчанию 
    DEFAULT_TEMPLATE. По умолчанию None.
    """

    # Путь к файлу шаблона отчета по умолчанию
    DEFAULT_TEMPLATE = "report/templates/template.docx"

    STYLES = {
        "table": {
            "title": "Т-название",
            "text": "Т-таблица",
            "text_heading": "Т-таблица-заголовок",
            "footer": "Т-примечание",
            "table": "Table Grid",
        },
        "figure": {
            "title": "Р-название",
            "fig": "Р-рисунок",
        },
        "heading": {
            1: "Heading 1",
            2: "Heading 2",
            3: "Heading 3",
            4: "Heading 4",
            5: "Heading 5",
            6: "Heading 6",
        },
    }

    TEXT = {
        "parameter": "Параметр",
        "figure": "Рисунок",
        "table": "Таблица",
    }

    def __init__(self, file_path=None) -> None:
        """Инициализирует новый экземпляр класса Report.

        Args:
            file_path (str, optional): Путь к файлу, в который добавляется отчет.
        Если не указан или файл не существует, отчет создается на основе шаблона по умолчанию.
        По умолчанию None.
        """
        self.doc = Document

        if file_path:
            print(f"\nРезультаты расчетов будут добавлены к файлу {file_path}")

            # Пытаемся загрузить существующий файл
            # и вставляем разрыв страницы
            self.set_template(file_path)
        else:
            # Создаем новый документ на основе стандартного шаблона
            self._init_empty_report()

    def _init_empty_report(self) -> None:
        """Метод инициализирует пустой отчет на основе стандартного шаблона
        и удаляет первый параграф.
        """
        self.set_template(self.DEFAULT_TEMPLATE)
        self.remove_paragraph()

    def add_paragraph(self, text: str, style=None) -> None:
        """Добавляет в документ абзац с указанным текстом и стилем.

        Args:
          text (str): Текст который необходимо добавить в документ в виде абзаца
          style (str): Стиль применяемый к добавляемому абзацу.
        """
        self.doc.add_paragraph(text, style=style)

    def add_heading(self, text, level=1) -> None:
        """Добавляет в документ абзац заголовка заданного уровня (1-6).

        Args:
          text (str): Текст заголовка
          level (int): Уровень заголовка (возможны значения 1-6). По умолчанию 1.
        """

        if not 1 <= level <= 6:
            print(
                "Уровень заголовка не может быть меньше 1 и больше 6, "
                f"заголовок '{text}' будет записан обычным текстом."
            )
            self.doc.add_paragraph(text)
            return

        self.doc.add_paragraph(text, style=self.STYLES["heading"][level])

    def insert_df_to_table(
        self,
        df,
        title=None,
        footer_text=None,
        col_names=None,
        col_widths=None,
        col_format=None,
        table_style=STYLES["table"]["table"],
        text_style=STYLES["table"]["text"],
        first_row_table_style=STYLES["table"]["text_heading"],
        row_names=None,
    ) -> None:
        """
        Метод вставляет Pandas DataFrame в файл отчета как отформатированную
        таблицу с возможностью указания заголовков для колонок, стилей таблицы и
        текста, ширины столбцов.

        Args:
            df (Pandas.DataFrame): DataFrame содержащий данные, которые необходимо вставить
        в таблицу документа Word.
            title (str): Параметр title используется для указания заголовка (названия) таблицы,
        которая будет вставлена в документ.
            col_names(tuple, list): Параметр переопределяет заголовки колонок в таблице.
        По умолчанию заголовки соответствуют названию колонок в DataFrame.
            col_widths (tuple): При указании, кортеж переопределяет ширину колонок
        таблицы по порядку.
            col_format (tuple): Параметр используется для указания стиля форматирования
        значений для каждого столбца таблицы.
        Форматы соответствует f-строке (пример: ":g", ":g", ":.2f")
            table_style (str): Название устанавливаемого стиля таблицы.
        Стиль должен присутствовать в файле шаблона отчета. По умолчанию "Table Grid".
            text_style: Название устанавливаемого основное стиля текста в таблице.
        Стиль должен присутствовать в файле шаблона отчета. По умолчанию "Т-таблица".
            first_row_table_style: Название устанавливаемого стиля строки заголовков таблицы.
        Стиль должен присутствовать в файле шаблона отчета. По умолчанию "Т-таблица-заголовок".

        Returns:
            Функция возвращает экземпляр docx.table.Table с произведенными изменениям.
        """
        doc = self.doc

        # Проверка на тип данных
        if not isinstance(df, pd.DataFrame):
            print("\nОшибка! Для вставки таблиц необходимо передать Pandas.DataFrame.")
            print(f"Таблица '{title}' не будет вставлена в отчет.\n")
            return

        # Вставляем заголовок таблицы
        if title:
            doc.add_paragraph(
                f"{self.TEXT["table"]} — {title}", style=self.STYLES["table"]["title"]
            )

        # Количество строк и столбцов в таблице
        rows = df.shape[0]
        columns = df.shape[1]

        # Учитываем дополнительную колонку для row_names
        add_row_names = isinstance(row_names, (list, tuple, str))
        if add_row_names:
            columns += 1

        # Добавляем таблицу в документ
        table = doc.add_table(rows + 1, columns, style=table_style)

        # Получаем доступ к ячейкам экземпляра таблицы для
        # увеличения производительности, все последующие
        # операции производим с ячейками а не с экземпляром таблицы
        cells = table._cells

        # Устанавливаем 1ю строку заголовков
        for column_idx, column_name in enumerate(df.columns):
            # Сдвиг на 1, если если присутствуют названия строк
            header_idx = column_idx + (1 if add_row_names else 0)
            if col_names:
                cells[header_idx].text = str(col_names[column_idx])
            else:
                cells[header_idx].text = str(column_name)

        # Добавляем заголовок для первой колонки, если row_names указано
        if add_row_names:
            cells[0].text = self.TEXT["parameter"]
            for row_idx, name in enumerate(row_names):
                cells[(row_idx + 1) * columns].text = str(name)

        # Записываем данные df в таблицу
        for row_idx in range(rows):
            for column_idx in range(df.shape[1]):
                cell_value = df.iat[row_idx, column_idx]

                cell_idx = (
                    (row_idx + 1) * columns + column_idx + (1 if add_row_names else 0)
                )

                # Если задан список стилей для форматирования текста
                # устанавливаем формат для значения каждой ячейки
                # иначе просто записываем строку значения в ячейку
                if col_format and isinstance(cell_value, (float, int)):
                    s = f"{{{col_format[column_idx]}}}"
                    cells[cell_idx].text = s.format(cell_value)
                else:
                    cells[cell_idx].text = str(cell_value)

        if col_widths:
            self._set_table_columns_width(table, col_widths)

        self._set_table_style(table, text_style, first_row_table_style)

        # Записываем служебный параграф после таблиц
        if footer_text:
            doc.add_paragraph(footer_text, style=self.STYLES["table"]["footer"])
        else:
            doc.add_paragraph("", style=self.STYLES["table"]["footer"])

        return table

    def insert_mpl_figure(self, chart, title="", dpi=200, width=16.5) -> None:
        """Метод вставляет график Matplotlib.plt в документ, предварительно
        сохранив его во временный файл. Устанавливает стили, добавляет заголовок
        и затем удаляет временный файл.

        Args:
        chart (Matplotlib.plt): график который необходимо вставить в документ
        title (str): заголовок графика
        dpi (int): разрешение с которым будет вставлен график. По умолчанию 200
        width (float): ширина графика в сантиметрах
        """
        doc = self.doc

        # Временная директория в которую сохраняется график
        temp_dir = "TEMP"

        # Название временного файла графика
        filename = f"temp_{random.randrange(1, 10**3):03}.png"

        # Путь сохранения графика в файл
        file_path = f"{temp_dir}/{filename}"
        # Создаем временную папку, если она отсутствует
        Path(file_path).parents[0].mkdir(parents=True, exist_ok=True)
        # Сохраняем график в файл
        chart.savefig(file_path, dpi=dpi)
        # Вставляем график в документ отчёта
        doc.add_picture(file_path, width=Cm(width))
        # Устанавливаем стиль графика
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.style = self.STYLES["figure"]["fig"]
        doc.add_paragraph(
            f"{self.TEXT["figure"]} — {title}", style=self.STYLES["figure"]["title"]
        )

        # Пытаемся удалить файл графика
        try:
            os.remove(file_path)
        except FileNotFoundError:
            pass

    def insert_page_break(self) -> None:
        """Метод вставляет в документ Word разрыв страницы
        в том месте где вызывается.
        """

        doc = self.doc
        paragraphs = doc.paragraphs
        run = paragraphs[-1].add_run()
        run.add_break(WD_BREAK.PAGE)

    def _set_table_columns_width(self, table, col_widths: tuple) -> None:
        """Метод устанавливает ширину столбцов в таблице docx
        поочередно проходя по каждой ячейке таблицы.

        Args:
            table (docx.table.Table): Таблица в документе docx.
            col_widths (tuple): Кортеж, содержащий ширину столбцов
        в сантиметрах по порядку.
        """
        set_table_columns_width(table, col_widths)

    def _set_table_style(
        self,
        table,
        style=STYLES["table"]["text"],
        first_row_style=STYLES["table"]["text_heading"],
    ) -> None:
        """Метод проходил по всем ячейкам таблицы table и устанавливает
        заданный стиль style параграфов в таблице. При желании можно
        указать стиль для заголовков таблицы (1-ая строка).

        Args:
            table (docx.table.Table): Таблица в документе docx
            style (str): Название устанавливаемого стиля. Стиль должен присутствовать
            в файле шаблона отчета. По умолчанию "Т-таблица".
            first_row_style (str): Названия стиля для первой строки таблицы.
            (строка заголовков). По умолчанию None.
        """
        set_table_style(table, style, first_row_style)

    def _set_table_rows_style(
        self, table, rows=(0, 1), style=STYLES["table"]["text_heading"]
    ) -> None:
        """Метод предназначен для установки заданного стиля для указанных строк таблицы.

        Args:
            table (docx.table.Table): Таблица в документе docx
            style (str): Название устанавливаемого стиля. Стиль должен присутствовать
            в файле шаблона отчета. По умолчанию "Т-таблица-заголовок".
        """
        set_table_rows_style(table, rows, style)

    def set_last_paragraph_style(self, style: str) -> None:
        """Метод устанавливает стиль последнего по порядку
        параграфа в документе Word.

        Args:
            style (str): Название стиля в документе
        (стиль обязательно должен присутствовать в документе отчёта).
        """
        doc = self.doc
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.style = style

    def set_template(self, template: str) -> None:
        """Метод позволяет установить шаблон документа Microsoft Word.
        Если указанный файл шаблон не найден, используется шаблон по умолчанию
        ("report/assets/template.docx").

        Args:
            template (str): Путь к файлу шаблону
        """

        # Проверка типа аргумента template
        if not isinstance(template, str):
            print("Ошибка! Аргумент template должен быть строкой.")
            self._init_empty_report()
            return

        template_path = Path(template)

        # Проверка наличия файла и его расширения
        if template_path.is_file():
            if template_path.suffix.lower() in [".docx", ".doc"]:
                self.doc = Document(template)
            else:
                print(
                    f"Ошибка! Файл шаблона {template} должен иметь расширение .docx или .doc."
                )
                print("Отчет будет сохранен с файлом шаблона по умолчанию.")
                self._init_empty_report()
        else:
            print(f"Файл {template_path} не обнаружен.")
            print("Отчет будет сохранен с файлом шаблона по умолчанию.")
            self._init_empty_report()

    def remove_paragraph(self, index=-1) -> None:
        """Метод удаляет параграф из файла документа. По умолчанию удаляется текущий параграф.

        Args:
            index (int): Индекс удаляемого параграфа, -1 — текущий параграф, -2 предыдущий и т.д.
        По умолчанию: -1 (текущий)
        """
        # -1 будет текущий параграф, -2 будет предыдущий
        try:
            paragraph = self.doc.paragraphs[index]

            # Удаляем параграф из документа
            p = paragraph._element
            p.getparent().remove(p)
        except IndexError:
            print(
                f"Ошибка, не удалось удалить параграф с индексом {index}. "
                "Вероятно этот параграф не существует."
            )

    def save(self, filename: str) -> None:
        """Метод сохраняет отчёт в файл Microsoft Word по заданному пути.

        Args:
            filename (str): Путь к файлу в который сохранить отчёт
        """
        report_path = Path(filename)

        # Создаем папку до пути файла, если не существует
        report_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            self.doc.save(report_path)
            print("\n" + "-" * 74)
            print(f"Файл: {report_path} успешно сохранён!")
            print("-" * 74 + "\n")
        except PermissionError:
            print(
                "\nОшибка! Не удалось сохранить файл. "
                "Проверьте возможность записи файла по указанному пути."
            )
            print("Возможно записываемый файл уже существует и открыт.")

            self._handle_save_error(filename)

    def _handle_save_error(self, path: str) -> None:
        temp_path = (
            str(Path(path).parent) + f"/temp_{random.randrange(1, 10**3):03}.docx"
        )
        inp = input(f"\nСохранить во временный файл {temp_path} (y/n)? ")
        if inp.lower() == "y":
            self.save(temp_path)
        elif inp.lower() == "n":
            print("\nОтчет не сохранен.")
            return
        else:
            self._handle_save_error(path)


def set_table_columns_width(table, col_widths: tuple) -> None:
    """Функция устанавливает ширину столбцов в таблице docx
    поочередно проходя по каждой ячейке таблицы.

    Args:
        table (docx.table.Table): Таблица в документе docx.
        col_widths (tuple): Кортеж, содержащий ширину столбцов
    в сантиметрах по порядку.
    """
    # Получаем доступ к ячейкам таблицы из соображений производительности
    # и считаем количество колонок и строк
    cells = table._cells
    columns = len(table.columns)
    rows = len(table.rows)

    if columns != len(col_widths):
        print(
            "\nВнимание количество заданных столбцов "
            "не совпадает с количеством столбцов в таблице."
        )
        print(f"В таблице — {columns}, задано — {len(col_widths)}")

    for row_idx in range(rows):
        for column_idx in range(columns):
            # Номер ячейки по порядку
            cell_n = column_idx + row_idx * columns

            # Устанавливаем ширину столбцов.
            # Если ячеек заданных ширин меньше чем столбцов
            # устанавливаем последнее успешное значение
            try:
                success_width = Cm(col_widths[column_idx])
                cells[cell_n].width = success_width

            except IndexError:
                cells[cell_n].width = success_width


def set_table_style(
    table, style="Т-таблица", first_row_style="Т-таблица-заголовок"
) -> None:
    """Функция проходит по всем ячейкам таблицы table и устанавливает
    заданный стиль style параграфов в таблице. При желании можно
    указать стиль для заголовков таблицы (1-ая строка).

    Args:
        table (docx.table.Table): Таблица в документе docx
        style (str): Название устанавливаемого стиля. Стиль должен присутствовать
        в файле шаблона отчета. По умолчанию "Т-таблица".
        first_row_style (str): Названия стиля для первой строки таблицы.
        (строка заголовков). По умолчанию None.
    """
    # Получаем доступ к ячейкам таблицы из соображений производительности
    # и считаем количество колонок и строк
    cells = table._cells
    columns = len(table.columns)
    rows = len(table.rows)

    for row_idx in range(rows):
        for column_idx in range(columns):
            # Номер ячейки по порядку
            cell_n = column_idx + row_idx * columns

            # Если задан стиль первой строки устанавливаем
            # иначе проходим по всем ячейкам и устанавливаем
            # основной стиль таблицы
            for paragraph in cells[cell_n].paragraphs:
                paragraph.style = style

            if row_idx == 0 and first_row_style:
                for paragraph in cells[cell_n].paragraphs:
                    try:
                        paragraph.style = first_row_style
                    except TypeError:
                        print(
                            "Ошибка установки стиля первой строки, "
                            f"заголовок: {paragraph.text}, стиль: {first_row_style}"
                        )


def set_table_rows_style(table, rows=(0, 1), style="Т-таблица-заголовок") -> None:
    """Метод предназначен для установки заданного стиля для указанных строк таблицы.

    Args:
        table (docx.table.Table): Таблица в документе docx
        style (str): Название устанавливаемого стиля. Стиль должен присутствовать
        в файле шаблона отчета. По умолчанию "Т-таблица-заголовок".
    """
    cells = table._cells
    columns = len(table.columns)

    for row_idx in rows:
        for column_idx in range(columns):
            # Номер ячейки по порядку
            cell_n = column_idx + row_idx * columns

            # Если задан стиль первой строки устанавливаем
            # иначе проходим по всем ячейкам и устанавливаем
            # основной стиль таблицы
            for paragraph in cells[cell_n].paragraphs:
                paragraph.style = style
