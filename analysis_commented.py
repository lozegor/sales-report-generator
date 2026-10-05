from pathlib import Path
import argparse
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, Reference

# Читаем и сохраняем файлы рядом со скриптом.
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_FILE = BASE_DIR / "sales_report.xlsx"

INPUT_DIR = BASE_DIR / "input"

# По умолчанию оформляем A:C; ws и cell_border должны существовать к моменту вызова.
def add_borders(ws,start_row, end_row, border, start_col=1, end_col=3):
    # В iter_rows обе границы диапазона включаются.
    rows = ws.iter_rows(
        min_row=start_row,
        max_row=end_row,
        min_col=start_col,
        max_col=end_col
    )
    for row in rows:
        for cell in row:
            cell.border = border

def validate_data(df):
    if df.empty:
        raise ValueError("В исходном файле нет строк с продажами")

    required_columns = ["Менеджер", "Товар", "Количество", "Цена"]

    for column in required_columns:
        if column not in df.columns:
            raise ValueError(f"В исходном файле отсутствует столбец: {column}")

    for column in required_columns:
        if df[column].isna().any():
            raise ValueError(f"В столбце «{column}» есть пустые значения")
    
    for column in ["Менеджер", "Товар"]:
        df[column] = df[column].str.strip()

        if (df[column] == "").any():
            raise ValueError(
                f"В столбце «{column}» есть значение из одних пробелов"
            )
        
    for column in ["Количество", "Цена"]:
        try:
            df[column] = pd.to_numeric(df[column], errors="raise")
        except ValueError:
            raise ValueError(
                f"В столбце «{column}» найдено значение, которое нельзя преобразовать в число"
            ) from None
    for column in ["Количество", "Цена"]:
        if (df[column] <= 0).any():
            raise ValueError(
                f"В столбце «{column}» есть число, меньшее или равное нулю"
            )
    return df

def load_sales(input_dir):
    files = []

    for file in sorted(input_dir.glob("*.xlsx")):
        if file.is_file() and not file.name.startswith("~$"):
            files.append(file)
    if not files:
        raise ValueError("В папке input нет Excel-файлов")

    tables = []

    for file in files:
        table = pd.read_excel(file)

        try:
            table = validate_data(table)
        except ValueError as error:
            raise ValueError(
                f"Ошибка в файле «{file.name}»: {error}"
            ) from None
        table["Источник"] = file.name
        tables.append(table)
    return pd.concat(tables, ignore_index=True)
def main(input_dir, output_file):
    df = load_sales(input_dir)

    header_fill = PatternFill(#заливка
        fill_type="solid",
        fgColor="D9EAF7"
    )

    thin_side = Side(style="thin")

    cell_border = Border(
        left=thin_side,
        right=thin_side,
        top=thin_side,
        bottom=thin_side
    )


    df["Выручка"] = df["Количество"] * df["Цена"]#добавляем столбец выручки

    total_sales = df["Выручка"].sum()#считаем общую выручку

    if total_sales == 0:
        raise ValueError(
            "Общая выручка равна нулю: невозможно рассчитать доли"
        )

    source_report = df.groupby("Источник").agg(
        Выручка=("Выручка", "sum"),
        Количество=("Количество", "sum")
    )

    source_report = source_report.sort_values(
        by="Выручка",
        ascending=False
    )

    source_report = source_report.reset_index()
    source_report.insert(0, "№", range(1, len(source_report) + 1))


    sales_by_manedger = df.groupby("Менеджер")["Выручка"].sum()#считаем выручку каждого менеджера

    sales_by_manedger = sales_by_manedger.sort_values(ascending=False)

    best_manager = sales_by_manedger.idxmax()#Находим имя лучшего менедера
    best_manager_sales = sales_by_manedger.max()#И его продажи



    report = pd.DataFrame({
        "Показатель": [
            "Общая выручка",
            "Самый результативный менеджер",
            "Выручка менеджера",
            "Количество исходных файлов",
            "Количество строк продаж"
        ],
        "Значение": [
            total_sales,
            best_manager,
            best_manager_sales,
            len(source_report),
            len(df)
        ]    
    })

    report.insert(0,"№",range(1,len(report)+1))

    product_report = df.groupby("Товар").agg(
        Выручка=("Выручка", "sum"),
        Количество=("Количество", "sum")
    )

    product_report = product_report.sort_values(
        by="Выручка",
        ascending=False
    )

    product_report = product_report.reset_index()

    product_report["Доля"] = product_report["Выручка"] / total_sales

    product_report.insert(0,"№", range(1,len(product_report)+1))

    # reset_index() превращает имена менеджеров из индекса в обычный столбец.
    manager_report = sales_by_manedger.reset_index()
    manager_report["Доля"] = manager_report["Выручка"] / total_sales
    manager_report.insert(0,"№", range(1, len(manager_report)+1))# делаем правильную нумерацию
    manager_startrow = len(report) + 3

    product_startrow = (
        manager_startrow
        + len(manager_report)
        + 2
    )
    # startrow в pandas считается с 0: 6 означает строку 7 в Excel.
    # +2 учитывает заголовок менеджеров и одну пустую строку между таблицами.

    source_startrow = product_startrow + len(product_report) + 2
    
    with pd.ExcelWriter(output_file) as writer:#добавляем все в excel
        df.to_excel(writer, sheet_name="Данные", index=False)
        report.to_excel(writer, sheet_name="Отчёт", index=False)
        manager_report.to_excel(writer,sheet_name="Отчёт",index=False,startrow=manager_startrow)
        product_report.to_excel(writer,sheet_name="Отчёт",index=False,startrow=product_startrow)
        source_report.to_excel(
            writer,
            sheet_name="Отчёт",
            index=False,
            startrow=source_startrow
        )
    # После выхода из with файл записан и закрыт; теперь можно заняться оформлением.
    wb = load_workbook(output_file)
    ws = wb["Отчёт"]

    for cell in ws[1]:
        cell.font = Font(bold = True)

    for row in ws.iter_rows(min_row=manager_startrow+1, max_row=manager_startrow+1, min_col=1, max_col=4):
        for cell in row:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
            cell.fill = header_fill
            cell.border = cell_border

    # openpyxl считает строки с 1: заголовок товаров находится на startrow + 1.
    for cell in ws[product_startrow + 1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.fill = header_fill

    # range не включает правую границу; формат меняет отображение, а не само число.
    for row in range(manager_startrow + 2, manager_startrow + 2 + len(manager_report)):
        ws[f"C{row}"].number_format = '#,##0'
        ws[f"D{row}"].number_format = '0%'

    add_borders(ws, manager_startrow + 1, manager_startrow + 1 + len(manager_report), border=cell_border, end_col=4)

    for row in range(
        product_startrow + 2,
        product_startrow + 2 + len(product_report)
    ):
        ws[f"C{row}"].number_format = '#,##0'
        ws[f"D{row}"].number_format = '0'
        ws[f"E{row}"].number_format = '0%'

    add_borders(ws,product_startrow+1,product_startrow+1+len(product_report),border=cell_border, end_col=5)
    
    for cell in ws[source_startrow + 1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.fill = header_fill

    for row in range(
        source_startrow + 2,
        source_startrow + 2 + len(source_report)
    ):
        ws[f"C{row}"].number_format = "#,##0"
        ws[f"D{row}"].number_format = "0"
        ws[f"A{row}"].alignment = Alignment(horizontal="center")

    add_borders(
        ws,
        source_startrow + 1,
        source_startrow + 1 + len(source_report),
        border=cell_border,
        end_col=4
    )
        
    # Приблизительная ширина: длина самого длинного значения + небольшой запас.
    for column in ws.columns:
        max_length = 0

        for cell in column:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[column[0].column_letter].width = max_length + 2

    for row in range(manager_startrow + 2,manager_startrow + 2 + len(manager_report)):
        ws[f"A{row}"].alignment = Alignment(horizontal="center")

    for row in range(product_startrow+2,product_startrow+2+len(product_report)):
        ws[f"A{row}"].alignment = Alignment(horizontal="center")

    chart = BarChart()#Создание диаграмы
    chart.title = "Выручка по менеджерам"
    chart.y_axis.title = "Выручка"
    chart.x_axis.title = "Менеджер"

    data = Reference(#
        ws,
        min_col=3,
        min_row=manager_startrow + 1,
        max_row=manager_startrow + 1 + len(manager_report)
    )

    chart.add_data(data, titles_from_data=True)#

    labels = Reference(#
        ws,
        min_col=2,
        min_row=manager_startrow + 2,
        max_row=manager_startrow + 1 + len(manager_report)
    )

    data2 = Reference(
        ws,
        min_col=3,
        min_row=product_startrow + 1,
        max_row=product_startrow + 1 + len(product_report)
    )

    labels2 = Reference(
        ws,
        min_col=2,
        min_row=product_startrow + 2,
        max_row=product_startrow + 1 + len(product_report)
    )

    chart.set_categories(labels)#

    ws.add_chart(chart, "G7")

    chart2 = BarChart()#Создание диаграмы
    chart2.title = "Выручка по товарам"
    chart2.y_axis.title = "Выручка"
    chart2.x_axis.title = "Товар"

    ws.add_chart(chart2, "M7")
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(labels2)
    data_ws = wb["Данные"]
    data_ws.auto_filter.ref = data_ws.dimensions
    data_ws.freeze_panes = "A2"
    wb.save(output_file)
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Создание отчёта по продажам"
    )

    parser.add_argument(
        "--input",
        type=Path,
        default=INPUT_DIR,
        help="Папка с исходными Excel-файлами"
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_FILE,
        help="Путь к готовому Excel-отчёту"
    )
    args = parser.parse_args()
    main(args.input, args.output)

































