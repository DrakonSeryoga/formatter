import csv
from dataclasses import dataclass
from enum import Enum
from typing import Optional, Union, Iterable


def try_to_int(value: any):
    try:
        return float(value)
    except:
        try:
            return int(value)
        except:
            return str(value)


def is_int(value: any):
    try:
        int(value)
        return True
    except:
        return False


class Color(Enum):
    RED: str = '#FF0000'
    GREEN: str = '#00FF00'
    BLUE: str = '#0000FF'


@dataclass
class Url:
    url: str
    value: Union[str, float, int, bool]

    def __str__(self):
        return f'<a style="color: inherit" href="{self.url}">{self.value}</a>'


@dataclass
class RowValue:
    value: Union[str, int, bool, float, Url]
    color: Optional[Union[Color, str]] = None
    tooltip: Optional[Union[str, int, bool, float]] = None
    tooltip_color: Optional[Union[Color, str]] = None


class TableRow:
    def __init__(self):
        self.rows: Optional[list[RowValue]] = []

    def add(self, *rows: RowValue) -> 'TableRow':
        if isinstance(rows, Iterable):
            for i in rows:
                self.rows.append(i)
        else:
            self.rows.append(*rows)

        return self

    def __repr__(self):
        return f"{self.rows}"


class Formatter:
    def __init__(self, headers: TableRow, rows: list[TableRow], path_to_file_for_save_without_extension: str,
                 delimiter: str = ";"):
        self.headers: TableRow = headers
        self.rows: list[TableRow] = rows
        self.filename = path_to_file_for_save_without_extension
        self.delimiter = delimiter

    def to_csv(self) -> str:  # path to file
        with open(f"{self.filename}.csv", 'w', encoding="utf-8") as f:
            write = csv.writer(f, delimiter=self.delimiter, )
            write.writerow([r.value for r in self.headers.rows])
            for row in self.rows:
                row = [str(i.value) for i in row.rows]
                write.writerow(row)

        return f"{self.filename}.csv"

    def to_excel(self) -> str:  # path to file
        try:
            import xlsxwriter
        except ImportError:
            raise RuntimeError(
                "To use this method, you need to install the xlsxwriter library. Run the command: pip install XlsxWriter==3.2.0")

        workbook = xlsxwriter.Workbook(f'{self.filename}.xlsx')
        worksheet = workbook.add_worksheet()

        row_number = 0

        for index, header in enumerate(self.headers.rows):
            worksheet.write(row_number, index, header.value)

        row_number += 1

        column_length = {}

        for row in self.rows:
            for index, value in enumerate(row.rows):
                current_len = column_length.get(index, 0)
                column_length[index] = max(current_len, len(str(value.value)))

                if is_int(value=value.value):
                    number_format = workbook.add_format({'num_format': '0'})
                    worksheet.write(row_number, index, try_to_int(value=value.value), number_format)
                else:
                    worksheet.write(row_number, index, try_to_int(value=value.value))
            row_number += 1

        for cl in column_length:
            worksheet.set_column(cl, cl, column_length[cl] + 5)

        workbook.close()

        return f"{self.filename}.xlsx"

    def to_html_table(self) -> str:  # path to file
        table_headers = '\n'.join([
                                      f'<th style="color: {'#fff' if el.color is None else el.color if isinstance(el.color, str) else el.color.value}" onclick="sortTable({i})"><p id="H_{i}">{el.value}</p></th>'
                                      for i, el in enumerate(self.headers.rows)])
        table_rows = ""
        tooltip_js_strings = ""
        for i, header_element in enumerate(self.headers.rows):
            if not header_element.tooltip:
                continue
            tooltip_js_strings += f"""tippy('#H_{i}', """ + """{
            interactive: true,
            content: """ + f""" '<span style="color: {'#fff' if header_element.tooltip_color is None else header_element.tooltip_color if isinstance(header_element.tooltip_color, str) else header_element.tooltip_color.value};">{header_element.tooltip}</span>' """ + """,
            allowHTML: true,
            placement: 'bottom'
        });\n"""
        for i, row in enumerate(self.rows):
            table_row = ""
            for _i, element in enumerate(row.rows):
                if element.tooltip is not None:
                    tooltip_js_strings += f"""tippy('#R_{i}_{_i}', """ + """{
    interactive: true,
    content: """ + f""" `<span style="color: {'#fff' if element.tooltip_color is None else (element.tooltip_color if isinstance(element.tooltip_color, str) else element.tooltip_color.value)};">{element.tooltip}</span>` """ + """,
    allowHTML: true,
    placement: 'top'
});\n"""
                table_row += f'<td><p style="color: {'#fff' if element.color is None else (element.color if isinstance(element.color, str) else element.color.value)}" id="R_{i}_{_i}">{element.value}</p></td>\n'
            table_rows += f"<tr>{table_row}</tr>\n"

        with open(f"{self.filename}.html", 'w', encoding='utf-8') as f:
            html_tamplate = """<!DOCTYPE html>
<html lang="ru">
<head>
	<meta charset="UTF-8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<title>Таблица</title>
	<link rel="stylesheet" href="https://unpkg.com/tippy.js@6/animations/scale.css"/>
<style>
td,th{border:1px solid #e6c200}p{margin:0}a{color:#fff}.popup,table{white-space:nowrap}.popup,body,main{color:#fff}body{background-color:#1e1e2f}table{margin:0 auto;border-collapse:collapse;box-shadow:0 0 10px #e6c200}th{position:sticky;top:0;background-color:#e6c200;color:#1e1e2f;padding:5px;cursor:pointer}td{padding:10px;text-align:center}tr:nth-child(2n){background-color:#2a2a3b}tr:nth-child(odd){background-color:#3a3a4f}tr:hover{background-color:#55557a}.popup{position:fixed;top:50px;right:25px;transform:translate(-50%,-50%);background-color:#3cc;padding:10px 20px;border-radius:5px;display:none;z-index:1;font-size:20px}
</style>
</head>
<body>
	<main>
		<table id="sortableTable">
			<thead>
				  <tr>
		            """ + table_headers + """
				  </tr>
			</thead>
			<tbody>
            	""" + table_rows + """
            </tbody>
		</table>
	</main>	
	<div class="popup" id="popup">Текст скопирован</div>
<script src="https://unpkg.com/@popperjs/core@2/dist/umd/popper.min.js"></script>
<script src="https://unpkg.com/tippy.js@6/dist/tippy-bundle.umd.js"></script>
<script>
    let sortDirection = {};
    function sortTable(colIndex) {
        const table = document.getElementById("sortableTable");
        const tbody = table.querySelector("tbody");
        const rows = Array.from(tbody.rows);

        const isAscending = sortDirection[colIndex] = !sortDirection[colIndex];

        rows.sort((rowA, rowB) => {
            let cellA = rowA.cells[colIndex].innerText.trim();
            let cellB = rowB.cells[colIndex].innerText.trim();

            if (!isNaN(cellA) && !isNaN(cellB)) {
                return isAscending ? cellA - cellB : cellB - cellA;
            }
            return isAscending ? cellA.localeCompare(cellB) : cellB.localeCompare(cellA);
        });

        rows.forEach(row => tbody.appendChild(row));
    }

	const popup = document.getElementById("popup");

	document.querySelectorAll("td").forEach(td => {
		td.addEventListener("click", e => {
			navigator.clipboard.writeText(e.target.innerText)
				.then(() => {
					showPopup();
				})
				.catch(() => {
					alert("Что-то пошло не так");
				});
		});
	});

	function showPopup() {
		popup.style.display = "block";
		setTimeout(() => {
			popup.style.display = "none";
		}, 2000); // всплывашка исчезнет через 2 секунды
	}
</script>
""" + f"""<script>
{tooltip_js_strings}</script>
</body>
</html>"""
            f.write(html_tamplate)

        return f"{self.filename}.html"
