import csv


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


class Formatter:
    def __init__(self, headers: list[str], rows: list[list[any]], path_to_file_for_save_without_extension: str,
                 delimiter: str = ";"):
        self.headers = headers
        self.rows = rows
        self.filename = path_to_file_for_save_without_extension
        self.delimiter = delimiter

    def to_csv(self) -> str:  # path to file
        with open(f"{self.filename}.csv", 'w', encoding="utf-8") as f:
            write = csv.writer(f, delimiter=self.delimiter, )
            write.writerow(self.headers)
            for row in self.rows:
                row = [str(i) for i in row]
                write.writerow(row)

        return f"{self.filename}.csv"

    def to_excel(self) -> str:  # path to file
        try:
            import xlsxwriter
        except ImportError:
            raise RuntimeError("To use this method, you need to install the xlsxwriter library. Run the command: pip install xlsxwriter")
	
        workbook = xlsxwriter.Workbook(f'{self.filename}.xlsx')
        worksheet = workbook.add_worksheet()

        row_number = 0

        for index, header in enumerate(self.headers):
            worksheet.write(row_number, index, header)

        row_number += 1

        column_length = {}

        for row in self.rows:
            for index, value in enumerate(row):
                current_len = column_length.get(index, 0)
                column_length[index] = max(current_len, len(str(value)))

                if is_int(value=value):
                    number_format = workbook.add_format({'num_format': '0'})
                    worksheet.write(row_number, index, try_to_int(value=value), number_format)
                else:
                    worksheet.write(row_number, index, try_to_int(value=value))
            row_number += 1

        for cl in column_length:
            worksheet.set_column(cl, cl, column_length[cl] + 5)

        workbook.close()

        return f"{self.filename}.xlsx"

    def to_html_table(self) -> str:  # path to file
        table_headers = '\n'.join([f'<th onclick="sortTable({i})">{el}</th>' for i,el in enumerate(self.headers)])
        table_rows = ""
        for row in self.rows:
            table_row = ""
            for element in row:
                table_row += f"<td>{element}</td>\n"
            table_rows += f"<tr>{table_row}</tr>\n"

        with open(f"{self.filename}.html", 'w', encoding='utf-8') as f:
            html_tamplate = """<!DOCTYPE html>
<html lang="ru">
<head>
	<meta charset="UTF-8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<title>Document</title>
<style>
td,th{border:1px solid #e6c200}.popup,table{white-space:nowrap}.popup,body,main{color:#fff}body{background-color:#1e1e2f}table{margin:0 auto;border-collapse:collapse;box-shadow:0 0 10px #e6c200}th{position:sticky;top:0;background-color:#e6c200;color:#1e1e2f;padding:5px;cursor:pointer}td{padding:10px;text-align:center}tr:nth-child(2n){background-color:#2a2a3b}tr:nth-child(odd){background-color:#3a3a4f}tr:hover{background-color:#55557a}.popup{position:fixed;top:50px;right:25px;transform:translate(-50%,-50%);background-color:#3cc;padding:10px 20px;border-radius:5px;display:none;z-index:1;font-size:20px}
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
</body>
</html>"""
            f.write(html_tamplate)

        return f"{self.filename}.html"
