from flask import render_template, Flask, request, send_file
from app_utils import urltodf as udf


app = Flask(__name__)


@app.route("/home")
def upload():
    return render_template("urls_form.html")


@app.route("/success", methods=["POST"])
def success():
    success.max_tables = int(request.form['max_tables'])
    success.min_rows = int(request.form['min_rows'])
    success.urls = request.form['urls']
    success.excel = request.form['excel']

    return render_template("success.html", MAX_TABLES=success.max_tables,
                           MIN_ROWS=success.min_rows, URLS=success.urls, EXCEL=success.excel)


@app.route("/convert")
def converter():
    #pdfsplitter.cropper(success.start_page,success.end_page, success.file_name)
    excel_path = "Excel_Multiple_Sheets_Output.xlsx"
    urls = success.urls
    urls = urls.split(",")
    udf.urls_to_excel(urls, get_max=success.max_tables,
                      min_rows=success.min_rows, excel_path=excel_path)
    return render_template("download_files.html")


@app.route("/download")
def download():
    filename = "Excel_Multiple_Sheets_Output.xlsx"

    return send_file(filename, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
