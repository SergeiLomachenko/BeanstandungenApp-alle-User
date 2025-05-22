import os
import subprocess
import sys
from flask import Flask, render_template, request, redirect, flash, url_for, send_file

app = Flask(__name__)
app.secret_key = 'dein_geheimer_schluessel' 

# Убираем загрузку конфигурации, так как файл pdf4.py используется для логики, а не настроек
# app.config.from_pyfile(os.path.join(os.path.dirname(__file__), 'pdf4.py'))

# Определяем базовую директорию проекта (где хранятся входные и выходные файлы)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Получаем файлы из формы
        invoice_file = request.files.get('invoice')
        ca3_file = request.files.get('ca3')
        rrm_file = request.files.get('rrm')
        
        # Проверка, что все необходимые файлы загружены
        if not (invoice_file and ca3_file and rrm_file):
            flash("Bitte laden Sie alle erforderlichen Dateien hoch.")
            return redirect(request.url)
        
        # Сохраняем загруженные файлы в BASE_DIR
        invoice_path = os.path.join(BASE_DIR, 'invoice.pdf')
        ca3_path = os.path.join(BASE_DIR, 'ca3.xlsx')
        rrm_path = os.path.join(BASE_DIR, 'rrm.xlsx')
        
        invoice_file.save(invoice_path)
        ca3_file.save(ca3_path)
        rrm_file.save(rrm_path)
        
        # Вызываем функцию анализа, которая выполнит pdf4.py и обработает файлы
        run_analysis()
        
        # Удаляем исходные файлы, так как они больше не нужны
        for f in ['invoice.pdf', 'ca3.xlsx', 'rrm.xlsx']:
            path = os.path.join(BASE_DIR, f)
            if os.path.exists(path):
                os.remove(path)
        
        # Перенаправляем пользователя на страницу загрузок
        return redirect(url_for("download_page"))
    
    return render_template("index.html")


def run_analysis():
    """
    Функция run_analysis() выполняет скрипт pdf4.py, который ожидается создать файлы:
        file1.xlsx, file2.xlsx, file3.xlsx, file4.xlsx.
    После выполнения эти файлы переименовываются в:
        Gesamtinvoiceinfo.xlsx, Invoiceinfo.xlsx, Rechnungsprüfung.xlsx, Validierung.xlsx,
    и остаются в BASE_DIR для последующего скачивания.
    """
    # Имена файлов, создаваемых pdf4.py
    original_files = ['file1.xlsx', 'file2.xlsx', 'file3.xlsx', 'file4.xlsx']
    # Новые имена для этих файлов
    new_names = [
        'Gesamtinvoiceinfo.xlsx', 
        'Invoiceinfo.xlsx', 
        'Rechnungsprüfung.xlsx', 
        'Validierung.xlsx'
    ]
    
    try:
        result = subprocess.run(
            [sys.executable, 'pdf4.py'],
            cwd=BASE_DIR,
            capture_output=True,
            text=True
        )
        print("stdout:\n", result.stdout)
        print("stderr:\n", result.stderr)
        print("Return code:", result.returncode)
    except Exception as ex:
        print("Fehler beim Ausführen von pdf4.py:", ex)
    
    # Переименование и перемещение файлов
    for original, new_name in zip(original_files, new_names):
        src = os.path.join(BASE_DIR, original)
        dst = os.path.join(BASE_DIR, new_name)
        if os.path.exists(src):
            try:
                os.replace(src, dst)
                print(f"Datei {original} wurde in {new_name} umbenannt.")
            except Exception as e:
                print(f"Fehler beim Umbenennen der Datei {original}: {e}")
        else:
            print(f"Ausgabedatei {original} wurde nicht gefunden.")


@app.route("/downloads")
def download_page():
    # Список файлов для скачивания
    results = [
        'Gesamtinvoiceinfo.xlsx', 
        'Invoiceinfo.xlsx', 
        'Rechnungsprüfung.xlsx', 
        'Validierung.xlsx'
    ]
    return render_template("download.html", results=results)


@app.route("/download/<filename>")
def download_file(filename):
    file_path = os.path.join(BASE_DIR, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "Datei nicht gefunden", 404


if __name__ == '__main__':
    app.run(debug=True)
