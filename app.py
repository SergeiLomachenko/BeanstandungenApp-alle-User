import os
import subprocess
import sys
from flask import Flask, render_template, request, redirect, flash, url_for, send_file
import tempfile
import shutil

app = Flask(__name__)
app.secret_key = 'dein_geheimer_schluessel'

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
LOG_FILE   = os.path.join(BASE_DIR, 'analysis.log')

os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)


def log(message: str):
    """Schreibt einen Zeilen-Eintrag in analysis.log."""
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(message + '\n')


def run_analysis_in_temp_dir(month: str, recl_file_path: str, grp_file_path: str, temp_dir: str) -> str | None:
    """Führt die Analyse in einem temporären Verzeichnis durch."""
    script = os.path.join(BASE_DIR, 'Reads_excel_columns.py')
    log(f"\n=== Analyse starten für Monat {month} ===")
    log(f"Temporäres Verzeichnis: {temp_dir}")

    # Kopiere die hochgeladenen Dateien ins temporäre Verzeichnis
    temp_recl = os.path.join(temp_dir, 'recl.xlsx')
    temp_grp = os.path.join(temp_dir, 'grp.xlsx')
    
    shutil.copy2(recl_file_path, temp_recl)
    if grp_file_path and os.path.exists(grp_file_path):
        shutil.copy2(grp_file_path, temp_grp)

    # Subprocess aufrufen
    cmd = [sys.executable, script, month, '--recl', temp_recl]
    if grp_file_path and os.path.exists(grp_file_path):
        cmd.extend(['--grp', temp_grp])
    
    try:
        result = subprocess.run(
            cmd,
            cwd=temp_dir,
            capture_output=True,
            text=True
        )
        log("--- stdout ---\n" + result.stdout)
        log("--- stderr ---\n" + result.stderr)
        log(f"Return code: {result.returncode}")
        
        if result.returncode != 0:
            log("Script beendet mit Fehler!")
            return None
            
    except Exception as e:
        log("Subprocess-Fehler: " + str(e))
        return None

    # Ergebnisdatei finden (jetzt nur eine Datei mit zwei Registerkarten)
    result_filename = f'Ergebnis_{month}.xlsx'
    src_result = os.path.join(temp_dir, result_filename)
    
    if os.path.exists(src_result):
        log(f"Ergebnisdatei erstellt: {src_result}")
        return result_filename
    else:
        # Falls die Datei unter anderem Namen erstellt wurde
        src_main = os.path.join(temp_dir, 'file2_filtered.xlsx')
        if os.path.exists(src_main):
            dst_result = os.path.join(temp_dir, result_filename)
            try:
                os.rename(src_main, dst_result)
                log(f"Datei umbenannt: {dst_result}")
                return result_filename
            except Exception as e:
                log("Fehler beim Umbenennen der Datei: " + str(e))
                return None
    
    log("Keine Ergebnisdatei gefunden")
    return None


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        month     = request.form.get('month')
        recl_file = request.files.get('recl')
        grp_file  = request.files.get('grp')

        if not month or not recl_file:
            flash("Monat und Excel-Datei (recl) sind Pflicht.")
            return redirect(request.url)

        # Erstelle ein temporäres Verzeichnis für diese Anfrage
        temp_dir = tempfile.mkdtemp(prefix=f'analysis_{month}_')
        
        try:
            # Speichere hochgeladene Dateien temporär
            recl_path = os.path.join(temp_dir, 'upload_recl.xlsx')
            recl_file.save(recl_path)
            
            grp_path = None
            if grp_file and grp_file.filename:
                grp_path = os.path.join(temp_dir, 'upload_grp.xlsx')
                grp_file.save(grp_path)

            # Führe Analyse durch
            result_filename = run_analysis_in_temp_dir(month, recl_path, grp_path, temp_dir)
            
            if not result_filename:
                flash("Analyse fehlgeschlagen. Schau in analysis.log.")
                # Lösche temporäres Verzeichnis
                shutil.rmtree(temp_dir)
                return redirect(request.url)

            # Sende Ergebnisdatei
            result_path = os.path.join(temp_dir, result_filename)
            if os.path.exists(result_path):
                response = send_file(
                    result_path, 
                    as_attachment=True, 
                    download_name=result_filename
                )
                
                # Lösche temporäres Verzeichnis nach dem Senden
                try:
                    shutil.rmtree(temp_dir)
                    log(f"Temporäres Verzeichnis gelöscht: {temp_dir}")
                except Exception as e:
                    log(f"Fehler beim Löschen des temporären Verzeichnisses: {e}")
                
                return response
            else:
                flash("Ergebnisdatei nicht gefunden.")
                shutil.rmtree(temp_dir)
                return redirect(request.url)
                
        except Exception as e:
            # Im Fehlerfall temporäres Verzeichnis löschen
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
            log(f"Fehler bei der Verarbeitung: {e}")
            flash("Ein Fehler ist aufgetreten. Schau in analysis.log.")
            return redirect(request.url)

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)