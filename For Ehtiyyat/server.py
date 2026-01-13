from flask import Flask, request, send_file, jsonify
from werkzeug.utils import secure_filename
import os
import shutil
from datetime import datetime
import zipfile
import threading
import time

# Bütün forma8 import-ları
from sheets.forma8_1 import run_forma8_1
from sheets.forma8_2 import run_forma8_2
from sheets.forma8_3 import run_forma8_3
from sheets.forma8_4 import run_forma8_4
from sheets.forma8_5 import run_forma8_5
from sheets.forma8_6 import run_forma8_6
from sheets.forma8_11 import run_forma8_11
from sheets.forma8_7 import run_forma8_7
from sheets.forma8_10 import run_forma8_10
from sheets.forma8_8 import run_forma8_8
from sheets.forma8_12 import run_forma8_12
from sheets.forma8_9 import run_forma8_9
from sheets.forma8_13 import run_forma8_13
from sheets.forma8_14 import run_forma8_14
from sheets.yekun_reserv import run_yekun_reserv

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB max

# Upload folderi yarat
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def home():
    return '''
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8">
            <title>Forma8 Avtomatlaşdırma</title>
            <style>
                * { margin: 0; padding: 0; box-sizing: border-box; }
                body { 
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    min-height: 100vh;
                    padding: 20px;
                }
                .container {
                    max-width: 800px;
                    margin: 0 auto;
                    background: white;
                    border-radius: 20px;
                    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
                    padding: 40px;
                }
                h1 {
                    color: #667eea;
                    margin-bottom: 10px;
                    font-size: 32px;
                    text-align: center;
                }
                .subtitle {
                    text-align: center;
                    color: #666;
                    margin-bottom: 30px;
                }
                .form-group {
                    margin: 25px 0;
                }
                label {
                    display: block;
                    margin-bottom: 8px;
                    font-weight: 600;
                    color: #333;
                    font-size: 14px;
                }
                input[type="file"] {
                    width: 100%;
                    padding: 12px;
                    border: 2px dashed #ddd;
                    border-radius: 8px;
                    cursor: pointer;
                    transition: all 0.3s;
                }
                input[type="file"]:hover {
                    border-color: #667eea;
                    background: #f8f9ff;
                }
                input[type="date"] {
                    width: 100%;
                    padding: 12px;
                    border: 2px solid #ddd;
                    border-radius: 8px;
                    font-size: 14px;
                }
                small {
                    color: #999;
                    font-size: 12px;
                    display: block;
                    margin-top: 5px;
                }
                button {
                    width: 100%;
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    color: white;
                    padding: 16px;
                    border: none;
                    border-radius: 8px;
                    cursor: pointer;
                    font-size: 16px;
                    font-weight: 600;
                    transition: transform 0.2s;
                    margin-top: 20px;
                }
                button:hover {
                    transform: translateY(-2px);
                    box-shadow: 0 10px 25px rgba(102, 126, 234, 0.4);
                }
                button:disabled {
                    background: #ccc;
                    cursor: not-allowed;
                    transform: none;
                }
                #status {
                    margin-top: 30px;
                    padding: 20px;
                    border-radius: 8px;
                    display: none;
                    animation: slideIn 0.3s;
                }
                @keyframes slideIn {
                    from { opacity: 0; transform: translateY(-10px); }
                    to { opacity: 1; transform: translateY(0); }
                }
                .success {
                    background: #d4edda;
                    color: #155724;
                    border: 2px solid #c3e6cb;
                }
                .error {
                    background: #f8d7da;
                    color: #721c24;
                    border: 2px solid #f5c6cb;
                }
                .processing {
                    background: #d1ecf1;
                    color: #0c5460;
                    border: 2px solid #bee5eb;
                }
                .progress-bar {
                    width: 100%;
                    height: 8px;
                    background: #e0e0e0;
                    border-radius: 10px;
                    overflow: hidden;
                    margin-top: 15px;
                }
                .progress-fill {
                    height: 100%;
                    background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
                    width: 0%;
                    transition: width 0.3s;
                    animation: pulse 2s infinite;
                }
                @keyframes pulse {
                    0%, 100% { opacity: 1; }
                    50% { opacity: 0.7; }
                }
                .status-icon {
                    font-size: 24px;
                    margin-right: 10px;
                }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>📊 Forma8 Avtomatlaşdırma</h1>
                <p class="subtitle">Excel fayllarını yükləyin və prosesi başladın</p>
                
                <form id="uploadForm" enctype="multipart/form-data">
                    <div class="form-group">
                        <label>📁 UCOT Faylı (UcotA.xlsx):</label>
                        <input type="file" name="ucot_file" accept=".xlsx,.xls" required>
                        <small>Əsas UCOT məlumat bazası</small>
                    </div>
                    
                    <div class="form-group">
                        <label>📁 Template Faylı (ALL.xlsx):</label>
                        <input type="file" name="template_file" accept=".xlsx,.xls" required>
                        <small>Bütün forma8 sheet-ləri olan template</small>
                    </div>
                    
                    <div class="form-group">
                        <label>📁 Əvvəlki Ay Faylları (ZIP):</label>
                        <input type="file" name="previous_files" accept=".zip">
                        <small>Forma8_7, 8_10, 8_14 üçün lazımdır (məcburi deyil)</small>
                    </div>
                    
                    <div class="form-group">
                        <label>📁 Yekun Reserv Template:</label>
                        <input type="file" name="yekun_template" accept=".xlsx,.xls">
                        <small>Yekun Reserv Excel faylı (məcburi deyil)</small>
                    </div>
                    
                    <div class="form-group">
                        <label>📅 Referans Tarixi:</label>
                        <input type="date" name="reference_date" required>
                        <small>Hesablamalar üçün əsas tarix</small>
                    </div>
                    
                    <button type="submit" id="submitBtn">🚀 Prosesi Başlat</button>
                </form>
                
                <div id="status"></div>
            </div>
            
            <script>
                document.getElementById('uploadForm').onsubmit = async (e) => {
                    e.preventDefault();
                    
                    const status = document.getElementById('status');
                    const submitBtn = document.getElementById('submitBtn');
                    
                    status.style.display = 'block';
                    status.className = 'processing';
                    status.innerHTML = `
                        <span class="status-icon">⏳</span>
                        <strong>Fayllar yüklənir və işlənir...</strong>
                        <div class="progress-bar">
                            <div class="progress-fill" id="progressFill"></div>
                        </div>
                        <p style="margin-top: 10px; font-size: 14px;">
                            Zəhmət olmasa gözləyin. Bu proses bir neçə dəqiqə çəkə bilər.
                        </p>
                    `;
                    
                    submitBtn.disabled = true;
                    submitBtn.textContent = '⏳ İşlənir...';
                    
                    // Progress bar simulyasiyası
                    let progress = 0;
                    const progressInterval = setInterval(() => {
                        progress += Math.random() * 10;
                        if (progress > 90) progress = 90;
                        document.getElementById('progressFill').style.width = progress + '%';
                    }, 500);
                    
                    const formData = new FormData(e.target);
                    
                    try {
                        const response = await fetch('/process', {
                            method: 'POST',
                            body: formData
                        });
                        
                        clearInterval(progressInterval);
                        document.getElementById('progressFill').style.width = '100%';
                        
                        if (response.ok) {
                            const blob = await response.blob();
                            const url = window.URL.createObjectURL(blob);
                            const a = document.createElement('a');
                            a.href = url;
                            a.download = 'Forma8_Results_' + Date.now() + '.zip';
                            document.body.appendChild(a);
                            a.click();
                            document.body.removeChild(a);
                            window.URL.revokeObjectURL(url);
                            
                            status.className = 'success';
                            status.innerHTML = `
                                <span class="status-icon">✅</span>
                                <strong>Proses uğurla tamamlandı!</strong>
                                <p style="margin-top: 10px;">Nəticə faylları yüklənir...</p>
                            `;
                            
                            submitBtn.disabled = false;
                            submitBtn.textContent = '🚀 Prosesi Başlat';
                        } else {
                            const error = await response.text();
                            status.className = 'error';
                            status.innerHTML = `
                                <span class="status-icon">❌</span>
                                <strong>Xəta baş verdi:</strong>
                                <p style="margin-top: 10px;">${error}</p>
                            `;
                            
                            submitBtn.disabled = false;
                            submitBtn.textContent = '🚀 Prosesi Başlat';
                        }
                    } catch (error) {
                        clearInterval(progressInterval);
                        status.className = 'error';
                        status.innerHTML = `
                            <span class="status-icon">❌</span>
                            <strong>Əlaqə xətası:</strong>
                            <p style="margin-top: 10px;">${error.message}</p>
                        `;
                        
                        submitBtn.disabled = false;
                        submitBtn.textContent = '🚀 Prosesi Başlat';
                    }
                };
                
                // Bugünkü tarixi default olaraq qoy
                document.querySelector('input[type="date"]').valueAsDate = new Date();
            </script>
        </body>
    </html>
    '''

@app.route('/process', methods=['POST'])
def process():
    session_id = None
    try:
        # Faylları yüklə
        ucot_file = request.files.get('ucot_file')
        template_file = request.files.get('template_file')
        previous_files = request.files.get('previous_files')
        yekun_template = request.files.get('yekun_template')
        reference_date = request.form.get('reference_date')
        
        if not ucot_file or not template_file or not reference_date:
            return "UCOT faylı, Template faylı və Tarix məcburidir!", 400
        
        # Unikal session ID yarat
        session_id = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        os.makedirs(session_folder, exist_ok=True)
        
        print(f"\n{'='*60}")
        print(f"Session: {session_id}")
        print(f"Tarix: {reference_date}")
        print(f"{'='*60}\n")
        
        # Faylları saxla
        ucot_path = os.path.join(session_folder, 'UcotA.xlsx')
        template_path = os.path.join(session_folder, 'ALL.xlsx')
        output_folder = os.path.join(session_folder, 'output')
        previous_folder = os.path.join(session_folder, 'previous')
        
        os.makedirs(output_folder, exist_ok=True)
        os.makedirs(previous_folder, exist_ok=True)
        
        ucot_file.save(ucot_path)
        template_file.save(template_path)
        
        # Əvvəlki faylları extract et
        if previous_files:
            zip_path = os.path.join(session_folder, 'previous.zip')
            previous_files.save(zip_path)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(previous_folder)
            print("✓ Previous faylları extract edildi")
        
        # Yekun Reserv template
        yekun_path = None
        if yekun_template:
            yekun_path = os.path.join(session_folder, 'Yekun_Reserv.xlsx')
            yekun_template.save(yekun_path)
            print("✓ Yekun Reserv template yükləndi")
        
        # ==================== FORMA8 PROSESLƏRİ ====================
        
        print("\n▶ Forma8_1 işlənir...")
        run_forma8_1(
            ucot_file=ucot_path,
            template_file=template_path,
            reference_date=reference_date,
            output_folder=output_folder
        )
        print("✅ Forma8_1 tamamlandı")
        
        excel_files = [f for f in os.listdir(output_folder) if f.endswith(".xlsx")]
        print(f"\n📊 {len(excel_files)} product faylı yaradıldı")
        
        # Forma8_2
        print("\n▶ Forma8_2 işlənir...")
        total_f_values = {}
        success_count = 0
        for file in excel_files:
            excel_path = os.path.join(output_folder, file)
            try:
                total_f = run_forma8_2(excel_path, ucot_path, reference_date)
                if total_f and total_f > 0:
                    total_f_values[file] = total_f
                success_count += 1
                print(f"  ✓ {file}")
            except Exception as e:
                print(f"  ✗ {file}: {str(e)}")
        print(f"✅ Forma8_2 tamamlandı ({success_count}/{len(excel_files)})")
        
        # Forma8_3
        print("\n▶ Forma8_3 işlənir...")
        success_count = 0
        for file in excel_files:
            try:
                run_forma8_3(os.path.join(output_folder, file), ucot_path, reference_date)
                success_count += 1
                print(f"  ✓ {file}")
            except Exception as e:
                print(f"  ✗ {file}: {str(e)}")
        print(f"✅ Forma8_3 tamamlandı ({success_count}/{len(excel_files)})")
        
        # Forma8_4
        print("\n▶ Forma8_4 işlənir...")
        success_count = 0
        for file in excel_files:
            try:
                run_forma8_4(os.path.join(output_folder, file), reference_date, ucot_path)
                success_count += 1
                print(f"  ✓ {file}")
            except Exception as e:
                print(f"  ✗ {file}: {str(e)}")
        print(f"✅ Forma8_4 tamamlandı ({success_count}/{len(excel_files)})")
        
        # Forma8_5
        print("\n▶ Forma8_5 işlənir...")
        success_count = 0
        for file in excel_files:
            try:
                run_forma8_5(os.path.join(output_folder, file), ucot_path, reference_date)
                success_count += 1
                print(f"  ✓ {file}")
            except Exception as e:
                print(f"  ✗ {file}: {str(e)}")
        print(f"✅ Forma8_5 tamamlandı ({success_count}/{len(excel_files)})")
        
        # Forma8_6
        print("\n▶ Forma8_6 işlənir...")
        success_count = 0
        for file in excel_files:
            try:
                run_forma8_6(os.path.join(output_folder, file), reference_date)
                success_count += 1
                print(f"  ✓ {file}")
            except Exception as e:
                print(f"  ✗ {file}: {str(e)}")
        print(f"✅ Forma8_6 tamamlandı ({success_count}/{len(excel_files)})")
        
        # Forma8_11
        print("\n▶ Forma8_11 işlənir...")
        success_count = 0
        for file in excel_files:
            try:
                run_forma8_11(os.path.join(output_folder, file), ucot_path, reference_date)
                success_count += 1
                print(f"  ✓ {file}")
            except Exception as e:
                print(f"  ✗ {file}: {str(e)}")
        print(f"✅ Forma8_11 tamamlandı ({success_count}/{len(excel_files)})")
        
        # Forma8_7
        print("\n▶ Forma8_7 işlənir...")
        success_count = 0
        for file in excel_files:
            try:
                total_f = total_f_values.get(file, None)
                run_forma8_7(os.path.join(output_folder, file), previous_folder, reference_date, total_f)
                success_count += 1
                print(f"  ✓ {file}")
            except Exception as e:
                print(f"  ✗ {file}: {str(e)}")
        print(f"✅ Forma8_7 tamamlandı ({success_count}/{len(excel_files)})")
        
        # Forma8_10
        print("\n▶ Forma8_10 işlənir...")
        success_count = 0
        for file in excel_files:
            try:
                run_forma8_10(os.path.join(output_folder, file), previous_folder, reference_date)
                success_count += 1
                print(f"  ✓ {file}")
            except Exception as e:
                print(f"  ✗ {file}: {str(e)}")
        print(f"✅ Forma8_10 tamamlandı ({success_count}/{len(excel_files)})")
        
        # Forma8_8
        print("\n▶ Forma8_8 işlənir...")
        success_count = 0
        for file in excel_files:
            try:
                run_forma8_8(os.path.join(output_folder, file), reference_date, ucot_path)
                success_count += 1
                print(f"  ✓ {file}")
            except Exception as e:
                print(f"  ✗ {file}: {str(e)}")
        print(f"✅ Forma8_8 tamamlandı ({success_count}/{len(excel_files)})")
        
        # Forma8_12
        print("\n▶ Forma8_12 işlənir...")
        success_count = 0
        for file in excel_files:
            try:
                run_forma8_12(os.path.join(output_folder, file), reference_date, ucot_path)
                success_count += 1
                print(f"  ✓ {file}")
            except Exception as e:
                print(f"  ✗ {file}: {str(e)}")
        print(f"✅ Forma8_12 tamamlandı ({success_count}/{len(excel_files)})")
        
        # Forma8_9
        print("\n▶ Forma8_9 işlənir...")
        success_count = 0
        for file in excel_files:
            try:
                run_forma8_9(os.path.join(output_folder, file), reference_date, ucot_path)
                success_count += 1
                print(f"  ✓ {file}")
            except Exception as e:
                print(f"  ✗ {file}: {str(e)}")
        print(f"✅ Forma8_9 tamamlandı ({success_count}/{len(excel_files)})")
        
        # Forma8_13
        print("\n▶ Forma8_13 işlənir...")
        success_count = 0
        for file in excel_files:
            try:
                run_forma8_13(os.path.join(output_folder, file), reference_date, ucot_path)
                success_count += 1
                print(f"  ✓ {file}")
            except Exception as e:
                print(f"  ✗ {file}: {str(e)}")
        print(f"✅ Forma8_13 tamamlandı ({success_count}/{len(excel_files)})")
        
        # Forma8_14
        print("\n▶ Forma8_14 işlənir...")
        success_count = 0
        for file in excel_files:
            try:
                run_forma8_14(os.path.join(output_folder, file), reference_date, ucot_path, previous_folder)
                success_count += 1
                print(f"  ✓ {file}")
            except Exception as e:
                print(f"  ✗ {file}: {str(e)}")
        print(f"✅ Forma8_14 tamamlandı ({success_count}/{len(excel_files)})")
        
        # Yekun Reserv
        if yekun_path:
            print("\n▶ Yekun Reserv işlənir...")
            run_yekun_reserv(yekun_path, output_folder, output_folder, reference_date)
            print("✅ Yekun Reserv tamamlandı")
        
        # Nəticələri ZIP et
        print("\n▶ Nəticələr ZIP edilir...")
        zip_filename = f'Forma8_Results_{session_id}.zip'
        zip_path = os.path.join(session_folder, zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(output_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, output_folder)
                    zipf.write(file_path, arcname)
        
        print(f"✅ ZIP yaradıldı: {zip_filename}")
        print(f"\n{'='*60}")
        print("✅ PROSES TAMAMLANDI")
        print(f"{'='*60}\n")
        
        # Təmizlik (10 dəqiqə sonra)
        def cleanup():
            time.sleep(600)
            shutil.rmtree(session_folder, ignore_errors=True)
            print(f"🗑️  Session təmizləndi: {session_id}")
        
        threading.Thread(target=cleanup, daemon=True).start()
        
        return send_file(zip_path, as_attachment=True, download_name=zip_filename)
        
    except Exception as e:
        print(f"\n❌ XƏTA: {str(e)}\n")
        if session_id:
            shutil.rmtree(os.path.join(app.config['UPLOAD_FOLDER'], session_id), ignore_errors=True)
        return f"Proses zamanı xəta baş verdi: {str(e)}", 500

@app.route('/health')
def health():
    """Server status yoxlaması"""
    return jsonify({
        "status": "ok",
        "message": "Forma8 Server işləyir",
        "timestamp": datetime.now().isoformat()
    })

if __name__ == '__main__':
    print("\n" + "="*60)
    print("🚀 FORMA8 WEB SERVER BAŞLADI")
    print("="*60)
    print("\n⚠️  Server-i dayandırmaq üçün: Ctrl+C")
    print("="*60 + "\n")
    
    # Şəbəkədə bütün cihazların girişi üçün
    app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)