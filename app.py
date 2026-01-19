# -*- coding: utf-8 -*-
"""
邮件批量发送系统 - Web版（单文件版本）
全公司可通过浏览器访问使用
"""

import os
import smtplib
import pandas as pd
from flask import Flask, request, jsonify, session
from werkzeug.utils import secure_filename
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from email.utils import formataddr
import mimetypes

app = Flask(__name__)
app.secret_key = 'email_sender_secret_key_2024'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(os.path.join(app.config['UPLOAD_FOLDER'], 'attachments'), exist_ok=True)

DEFAULT_CONFIG = {
    'smtp_server': 'smtp.exmail.qq.com',
    'smtp_port': 465,
    'sender_email': '',
    'sender_password': '',
    'sender_name': ''
}

HTML_TEMPLATE = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>邮件批量发送系统</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Microsoft YaHei', sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; padding: 20px; }
        .container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 15px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); overflow: hidden; }
        .header { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: white; padding: 25px 30px; text-align: center; }
        .header h1 { font-size: 28px; margin-bottom: 5px; }
        .header p { opacity: 0.8; font-size: 14px; }
        .main-content { display: flex; min-height: 600px; }
        .left-panel { width: 45%; padding: 25px; border-right: 1px solid #eee; }
        .right-panel { width: 55%; padding: 25px; background: #fafafa; }
        .section { margin-bottom: 25px; }
        .section-title { font-size: 16px; font-weight: bold; color: #333; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #667eea; }
        .btn { padding: 12px 24px; border: none; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: bold; transition: all 0.3s; display: inline-flex; align-items: center; gap: 8px; }
        .btn-primary { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; }
        .btn-primary:hover { transform: translateY(-2px); box-shadow: 0 5px 20px rgba(102, 126, 234, 0.4); }
        .btn-success { background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); color: white; }
        .btn-success:hover { transform: translateY(-2px); box-shadow: 0 5px 20px rgba(17, 153, 142, 0.4); }
        .btn-warning { background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); color: white; }
        .btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none !important; }
        .file-input { display: none; }
        .upload-area { display: flex; gap: 15px; flex-wrap: wrap; }
        .status-box { background: #e8f5e9; border: 1px solid #a5d6a7; border-radius: 8px; padding: 15px; margin-top: 15px; }
        .company-table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 13px; }
        .company-table th, .company-table td { padding: 12px; text-align: left; border-bottom: 1px solid #eee; }
        .company-table th { background: #f5f5f5; font-weight: bold; color: #555; }
        .company-table tr:hover { background: #f0f7ff; }
        .company-table tr.active { background: #e3f2fd; }
        .company-table tr.sent { background: #e8f5e9; }
        .company-table tr.sent td:first-child::before { content: '✓ '; color: #4caf50; }
        .table-container { max-height: 300px; overflow-y: auto; border: 1px solid #ddd; border-radius: 8px; }
        .template-input { width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 8px; font-size: 14px; margin-bottom: 10px; }
        .template-textarea { width: 100%; height: 180px; padding: 12px; border: 1px solid #ddd; border-radius: 8px; font-size: 14px; resize: vertical; font-family: inherit; }
        .preview-box { background: white; border: 1px solid #ddd; border-radius: 8px; padding: 20px; margin-top: 15px; }
        .preview-recipient { font-size: 16px; font-weight: bold; color: #1976d2; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 1px solid #eee; }
        .preview-content { white-space: pre-wrap; font-size: 14px; line-height: 1.8; color: #333; }
        .preview-attachments { margin-top: 15px; padding-top: 15px; border-top: 1px solid #eee; color: #666; }
        .nav-buttons { display: flex; gap: 10px; margin-top: 20px; justify-content: center; }
        .btn-nav { padding: 10px 20px; background: #f5f5f5; border: 1px solid #ddd; border-radius: 8px; cursor: pointer; font-size: 14px; }
        .btn-nav:hover:not(:disabled) { background: #e0e0e0; }
        .progress-bar { width: 100%; height: 8px; background: #e0e0e0; border-radius: 4px; overflow: hidden; margin: 15px 0; }
        .progress-fill { height: 100%; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); transition: width 0.3s; }
        .progress-text { text-align: center; font-size: 14px; color: #666; }
        .send-section { text-align: center; padding: 20px; background: white; border-radius: 8px; margin-top: 20px; }
        .btn-send { padding: 18px 50px; font-size: 18px; }
        .hint { font-size: 12px; color: #999; margin-top: 10px; }
        .attachment-list { margin-top: 10px; padding: 10px; background: #f5f5f5; border-radius: 5px; font-size: 13px; }
        .attachment-item { padding: 5px 0; color: #666; }
        .loading { display: none; text-align: center; padding: 20px; }
        .loading.show { display: block; }
        .spinner { border: 3px solid #f3f3f3; border-top: 3px solid #667eea; border-radius: 50%; width: 30px; height: 30px; animation: spin 1s linear infinite; margin: 0 auto 10px; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .toast { position: fixed; top: 20px; right: 20px; padding: 15px 25px; border-radius: 8px; color: white; font-weight: bold; z-index: 1000; animation: slideIn 0.3s ease; }
        .toast.success { background: #4caf50; }
        .toast.error { background: #f44336; }
        @keyframes slideIn { from { transform: translateX(100%); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📧 邮件批量发送系统</h1>
            <p>导入Excel公司列表 → 编辑邮件模板 → 预览确认 → 逐个发送</p>
        </div>
        <div class="main-content">
            <div class="left-panel">
                <div class="section">
                    <div class="section-title">1. 导入数据</div>
                    <div class="upload-area">
                        <input type="file" id="excelFile" class="file-input" accept=".xlsx,.xls">
                        <button class="btn btn-primary" onclick="document.getElementById('excelFile').click()">📂 导入公司列表(Excel)</button>
                        <input type="file" id="attachFiles" class="file-input" multiple>
                        <button class="btn btn-warning" onclick="document.getElementById('attachFiles').click()">📎 添加附件</button>
                    </div>
                    <div id="importStatus" class="status-box" style="display:none;"></div>
                    <div id="attachStatus" class="attachment-list" style="display:none;"></div>
                </div>
                <div class="section">
                    <div class="section-title">2. 公司列表</div>
                    <div class="table-container">
                        <table class="company-table" id="companyTable">
                            <thead><tr><th>序号</th><th>公司名称</th><th>负责人</th><th>邮箱</th></tr></thead>
                            <tbody id="companyList"><tr><td colspan="4" style="text-align:center;color:#999;padding:30px;">请先导入Excel公司列表</td></tr></tbody>
                        </table>
                    </div>
                </div>
                <div class="section">
                    <div class="section-title">3. 发件人设置</div>
                    <input type="text" id="senderName" class="template-input" placeholder="发件人姓名" value="">
                    <input type="email" id="senderEmail" class="template-input" placeholder="发件人邮箱" value="">
                    <input type="password" id="senderPassword" class="template-input" placeholder="邮箱授权码（非登录密码）">
                </div>
            </div>
            <div class="right-panel">
                <div class="section">
                    <div class="section-title">4. 邮件模板</div>
                    <input type="text" id="emailSubject" class="template-input" placeholder="邮件主题" value="合作邀请函">
                    <textarea id="emailTemplate" class="template-textarea" placeholder="邮件正文">尊敬的{company_name}的{contact_person}：

您好！

感谢您百忙之中阅读此邮件。

我们诚挚地希望能与贵公司建立合作关系，共同探讨未来的发展机会。

如有任何问题，欢迎随时与我们联系。

祝好！

{sender_name}</textarea>
                    <div class="hint">可用变量: {company_name}=公司名, {contact_person}=负责人, {sender_name}=你的名字</div>
                </div>
                <div class="section">
                    <div class="section-title">5. 邮件预览</div>
                    <div class="progress-bar"><div class="progress-fill" id="progressFill" style="width:0%"></div></div>
                    <div class="progress-text" id="progressText">进度: 0/0</div>
                    <div class="preview-box">
                        <div class="preview-recipient" id="previewRecipient">请先导入公司列表</div>
                        <div class="preview-content" id="previewContent"></div>
                        <div class="preview-attachments" id="previewAttachments" style="display:none;"></div>
                    </div>
                    <div class="nav-buttons">
                        <button class="btn-nav" id="btnPrev" onclick="prevCompany()" disabled>◀ 上一个</button>
                        <button class="btn-nav" onclick="refreshPreview()">🔄 刷新预览</button>
                        <button class="btn-nav" id="btnNext" onclick="nextCompany()" disabled>下一个 ▶</button>
                    </div>
                </div>
                <div class="send-section">
                    <div style="margin-bottom: 15px;">
                        <label style="font-size: 14px; margin-right: 20px;"><input type="radio" name="sendMode" value="single" checked onchange="updateSendMode()"> 单个发送（逐个确认）</label>
                        <label style="font-size: 14px;"><input type="radio" name="sendMode" value="batch" onchange="updateSendMode()"> 批量发送（一次性全部发送）</label>
                    </div>
                    <div id="singleSendArea">
                        <button class="btn btn-success btn-send" id="btnSend" onclick="sendEmail()" disabled>✉️ 确认发送当前邮件</button>
                        <div class="hint">点击后将发送邮件给当前选中的公司</div>
                    </div>
                    <div id="batchSendArea" style="display: none;">
                        <button class="btn btn-success btn-send" id="btnSendAll" onclick="sendAllEmails()" disabled style="background: linear-gradient(135deg, #ff6b6b 0%, #ee5a24 100%);">🚀 一键发送全部邮件</button>
                        <div class="hint">点击后将一次性发送给所有公司（请先确认预览无误）</div>
                    </div>
                </div>
                <div class="loading" id="loading"><div class="spinner"></div><div>正在发送...</div></div>
            </div>
        </div>
    </div>
    <script>
        let companies = [];
        let currentIndex = 0;
        let sentStatus = [];
        
        document.getElementById('excelFile').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            const formData = new FormData();
            formData.append('file', file);
            fetch('/upload_excel', { method: 'POST', body: formData })
            .then(res => res.json())
            .then(data => {
                if (data.success) {
                    companies = data.companies;
                    sentStatus = new Array(companies.length).fill(false);
                    currentIndex = 0;
                    document.getElementById('importStatus').style.display = 'block';
                    document.getElementById('importStatus').innerHTML = '✓ 成功导入 <strong>' + data.count + '</strong> 个公司';
                    renderCompanyTable();
                    updateUI();
                    showToast('成功导入 ' + data.count + ' 个公司', 'success');
                } else { showToast(data.error, 'error'); }
            }).catch(err => showToast('上传失败: ' + err, 'error'));
        });
        
        document.getElementById('attachFiles').addEventListener('change', function(e) {
            const files = e.target.files;
            if (!files.length) return;
            const formData = new FormData();
            for (let file of files) { formData.append('files', file); }
            fetch('/upload_attachments', { method: 'POST', body: formData })
            .then(res => res.json())
            .then(data => {
                if (data.success) {
                    const attachStatus = document.getElementById('attachStatus');
                    attachStatus.style.display = 'block';
                    attachStatus.innerHTML = '<strong>📎 已添加附件:</strong><br>' + data.files.map(f => '<div class="attachment-item">• ' + f + '</div>').join('');
                    refreshPreview();
                    showToast('成功添加 ' + data.count + ' 个附件', 'success');
                } else { showToast(data.error, 'error'); }
            }).catch(err => showToast('上传失败: ' + err, 'error'));
        });
        
        function renderCompanyTable() {
            const tbody = document.getElementById('companyList');
            if (companies.length === 0) { tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:#999;padding:30px;">请先导入Excel公司列表</td></tr>'; return; }
            tbody.innerHTML = companies.map((c, i) => '<tr class="' + (i === currentIndex ? 'active' : '') + ' ' + (sentStatus[i] ? 'sent' : '') + '" onclick="selectCompany(' + i + ')" style="cursor:pointer;"><td>' + (i + 1) + '</td><td>' + c.name + '</td><td>' + c.contact + '</td><td>' + c.email + '</td></tr>').join('');
        }
        
        function selectCompany(index) { currentIndex = index; updateUI(); }
        function prevCompany() { if (currentIndex > 0) { currentIndex--; updateUI(); } }
        function nextCompany() { if (currentIndex < companies.length - 1) { currentIndex++; updateUI(); } }
        
        function updateUI() {
            renderCompanyTable();
            document.getElementById('btnPrev').disabled = currentIndex <= 0;
            document.getElementById('btnNext').disabled = currentIndex >= companies.length - 1;
            document.getElementById('btnSend').disabled = companies.length === 0;
            document.getElementById('btnSendAll').disabled = companies.length === 0;
            const progress = companies.length > 0 ? ((currentIndex + 1) / companies.length * 100) : 0;
            document.getElementById('progressFill').style.width = progress + '%';
            document.getElementById('progressText').textContent = '进度: ' + (currentIndex + 1) + '/' + companies.length + ' | 已发送: ' + sentStatus.filter(s => s).length;
            if (sentStatus[currentIndex]) {
                document.getElementById('btnSend').textContent = '✓ 已发送 (点击重发)';
                document.getElementById('btnSend').style.background = '#9e9e9e';
            } else {
                document.getElementById('btnSend').textContent = '✉️ 确认发送当前邮件';
                document.getElementById('btnSend').style.background = '';
            }
            const unsent = sentStatus.filter(s => !s).length;
            if (unsent === 0 && companies.length > 0) {
                document.getElementById('btnSendAll').textContent = '✓ 全部已发送';
                document.getElementById('btnSendAll').style.background = '#9e9e9e';
            } else {
                document.getElementById('btnSendAll').textContent = '🚀 一键发送全部邮件 (' + unsent + '封待发)';
            }
            refreshPreview();
        }
        
        function refreshPreview() {
            if (companies.length === 0) { document.getElementById('previewRecipient').textContent = '请先导入公司列表'; document.getElementById('previewContent').textContent = ''; return; }
            fetch('/preview_email', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ index: currentIndex, template: document.getElementById('emailTemplate').value, subject: document.getElementById('emailSubject').value, sender_name: document.getElementById('senderName').value }) })
            .then(res => res.json())
            .then(data => {
                if (data.success) {
                    document.getElementById('previewRecipient').textContent = '收件人: ' + data.recipient;
                    document.getElementById('previewContent').textContent = '【主题】' + data.subject + '\\n\\n【正文】\\n' + data.body;
                    if (data.attachments && data.attachments.length > 0) { document.getElementById('previewAttachments').style.display = 'block'; document.getElementById('previewAttachments').innerHTML = '<strong>📎 附件:</strong> ' + data.attachments.join(', '); }
                    else { document.getElementById('previewAttachments').style.display = 'none'; }
                }
            });
        }
        
        function sendEmail() {
            if (companies.length === 0) return;
            const company = companies[currentIndex];
            if (!confirm('确定要发送邮件给:\\n\\n公司: ' + company.name + '\\n负责人: ' + company.contact + '\\n邮箱: ' + company.email + '\\n\\n请确认预览内容无误！')) return;
            document.getElementById('loading').classList.add('show');
            document.getElementById('btnSend').disabled = true;
            fetch('/send_email', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ index: currentIndex, template: document.getElementById('emailTemplate').value, subject: document.getElementById('emailSubject').value, sender_name: document.getElementById('senderName').value, sender_email: document.getElementById('senderEmail').value, sender_password: document.getElementById('senderPassword').value }) })
            .then(res => res.json())
            .then(data => {
                document.getElementById('loading').classList.remove('show');
                if (data.success) { sentStatus[currentIndex] = true; showToast(data.message, 'success'); if (currentIndex < companies.length - 1) { currentIndex++; } updateUI(); }
                else { showToast(data.error, 'error'); document.getElementById('btnSend').disabled = false; }
            }).catch(err => { document.getElementById('loading').classList.remove('show'); showToast('发送失败: ' + err, 'error'); document.getElementById('btnSend').disabled = false; });
        }
        
        function showToast(message, type) { const toast = document.createElement('div'); toast.className = 'toast ' + type; toast.textContent = message; document.body.appendChild(toast); setTimeout(() => { toast.remove(); }, 3000); }
        function updateSendMode() { const mode = document.querySelector('input[name="sendMode"]:checked').value; document.getElementById('singleSendArea').style.display = mode === 'single' ? 'block' : 'none'; document.getElementById('batchSendArea').style.display = mode === 'batch' ? 'block' : 'none'; }
        
        function sendAllEmails() {
            if (companies.length === 0) return;
            const unsent = sentStatus.filter(s => !s).length;
            if (!confirm('确定要一次性发送邮件给所有公司吗？\\n\\n总计: ' + companies.length + ' 个公司\\n待发送: ' + unsent + ' 封\\n\\n请确认邮件模板和附件无误！')) return;
            document.getElementById('loading').classList.add('show');
            document.getElementById('btnSendAll').disabled = true;
            fetch('/send_all', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ template: document.getElementById('emailTemplate').value, subject: document.getElementById('emailSubject').value, sender_name: document.getElementById('senderName').value, sender_email: document.getElementById('senderEmail').value, sender_password: document.getElementById('senderPassword').value }) })
            .then(res => res.json())
            .then(data => {
                document.getElementById('loading').classList.remove('show');
                if (data.success) { data.results.forEach(r => { if (r.success) { sentStatus[r.index] = true; } }); updateUI(); showToast('发送完成！成功: ' + data.success_count + ', 失败: ' + data.fail_count, data.fail_count > 0 ? 'error' : 'success'); if (data.fail_count > 0) { const failedItems = data.results.filter(r => !r.success); alert('以下邮件发送失败:\\n\\n' + failedItems.map(r => companies[r.index].name + ': ' + r.message).join('\\n')); } }
                else { showToast(data.error, 'error'); }
                document.getElementById('btnSendAll').disabled = false;
            }).catch(err => { document.getElementById('loading').classList.remove('show'); showToast('发送失败: ' + err, 'error'); document.getElementById('btnSendAll').disabled = false; });
        }
        
        document.getElementById('emailTemplate').addEventListener('input', refreshPreview);
        document.getElementById('emailSubject').addEventListener('input', refreshPreview);
        document.getElementById('senderName').addEventListener('input', refreshPreview);
    </script>
</body>
</html>'''


@app.route('/')
def index():
    return HTML_TEMPLATE


@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '没有选择文件'})
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '没有选择文件'})
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'success': False, 'error': '请上传Excel文件(.xlsx或.xls)'})
    try:
        filename = secure_filename(file.filename) or 'upload.xlsx'
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        df = pd.read_excel(filepath, engine='openpyxl')
        if '公司名称' not in df.columns or '邮箱地址' not in df.columns:
            return jsonify({'success': False, 'error': 'Excel必须包含"公司名称"和"邮箱地址"列'})
        has_contact = '负责人' in df.columns
        companies = []
        for _, row in df.iterrows():
            name = str(row['公司名称']).strip()
            email = str(row['邮箱地址']).strip()
            contact = str(row['负责人']).strip() if has_contact else '负责人'
            if contact == 'nan' or not contact:
                contact = '负责人'
            if name and email and name != 'nan' and email != 'nan':
                companies.append({'name': name, 'email': email, 'contact': contact})
        session['companies'] = companies
        session['sent_status'] = [False] * len(companies)
        return jsonify({'success': True, 'count': len(companies), 'companies': companies})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/upload_attachments', methods=['POST'])
def upload_attachments():
    if 'files' not in request.files:
        return jsonify({'success': False, 'error': '没有选择文件'})
    files = request.files.getlist('files')
    if not files or files[0].filename == '':
        return jsonify({'success': False, 'error': '没有选择文件'})
    try:
        attach_folder = os.path.join(app.config['UPLOAD_FOLDER'], 'attachments')
        for f in os.listdir(attach_folder):
            os.remove(os.path.join(attach_folder, f))
        saved_files = []
        for file in files:
            filename = secure_filename(file.filename) or file.filename
            filepath = os.path.join(attach_folder, filename)
            file.save(filepath)
            saved_files.append({'name': filename, 'path': filepath})
        session['attachments'] = saved_files
        return jsonify({'success': True, 'count': len(saved_files), 'files': [f['name'] for f in saved_files]})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/preview_email', methods=['POST'])
def preview_email():
    data = request.json
    index = data.get('index', 0)
    template = data.get('template', '')
    subject = data.get('subject', '合作邀请函')
    sender_name = data.get('sender_name', '')
    companies = session.get('companies', [])
    if not companies or index >= len(companies):
        return jsonify({'success': False, 'error': '没有公司数据'})
    company = companies[index]
    try:
        body = template.format(company_name=company['name'], contact_person=company['contact'], sender_name=sender_name)
    except KeyError as e:
        body = f'[模板错误: 缺少变量 {e}]'
    attachments = session.get('attachments', [])
    return jsonify({'success': True, 'recipient': f"{company['name']} - {company['contact']} <{company['email']}>", 'subject': subject, 'body': body, 'attachments': [a['name'] for a in attachments], 'sent': session.get('sent_status', [])[index] if session.get('sent_status') else False})


@app.route('/send_email', methods=['POST'])
def send_email():
    data = request.json
    index = data.get('index', 0)
    template = data.get('template', '')
    subject = data.get('subject', '合作邀请函')
    sender_email = data.get('sender_email', '')
    sender_password = data.get('sender_password', '')
    sender_name = data.get('sender_name', '')
    
    if not sender_email or not sender_password:
        return jsonify({'success': False, 'error': '请填写发件人邮箱和授权码'})
    
    companies = session.get('companies', [])
    if not companies or index >= len(companies):
        return jsonify({'success': False, 'error': '没有公司数据'})
    company = companies[index]
    try:
        msg = MIMEMultipart()
        msg['From'] = formataddr((sender_name, sender_email))
        msg['To'] = formataddr((company['name'], company['email']))
        msg['Subject'] = Header(subject, 'utf-8')
        body = template.format(company_name=company['name'], contact_person=company['contact'], sender_name=sender_name)
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        attachments = session.get('attachments', [])
        for attach in attachments:
            filepath = attach['path']
            if os.path.exists(filepath):
                filename = attach['name']
                mime_type, _ = mimetypes.guess_type(filepath)
                if mime_type is None:
                    mime_type = 'application/octet-stream'
                main_type, sub_type = mime_type.split('/', 1)
                with open(filepath, 'rb') as f:
                    part = MIMEBase(main_type, sub_type)
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                encoded_filename = Header(filename, 'utf-8').encode()
                part.add_header('Content-Disposition', 'attachment', filename=encoded_filename)
                msg.attach(part)
        server = smtplib.SMTP_SSL(DEFAULT_CONFIG['smtp_server'], DEFAULT_CONFIG['smtp_port'])
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        sent_status = session.get('sent_status', [False] * len(companies))
        sent_status[index] = True
        session['sent_status'] = sent_status
        return jsonify({'success': True, 'message': f'邮件已成功发送给 {company["name"]}'})
    except smtplib.SMTPAuthenticationError:
        return jsonify({'success': False, 'error': '邮箱认证失败，请检查邮箱和授权码'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/send_all', methods=['POST'])
def send_all():
    data = request.json
    template = data.get('template', '')
    subject = data.get('subject', '合作邀请函')
    sender_email = data.get('sender_email', '')
    sender_password = data.get('sender_password', '')
    sender_name = data.get('sender_name', '')
    
    if not sender_email or not sender_password:
        return jsonify({'success': False, 'error': '请填写发件人邮箱和授权码'})
    
    companies = session.get('companies', [])
    if not companies:
        return jsonify({'success': False, 'error': '没有公司数据'})
    sent_status = session.get('sent_status', [False] * len(companies))
    results = []
    success_count = 0
    fail_count = 0
    for index, company in enumerate(companies):
        if sent_status[index]:
            results.append({'index': index, 'success': True, 'message': '已跳过（之前已发送）'})
            continue
        try:
            msg = MIMEMultipart()
            msg['From'] = formataddr((sender_name, sender_email))
            msg['To'] = formataddr((company['name'], company['email']))
            msg['Subject'] = Header(subject, 'utf-8')
            body = template.format(company_name=company['name'], contact_person=company['contact'], sender_name=sender_name)
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            attachments = session.get('attachments', [])
            for attach in attachments:
                filepath = attach['path']
                if os.path.exists(filepath):
                    filename = attach['name']
                    mime_type, _ = mimetypes.guess_type(filepath)
                    if mime_type is None:
                        mime_type = 'application/octet-stream'
                    main_type, sub_type = mime_type.split('/', 1)
                    with open(filepath, 'rb') as f:
                        part = MIMEBase(main_type, sub_type)
                        part.set_payload(f.read())
                    encoders.encode_base64(part)
                    encoded_filename = Header(filename, 'utf-8').encode()
                    part.add_header('Content-Disposition', 'attachment', filename=encoded_filename)
                    msg.attach(part)
            server = smtplib.SMTP_SSL(DEFAULT_CONFIG['smtp_server'], DEFAULT_CONFIG['smtp_port'])
            server.login(sender_email, sender_password)
            server.send_message(msg)
            server.quit()
            sent_status[index] = True
            success_count += 1
            results.append({'index': index, 'success': True, 'message': '发送成功'})
        except Exception as e:
            fail_count += 1
            results.append({'index': index, 'success': False, 'message': str(e)})
    session['sent_status'] = sent_status
    return jsonify({'success': True, 'total': len(companies), 'success_count': success_count, 'fail_count': fail_count, 'results': results})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
