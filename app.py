# -*- coding: utf-8 -*-
"""
邮件批量发送系统 - Web版
全公司可通过浏览器访问使用
"""

import os
import smtplib
import pandas as pd
from flask import Flask, render_template, request, jsonify, session
from werkzeug.utils import secure_filename
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header, make_header, decode_header
from email.utils import formataddr
import mimetypes

app = Flask(__name__, template_folder='.')
app.secret_key = 'email_sender_secret_key_2024'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

# 确保上传目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(os.path.join(app.config['UPLOAD_FOLDER'], 'attachments'), exist_ok=True)

# 默认邮件配置
DEFAULT_CONFIG = {
    'smtp_server': 'smtp.exmail.qq.com',
    'smtp_port': 465,
    'sender_email': 'yuqing.zhang@insightst.com',
    'sender_password': 'XhCVbJ9zfa6RqBXb',
    'sender_name': '张玉青'
}


@app.route('/')
def index():
    """主页"""
    return render_template('index.html')


@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    """上传Excel文件"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '没有选择文件'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '没有选择文件'})
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'success': False, 'error': '请上传Excel文件(.xlsx或.xls)'})
    
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # 读取Excel
        df = pd.read_excel(filepath, engine='openpyxl')
        
        if '公司名称' not in df.columns or '邮箱地址' not in df.columns:
            return jsonify({
                'success': False, 
                'error': 'Excel必须包含"公司名称"和"邮箱地址"列'
            })
        
        has_contact = '负责人' in df.columns
        
        companies = []
        for _, row in df.iterrows():
            name = str(row['公司名称']).strip()
            email = str(row['邮箱地址']).strip()
            contact = str(row['负责人']).strip() if has_contact else '负责人'
            if contact == 'nan' or not contact:
                contact = '负责人'
            if name and email and name != 'nan' and email != 'nan':
                companies.append({
                    'name': name,
                    'email': email,
                    'contact': contact
                })
        
        # 存储到session
        session['companies'] = companies
        session['sent_status'] = [False] * len(companies)
        
        return jsonify({
            'success': True,
            'count': len(companies),
            'companies': companies
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/upload_attachments', methods=['POST'])
def upload_attachments():
    """上传附件"""
    if 'files' not in request.files:
        return jsonify({'success': False, 'error': '没有选择文件'})
    
    files = request.files.getlist('files')
    if not files or files[0].filename == '':
        return jsonify({'success': False, 'error': '没有选择文件'})
    
    try:
        # 清空之前的附件
        attach_folder = os.path.join(app.config['UPLOAD_FOLDER'], 'attachments')
        for f in os.listdir(attach_folder):
            os.remove(os.path.join(attach_folder, f))
        
        saved_files = []
        for file in files:
            filename = secure_filename(file.filename)
            # 处理中文文件名
            if not filename:
                filename = file.filename
            filepath = os.path.join(attach_folder, filename)
            file.save(filepath)
            saved_files.append({'name': filename, 'path': filepath})
        
        session['attachments'] = saved_files
        
        return jsonify({
            'success': True,
            'count': len(saved_files),
            'files': [f['name'] for f in saved_files]
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/preview_email', methods=['POST'])
def preview_email():
    """预览邮件"""
    data = request.json
    index = data.get('index', 0)
    template = data.get('template', '')
    subject = data.get('subject', '合作邀请函')
    sender_name = data.get('sender_name', DEFAULT_CONFIG['sender_name'])
    
    companies = session.get('companies', [])
    if not companies or index >= len(companies):
        return jsonify({'success': False, 'error': '没有公司数据'})
    
    company = companies[index]
    
    try:
        body = template.format(
            company_name=company['name'],
            contact_person=company['contact'],
            sender_name=sender_name
        )
    except KeyError as e:
        body = f'[模板错误: 缺少变量 {e}]'
    
    attachments = session.get('attachments', [])
    
    return jsonify({
        'success': True,
        'recipient': f"{company['name']} - {company['contact']} <{company['email']}>",
        'subject': subject,
        'body': body,
        'attachments': [a['name'] for a in attachments],
        'sent': session.get('sent_status', [])[index] if session.get('sent_status') else False
    })


@app.route('/send_email', methods=['POST'])
def send_email():
    """发送邮件"""
    data = request.json
    index = data.get('index', 0)
    template = data.get('template', '')
    subject = data.get('subject', '合作邀请函')
    sender_email = data.get('sender_email', DEFAULT_CONFIG['sender_email'])
    sender_password = data.get('sender_password', DEFAULT_CONFIG['sender_password'])
    sender_name = data.get('sender_name', DEFAULT_CONFIG['sender_name'])
    
    companies = session.get('companies', [])
    if not companies or index >= len(companies):
        return jsonify({'success': False, 'error': '没有公司数据'})
    
    company = companies[index]
    
    try:
        # 组装邮件
        msg = MIMEMultipart()
        msg['From'] = formataddr((sender_name, sender_email))
        msg['To'] = formataddr((company['name'], company['email']))
        msg['Subject'] = Header(subject, 'utf-8')
        
        body = template.format(
            company_name=company['name'],
            contact_person=company['contact'],
            sender_name=sender_name
        )
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # 添加附件
        attachments = session.get('attachments', [])
        for attach in attachments:
            filepath = attach['path']
            if os.path.exists(filepath):
                filename = attach['name']
                # 获取文件MIME类型
                mime_type, _ = mimetypes.guess_type(filepath)
                if mime_type is None:
                    mime_type = 'application/octet-stream'
                main_type, sub_type = mime_type.split('/', 1)
                
                with open(filepath, 'rb') as f:
                    part = MIMEBase(main_type, sub_type)
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                # 使用RFC 2231编码处理中文文件名
                encoded_filename = Header(filename, 'utf-8').encode()
                part.add_header('Content-Disposition', 'attachment',
                                filename=encoded_filename)
                msg.attach(part)
        
        # 发送
        server = smtplib.SMTP_SSL(DEFAULT_CONFIG['smtp_server'], DEFAULT_CONFIG['smtp_port'])
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        
        # 更新发送状态
        sent_status = session.get('sent_status', [False] * len(companies))
        sent_status[index] = True
        session['sent_status'] = sent_status
        
        return jsonify({
            'success': True,
            'message': f'邮件已成功发送给 {company["name"]}'
        })
        
    except smtplib.SMTPAuthenticationError:
        return jsonify({'success': False, 'error': '邮箱认证失败，请检查邮箱和授权码'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/get_status')
def get_status():
    """获取发送状态"""
    companies = session.get('companies', [])
    sent_status = session.get('sent_status', [])
    
    sent_count = sum(1 for s in sent_status if s)
    
    return jsonify({
        'total': len(companies),
        'sent': sent_count,
        'status': sent_status
    })


@app.route('/send_all', methods=['POST'])
def send_all():
    """批量发送所有邮件"""
    data = request.json
    template = data.get('template', '')
    subject = data.get('subject', '合作邀请函')
    sender_email = data.get('sender_email', DEFAULT_CONFIG['sender_email'])
    sender_password = data.get('sender_password', DEFAULT_CONFIG['sender_password'])
    sender_name = data.get('sender_name', DEFAULT_CONFIG['sender_name'])
    
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
            # 组装邮件
            msg = MIMEMultipart()
            msg['From'] = formataddr((sender_name, sender_email))
            msg['To'] = formataddr((company['name'], company['email']))
            msg['Subject'] = Header(subject, 'utf-8')
            
            body = template.format(
                company_name=company['name'],
                contact_person=company['contact'],
                sender_name=sender_name
            )
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # 添加附件
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
                    part.add_header('Content-Disposition', 'attachment',
                                    filename=encoded_filename)
                    msg.attach(part)
            
            # 发送
            server = smtplib.SMTP_SSL(DEFAULT_CONFIG['smtp_server'], DEFAULT_CONFIG['smtp_port'])
            server.login(sender_email, sender_password)
            server.send_message(msg)
            server.quit()
            
            sent_status[index] = True
            success_count += 1
            results.append({'index': index, 'success': True, 'message': f'发送成功'})
            
        except Exception as e:
            fail_count += 1
            results.append({'index': index, 'success': False, 'message': str(e)})
    
    session['sent_status'] = sent_status
    
    return jsonify({
        'success': True,
        'total': len(companies),
        'success_count': success_count,
        'fail_count': fail_count,
        'results': results
    })


if __name__ == '__main__':
    # 局域网内可访问，使用 0.0.0.0
    app.run(host='0.0.0.0', port=5000, debug=True)

