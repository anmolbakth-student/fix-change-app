
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, send_file
import sqlite3
from datetime import datetime, timedelta
import csv
import io
import os
import json
import traceback
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

app = Flask(__name__)
app.secret_key = 'web-app-secret-key-2026'

DATABASE = 'attendance.db'
ACCESS_CODE = 'alusman@123'
ADMIN_CODE = 'admin@2026'
LICENSE_FILE = '.license'
EXPIRY_DAYS = 60
ACTIVATION_KEY = 'azeem@taizco.com'

def is_license_valid():
    if not os.path.exists(LICENSE_FILE):
        return False
    try:
        with open(LICENSE_FILE, 'r') as f:
            data = json.load(f)
        expiry = datetime.strptime(data['expiry'], '%Y-%m-%d %H:%M:%S')
        return datetime.now() <= expiry
    except:
        return False

def license_required(f):
    def decorated_function(*args, **kwargs):
        if not is_license_valid():
            return redirect(url_for('license_expired'))
        return f(*args, **kwargs)
    decorated_function.__name__ = f.__name__
    return decorated_function

def get_db():
    conn = sqlite3.connect(DATABASE, timeout=20)
    conn.row_factory = sqlite3.Row
    return conn

def get_all_active_users(conn, floor_filter=None):
    cursor = conn.cursor()
    query = "SELECT user_id, name, device_floor FROM users WHERE is_active = 1"
    params = []
    if floor_filter:
        query += " AND device_floor = ?"
        params.append(floor_filter)
    query += " ORDER BY name"
    cursor.execute(query, params)
    return cursor.fetchall()

def calculate_remark_in(time_str, day_name):
    if not time_str or time_str == '--':
        return 'Missing In'
    if day_name == 'Sunday':
        return 'Over Time'
    try:
        time_obj = datetime.strptime(time_str, '%H:%M')
        if day_name == 'Saturday':
            start = datetime.strptime('10:00', '%H:%M').time()
            grace = datetime.strptime('10:30', '%H:%M').time()
        else:
            start = datetime.strptime('09:15', '%H:%M').time()
            grace = datetime.strptime('09:30', '%H:%M').time()
        check_time = time_obj.time()
        if check_time <= start:
            return 'On Time'
        elif check_time <= grace:
            return 'Grace'
        else:
            return 'Late'
    except:
        return 'On Time'

def calculate_remark_out(time_str, day_name):
    if not time_str or time_str == '--':
        return 'Missing Out'
    if day_name == 'Sunday':
        return 'Over Time'
    try:
        time_obj = datetime.strptime(time_str, '%H:%M')
        if day_name == 'Saturday':
            early = datetime.strptime('14:00', '%H:%M').time()
            overtime = datetime.strptime('16:00', '%H:%M').time()
        else:
            early = datetime.strptime('17:00', '%H:%M').time()
            overtime = datetime.strptime('18:00', '%H:%M').time()
        check_time = time_obj.time()
        if check_time < early:
            return 'Early Dept'
        elif check_time > overtime:
            return 'Overtime'
        else:
            return 'On Time'
    except:
        return 'On Time'

def calculate_status(in_time, out_time, day_name, has_in, has_out, status_override=None):
    if status_override:
        return status_override
    if day_name == 'Sunday':
        if has_in and has_out:
            return 'Sunday / Over Time'
        else:
            return 'Weekly Off'
    in_present = has_in
    out_present = has_out
    if in_present and out_present:
        in_remark = calculate_remark_in(in_time, day_name)
        out_remark = calculate_remark_out(out_time, day_name)
        return f"Present / {in_remark} / {out_remark}"
    elif in_present and not out_present:
        return f"Present / Missing Out"
    elif not in_present and out_present:
        return f"Present / Missing In"
    else:
        return "Leave / Absent(Document your Truancy) / Missing In-Out"

def get_display_time(time_str):
    if not time_str or time_str == '--':
        return '--'
    return time_str

def get_export_data_grouped(date_from, date_to, selected_floor, search_user):
    conn = get_db()
    cursor = conn.cursor()
    users = get_all_active_users(conn, selected_floor)
    
    query = '''
        SELECT 
            a.user_id,
            a.device_floor,
            DATE(a.timestamp) as date_only,
            strftime('%w', a.timestamp) as weekday,
            strftime('%H:%M', a.timestamp) as time,
            a.punch_type,
            u.name as user_name,
            a.remark_in,
            a.remark_out,
            a.status_override,
            a.is_manually_edited
        FROM attendance_logs a
        LEFT JOIN users u ON a.user_id = u.user_id AND a.device_floor = u.device_floor
        WHERE 1=1
    '''
    params = []
    if date_from:
        query += ' AND DATE(a.timestamp) >= ?'
        params.append(date_from)
    if date_to:
        query += ' AND DATE(a.timestamp) <= ?'
        params.append(date_to)
    if selected_floor:
        query += ' AND a.device_floor = ?'
        params.append(selected_floor)
    if search_user:
        query += ' AND (a.user_id LIKE ? OR u.name LIKE ?)'
        search_pattern = f'%{search_user}%'
        params.extend([search_pattern, search_pattern])
    query += ' ORDER BY a.timestamp ASC'
    cursor.execute(query, params)
    raw_logs = cursor.fetchall()
    
    weekday_map = {
        '0': 'Sunday', '1': 'Monday', '2': 'Tuesday', '3': 'Wednesday',
        '4': 'Thursday', '5': 'Friday', '6': 'Saturday'
    }
    
    log_map = {}
    for log in raw_logs:
        key = f"{log['user_id']}|{log['date_only']}|{log['device_floor']}"
        if key not in log_map:
            log_map[key] = {
                'user_id': log['user_id'],
                'user_name': log['user_name'] or 'Unknown',
                'device_floor': log['device_floor'],
                'date_only': log['date_only'],
                'day_name': weekday_map.get(log['weekday'], 'Unknown'),
                'in_times': [],
                'out_times': [],
                'remark_in': None,
                'remark_out': None,
                'status_override': None,
                'is_manually_edited': 0
            }
        if log['punch_type'] == 0:
            log_map[key]['in_times'].append(log['time'])
            log_map[key]['remark_in'] = log['remark_in']
        else:
            log_map[key]['out_times'].append(log['time'])
            log_map[key]['remark_out'] = log['remark_out']
        if log['is_manually_edited'] == 1:
            log_map[key]['is_manually_edited'] = 1
        if log['status_override']:
            log_map[key]['status_override'] = log['status_override']
    
    table_data = []
    for user in users:
        user_id = user['user_id']
        user_name = user['name']
        device_floor = user['device_floor']
        
        if date_from and date_to:
            start = datetime.strptime(date_from, '%Y-%m-%d')
            end = datetime.strptime(date_to, '%Y-%m-%d')
            date_list = []
            current = start
            while current <= end:
                date_list.append(current.strftime('%Y-%m-%d'))
                current += timedelta(days=1)
        else:
            date_list = [datetime.now().strftime('%Y-%m-%d')]
        
        for date_str in date_list:
            key = f"{user_id}|{date_str}|{device_floor}"
            if key in log_map:
                data = log_map[key]
                day_name = data['day_name']
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                date_formatted = date_obj.strftime('%d-%b-%Y')
                in_times = data['in_times']
                out_times = data['out_times']
                in_time = min(in_times) if in_times else None
                out_time = max(out_times) if out_times else None
                has_in = in_time is not None
                has_out = out_time is not None
                chk_in = get_display_time(in_time) if has_in else '--'
                chk_out = get_display_time(out_time) if has_out else '--'
                
                if data.get('is_manually_edited', 0) == 1:
                    remark_in = data.get('remark_in') or 'Missing In'
                    remark_out = data.get('remark_out') or 'Missing Out'
                    status = data.get('status_override') or calculate_status(in_time, out_time, day_name, has_in, has_out)
                else:
                    remark_in = calculate_remark_in(in_time, day_name) if has_in else ('Over Time' if day_name == 'Sunday' else 'Missing In')
                    remark_out = calculate_remark_out(out_time, day_name) if has_out else ('Over Time' if day_name == 'Sunday' else 'Missing Out')
                    status = calculate_status(in_time, out_time, day_name, has_in, has_out, data.get('status_override'))
                
                table_data.append({
                    'user_id': user_id,
                    'user_name': user_name,
                    'device_floor': device_floor,
                    'date': date_formatted,
                    'day': day_name,
                    'chk_in': chk_in,
                    'remark_in': remark_in,
                    'chk_out': chk_out,
                    'remark_out': remark_out,
                    'status': status
                })
    conn.close()
    return table_data

@app.route('/license-expired')
def license_expired():
    return render_template('license_expired.html'), 403

@app.route('/activate', methods=['GET', 'POST'])
def activate():
    if request.method == 'POST':
        key = request.form.get('key', '').strip()
        if key == ACTIVATION_KEY:
            data = {
                'expiry': (datetime.now() + timedelta(days=EXPIRY_DAYS)).strftime('%Y-%m-%d %H:%M:%S'),
                'created': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            with open(LICENSE_FILE, 'w') as f:
                json.dump(data, f)
            flash('License activated!', 'success')
            return redirect(url_for('login'))
        flash('Invalid key', 'danger')
    return render_template('activate.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'access_code' in session:
        return redirect(url_for('index'))
    if request.method == 'POST':
        code = request.form.get('access_code', '').strip()
        if code == ADMIN_CODE:
            session['access_code'] = code
            session['role'] = 'admin'
            return redirect(url_for('index'))
        elif code == ACCESS_CODE:
            session['access_code'] = code
            session['role'] = 'viewer'
            return redirect(url_for('index'))
        flash('Invalid code', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash('Logged out', 'info')
    return redirect(url_for('login'))

@app.route('/')
@license_required
def index():
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    page = request.args.get('page', 1, type=int)
    per_page = 30
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')
    
    if not date_from and not date_to:
        date_to = datetime.now().strftime('%Y-%m-%d')
        date_from = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
    
    conn = get_db()
    cursor = conn.cursor()
    
    users = get_all_active_users(conn, selected_floor)
    if search_user:
        users = [u for u in users if search_user.lower() in u['name'].lower() or search_user in str(u['user_id'])]
    
    query = '''
        SELECT 
            a.id as log_id,
            a.user_id,
            a.device_floor,
            DATE(a.timestamp) as date_only,
            strftime('%w', a.timestamp) as weekday,
            strftime('%H:%M', a.timestamp) as time,
            a.punch_type,
            u.name as user_name,
            a.remark_in,
            a.remark_out,
            a.status_override,
            a.is_manually_edited
        FROM attendance_logs a
        LEFT JOIN users u ON a.user_id = u.user_id AND a.device_floor = u.device_floor
        WHERE 1=1
    '''
    params = []
    if date_from:
        query += ' AND DATE(a.timestamp) >= ?'
        params.append(date_from)
    if date_to:
        query += ' AND DATE(a.timestamp) <= ?'
        params.append(date_to)
    if selected_floor:
        query += ' AND a.device_floor = ?'
        params.append(selected_floor)
    if search_user:
        query += ' AND (a.user_id LIKE ? OR u.name LIKE ?)'
        search_pattern = f'%{search_user}%'
        params.extend([search_pattern, search_pattern])
    query += ' ORDER BY a.timestamp ASC'
    cursor.execute(query, params)
    raw_logs = cursor.fetchall()
    
    weekday_map = {
        '0': 'Sunday', '1': 'Monday', '2': 'Tuesday', '3': 'Wednesday',
        '4': 'Thursday', '5': 'Friday', '6': 'Saturday'
    }
    
    log_map = {}
    for log in raw_logs:
        key = f"{log['user_id']}|{log['date_only']}|{log['device_floor']}"
        if key not in log_map:
            log_map[key] = {
                'user_id': log['user_id'],
                'user_name': log['user_name'] or 'Unknown',
                'device_floor': log['device_floor'],
                'date_only': log['date_only'],
                'day_name': weekday_map.get(log['weekday'], 'Unknown'),
                'in_times': [],
                'out_times': [],
                'in_id': None,
                'out_id': None,
                'remark_in': None,
                'remark_out': None,
                'status_override': None,
                'is_manually_edited': 0
            }
        if log['punch_type'] == 0:
            log_map[key]['in_times'].append(log['time'])
            if log_map[key]['in_id'] is None:
                log_map[key]['in_id'] = log['log_id']
            log_map[key]['remark_in'] = log['remark_in']
        else:
            log_map[key]['out_times'].append(log['time'])
            if log_map[key]['out_id'] is None:
                log_map[key]['out_id'] = log['log_id']
            log_map[key]['remark_out'] = log['remark_out']
        if log['is_manually_edited'] == 1:
            log_map[key]['is_manually_edited'] = 1
        if log['status_override']:
            log_map[key]['status_override'] = log['status_override']
    
    all_table_data = []
    start = datetime.strptime(date_from, '%Y-%m-%d')
    end = datetime.strptime(date_to, '%Y-%m-%d')
    date_list = []
    current = start
    while current <= end:
        date_list.append(current.strftime('%Y-%m-%d'))
        current += timedelta(days=1)
    
    for user in users:
        user_id = user['user_id']
        user_name = user['name']
        device_floor = user['device_floor']
        for date_str in date_list:
            key = f"{user_id}|{date_str}|{device_floor}"
            if key in log_map:
                data = log_map[key]
                day_name = data['day_name']
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                date_formatted = date_obj.strftime('%d-%b-%Y')
                in_times = data['in_times']
                out_times = data['out_times']
                in_time = min(in_times) if in_times else None
                out_time = max(out_times) if out_times else None
                has_in = in_time is not None
                has_out = out_time is not None
                chk_in = get_display_time(in_time) if has_in else '--'
                chk_out = get_display_time(out_time) if has_out else '--'
                if data.get('is_manually_edited', 0) == 1:
                    remark_in = data.get('remark_in') or 'Missing In'
                    remark_out = data.get('remark_out') or 'Missing Out'
                    status = data.get('status_override') or calculate_status(in_time, out_time, day_name, has_in, has_out)
                else:
                    remark_in = calculate_remark_in(in_time, day_name) if has_in else ('Over Time' if day_name == 'Sunday' else 'Missing In')
                    remark_out = calculate_remark_out(out_time, day_name) if has_out else ('Over Time' if day_name == 'Sunday' else 'Missing Out')
                    status = calculate_status(in_time, out_time, day_name, has_in, has_out, data.get('status_override'))
                all_table_data.append({
                    'user_id': user_id,
                    'user_name': user_name,
                    'device_floor': device_floor,
                    'date': date_formatted,
                    'day': day_name,
                    'chk_in': chk_in,
                    'remark_in': remark_in,
                    'chk_out': chk_out,
                    'remark_out': remark_out,
                    'status': status,
                    'is_sunday': day_name == 'Sunday',
                    'in_id': data.get('in_id'),
                    'out_id': data.get('out_id')
                })
            else:
                day_name = datetime.strptime(date_str, '%Y-%m-%d').strftime('%A')
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                date_formatted = date_obj.strftime('%d-%b-%Y')
                if day_name == 'Sunday':
                    all_table_data.append({
                        'user_id': user_id,
                        'user_name': user_name,
                        'device_floor': device_floor,
                        'date': date_formatted,
                        'day': day_name,
                        'chk_in': '--',
                        'remark_in': '',
                        'chk_out': '--',
                        'remark_out': '',
                        'status': 'Weekly Off',
                        'is_sunday': True,
                        'in_id': None,
                        'out_id': None
                    })
                else:
                    all_table_data.append({
                        'user_id': user_id,
                        'user_name': user_name,
                        'device_floor': device_floor,
                        'date': date_formatted,
                        'day': day_name,
                        'chk_in': '--',
                        'remark_in': 'Missing In',
                        'chk_out': '--',
                        'remark_out': 'Missing Out',
                        'status': 'Leave / Absent(Document your Truancy) / Missing In-Out',
                        'is_sunday': False,
                        'in_id': None,
                        'out_id': None
                    })
    
    all_table_data.sort(key=lambda x: datetime.strptime(x['date'], '%d-%b-%Y'), reverse=True)
    total_count = len(all_table_data)
    total_pages = (total_count + per_page - 1) // per_page if total_count > 0 else 1
    start_idx = (page - 1) * per_page
    end_idx = start_idx + per_page
    page_data = all_table_data[start_idx:end_idx]
    
    cursor.execute('SELECT COUNT(*) as total FROM attendance_logs')
    total_logs = cursor.fetchone()['total']
    cursor.execute('SELECT COUNT(*) as total FROM users')
    total_users = cursor.fetchone()['total']
    
    conn.close()
    
    return render_template('index.html', 
                         logs=page_data,
                         page=page,
                         total_pages=total_pages,
                         total_count=total_count,
                         date_from=date_from,
                         date_to=date_to,
                         selected_floor=selected_floor,
                         search_user=search_user,
                         total_logs=total_logs,
                         total_users=total_users)

@app.route('/save_record', methods=['POST'])
@license_required
def save_record():
    if 'access_code' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    if session.get('role') != 'admin':
        return jsonify({'error': 'Admin required'}), 403
    
    try:
        data = request.json
        in_id = data.get('in_id')
        out_id = data.get('out_id')
        chk_in = data.get('chk_in', '').strip()
        chk_out = data.get('chk_out', '').strip()
        remark_in = data.get('remark_in', '').strip()
        remark_out = data.get('remark_out', '').strip()
        status_override = data.get('status', '').strip()
        
        if not in_id and not out_id:
            return jsonify({'error': 'No record ID'}), 400
        
        conn = get_db()
        cursor = conn.cursor()
        updated = []
        
        # Update IN record
        if in_id and chk_in:
            cursor.execute('SELECT timestamp FROM attendance_logs WHERE id = ?', (in_id,))
            existing = cursor.fetchone()
            if existing:
                current_date = datetime.strptime(existing['timestamp'], '%Y-%m-%d %H:%M:%S').date()
                day_name = current_date.strftime('%A')
                try:
                    time_obj = datetime.strptime(chk_in, '%H:%M').time()
                    new_timestamp = datetime.combine(current_date, time_obj).strftime('%Y-%m-%d %H:%M:%S')
                    if not remark_in:
                        remark_in = calculate_remark_in(chk_in, day_name)
                    cursor.execute('''
                        UPDATE attendance_logs 
                        SET timestamp = ?, remark_in = ?, is_manually_edited = 1, last_edited_at = CURRENT_TIMESTAMP
                        WHERE id = ?
                    ''', (new_timestamp, remark_in, in_id))
                    updated.append('In')
                except:
                    pass
        
        # Handle OUT record
        if chk_out:
            if out_id:
                cursor.execute('SELECT timestamp FROM attendance_logs WHERE id = ?', (out_id,))
                existing = cursor.fetchone()
                if existing:
                    current_date = datetime.strptime(existing['timestamp'], '%Y-%m-%d %H:%M:%S').date()
                    day_name = current_date.strftime('%A')
                    try:
                        time_obj = datetime.strptime(chk_out, '%H:%M').time()
                        new_timestamp = datetime.combine(current_date, time_obj).strftime('%Y-%m-%d %H:%M:%S')
                        if not remark_out:
                            remark_out = calculate_remark_out(chk_out, day_name)
                        cursor.execute('''
                            UPDATE attendance_logs 
                            SET timestamp = ?, remark_out = ?, is_manually_edited = 1, last_edited_at = CURRENT_TIMESTAMP
                            WHERE id = ?
                        ''', (new_timestamp, remark_out, out_id))
                        updated.append('Out')
                    except:
                        pass
            else:
                # No OUT record - CREATE one
                if in_id:
                    cursor.execute('SELECT user_id, device_floor, timestamp FROM attendance_logs WHERE id = ?', (in_id,))
                    in_record = cursor.fetchone()
                    if in_record:
                        current_date = datetime.strptime(in_record['timestamp'], '%Y-%m-%d %H:%M:%S').date()
                        day_name = current_date.strftime('%A')
                        try:
                            time_obj = datetime.strptime(chk_out, '%H:%M').time()
                            new_timestamp = datetime.combine(current_date, time_obj).strftime('%Y-%m-%d %H:%M:%S')
                            if not remark_out:
                                remark_out = calculate_remark_out(chk_out, day_name)
                            cursor.execute('''
                                INSERT INTO attendance_logs 
                                (user_id, device_floor, timestamp, punch_type, remark_out, is_manually_edited, last_edited_at)
                                VALUES (?, ?, ?, 1, ?, 1, CURRENT_TIMESTAMP)
                            ''', (in_record['user_id'], in_record['device_floor'], new_timestamp, remark_out))
                            updated.append('Out (created)')
                        except Exception as e:
                            print(f"Error creating OUT record: {e}")
        
        # Status override
        if status_override:
            if in_id:
                cursor.execute('''
                    UPDATE attendance_logs SET status_override = ?, is_manually_edited = 1
                    WHERE id = ?
                ''', (status_override, in_id))
            if out_id:
                cursor.execute('''
                    UPDATE attendance_logs SET status_override = ?, is_manually_edited = 1
                    WHERE id = ?
                ''', (status_override, out_id))
            updated.append('Status')
        
        conn.commit()
        conn.close()
        
        if updated:
            return jsonify({'success': True, 'message': f'Updated: {", ".join(updated)}'})
        return jsonify({'error': 'No changes'}), 400
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/users')
@license_required
def users():
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    search = request.args.get('search', '')
    floor = request.args.get('floor', '')
    status = request.args.get('status', '')
    
    conn = get_db()
    cursor = conn.cursor()
    
    query = '''
        SELECT u.*, 
               (SELECT COUNT(*) FROM attendance_logs a WHERE a.user_id = u.user_id AND a.device_floor = u.device_floor) as log_count
        FROM users u WHERE 1=1
    '''
    params = []
    
    if search:
        query += ' AND (u.user_id LIKE ? OR u.name LIKE ?)'
        params.extend([f'%{search}%', f'%{search}%'])
    if floor:
        query += ' AND u.device_floor = ?'
        params.append(floor)
    if status != '':
        query += ' AND u.is_active = ?'
        params.append(int(status))
    
    query += ' ORDER BY u.device_floor, u.name'
    cursor.execute(query, params)
    users = cursor.fetchall()
    
    cursor.execute('''
        SELECT device_floor, 
               COUNT(*) as total,
               SUM(CASE WHEN is_active = 1 THEN 1 ELSE 0 END) as active,
               SUM(CASE WHEN is_active = 0 THEN 1 ELSE 0 END) as inactive
        FROM users 
        GROUP BY device_floor
    ''')
    floor_stats = cursor.fetchall()
    
    total_active = 0
    total_inactive = 0
    for stat in floor_stats:
        total_active += stat['active']
        total_inactive += stat['inactive']
    
    conn.close()
    
    return render_template('users.html', 
                         users=users, 
                         floor_stats=floor_stats,
                         total_active=total_active,
                         total_inactive=total_inactive,
                         search_query=search,
                         selected_floor=floor,
                         selected_status=status)

@app.route('/toggle_user_status', methods=['POST'])
@license_required
def toggle_user_status():
    if 'access_code' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    if session.get('role') != 'admin':
        return jsonify({'error': 'Admin required'}), 403
    
    try:
        data = request.json
        user_id = data.get('user_id')
        device_floor = data.get('device_floor')
        action = data.get('action')
        
        if not user_id or not device_floor or not action:
            return jsonify({'error': 'Missing fields'}), 400
        
        new_status = 1 if action == 'activate' else 0
        
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE users SET is_active = ? 
            WHERE user_id = ? AND device_floor = ?
        ''', (new_status, user_id, device_floor))
        
        if cursor.rowcount == 0:
            conn.close()
            return jsonify({'error': 'User not found'}), 404
        
        conn.commit()
        conn.close()
        return jsonify({'success': True})
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/reports')
@license_required
def reports():
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    selected_employee = request.args.get('employee', '')
    emp_date_from = request.args.get('emp_date_from', '')
    emp_date_to = request.args.get('emp_date_to', '')
    
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('SELECT COUNT(*) as total FROM users WHERE is_active = 1')
    active_users = cursor.fetchone()['total']
    cursor.execute('SELECT COUNT(*) as total FROM users WHERE is_active = 0')
    inactive_users = cursor.fetchone()['total'] or 0
    total_users = active_users + inactive_users
    
    cursor.execute('SELECT COUNT(*) as total FROM attendance_logs')
    total_logs = cursor.fetchone()['total']
    
    cursor.execute('SELECT device_floor, COUNT(*) as count FROM users WHERE is_active = 1 GROUP BY device_floor')
    floor_users = {row['device_floor']: row['count'] for row in cursor.fetchall()}
    
    cursor.execute('SELECT device_floor, COUNT(*) as count FROM attendance_logs GROUP BY device_floor')
    floor_logs = {row['device_floor']: row['count'] for row in cursor.fetchall()}
    
    daily_stats = []
    for i in range(7):
        date = (datetime.now() - timedelta(days=i)).strftime('%Y-%m-%d')
        cursor.execute('''
            SELECT 
                COUNT(*) as total,
                SUM(CASE WHEN punch_type = 0 THEN 1 ELSE 0 END) as in_count,
                SUM(CASE WHEN punch_type = 1 THEN 1 ELSE 0 END) as out_count
            FROM attendance_logs
            WHERE DATE(timestamp) = ?
        ''', (date,))
        row = cursor.fetchone()
        daily_stats.append({
            'date': date,
            'total': row['total'] or 0,
            'in_count': row['in_count'] or 0,
            'out_count': row['out_count'] or 0
        })
    
    cursor.execute('SELECT user_id, name, device_floor FROM users WHERE is_active = 1 ORDER BY name')
    employees = cursor.fetchall()
    
    employee_logs = []
    employee_name = ''
    
    if selected_employee:
        parts = selected_employee.split('|')
        if len(parts) == 2:
            user_id = parts[0]
            device_floor = parts[1]
            cursor.execute('SELECT name FROM users WHERE user_id = ? AND device_floor = ?', (user_id, device_floor))
            user = cursor.fetchone()
            if user:
                employee_name = user['name']
            
            if emp_date_from and emp_date_to:
                logs = get_export_data_grouped(emp_date_from, emp_date_to, device_floor, user_id)
                employee_logs = [row for row in logs if row['user_id'] == user_id]
    
    conn.close()
    
    return render_template('reports.html',
                         active_tab='summary',
                         total_users=total_users,
                         active_users=active_users,
                         inactive_users=inactive_users,
                         total_logs=total_logs,
                         floor_users=floor_users,
                         floor_logs=floor_logs,
                         daily_stats=daily_stats,
                         employees=employees,
                         selected_employee=selected_employee,
                         employee_name=employee_name,
                         employee_logs=employee_logs,
                         emp_date_from=emp_date_from,
                         emp_date_to=emp_date_to)

@app.route('/activity')
@license_required
def activity():
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    selected_employee = request.args.get('employee', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    status_override = request.args.get('status_override', '')
    apply_all = request.args.get('apply_all', '')
    revert = request.args.get('revert', '')
    override_date = request.args.get('override_date', '')
    override_employee = request.args.get('override_employee', '')
    
    if not date_from and not date_to:
        date_to = datetime.now().strftime('%Y-%m-%d')
        date_from = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
    
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('SELECT user_id, name, device_floor FROM users WHERE is_active = 1 ORDER BY name')
    employees = cursor.fetchall()
    
    if status_override and apply_all == 'true' and override_date:
        try:
            cursor.execute('''
                UPDATE attendance_logs 
                SET status_override = ?, is_manually_edited = 1
                WHERE DATE(timestamp) = ?
            ''', (status_override, override_date))
            affected = cursor.rowcount
            conn.commit()
            flash(f'Status override "{status_override}" applied to {affected} records for {override_date}', 'success')
        except Exception as e:
            flash(f'Error: {e}', 'danger')
    
    if revert == 'true' and override_date:
        try:
            cursor.execute('''
                UPDATE attendance_logs 
                SET status_override = NULL
                WHERE DATE(timestamp) = ?
            ''', (override_date,))
            affected = cursor.rowcount
            conn.commit()
            flash(f'Status override reverted for {affected} records on {override_date}', 'success')
        except Exception as e:
            flash(f'Error: {e}', 'danger')
    
    raw_logs = []
    selected_employee_name = ''
    
    if selected_employee and date_from and date_to:
        parts = selected_employee.split('|')
        if len(parts) == 2:
            user_id = parts[0]
            device_floor = parts[1]
            
            cursor.execute('SELECT name FROM users WHERE user_id = ? AND device_floor = ?', (user_id, device_floor))
            user = cursor.fetchone()
            if user:
                selected_employee_name = user['name']
            
            query = """
                SELECT
                    DATE(a.timestamp) as date,
                    strftime('%H:%M', a.timestamp) as time,
                    a.punch_type
                FROM attendance_logs a
                WHERE a.user_id = ? AND a.device_floor = ?
                    AND DATE(a.timestamp) >= ? AND DATE(a.timestamp) <= ?
                ORDER BY a.timestamp ASC
            """
            cursor.execute(query, (user_id, device_floor, date_from, date_to))
            punches = cursor.fetchall()
            
            punches_by_date = {}
            for punch in punches:
                date_key = punch['date']
                if date_key not in punches_by_date:
                    punches_by_date[date_key] = {'in': None, 'out': None}
                if punch['punch_type'] == 0:
                    punches_by_date[date_key]['in'] = punch['time']
                else:
                    punches_by_date[date_key]['out'] = punch['time']
            
            start_date = datetime.strptime(date_from, '%Y-%m-%d')
            end_date = datetime.strptime(date_to, '%Y-%m-%d')
            current_date = start_date
            while current_date <= end_date:
                date_str = current_date.strftime('%Y-%m-%d')
                date_formatted = current_date.strftime('%d-%b-%Y')
                day_name = current_date.strftime('%A')
                
                in_time = punches_by_date[date_str]['in'] if date_str in punches_by_date else None
                out_time = punches_by_date[date_str]['out'] if date_str in punches_by_date else None
                
                raw_logs.append({
                    'user_id': user_id,
                    'user_name': selected_employee_name,
                    'device_floor': device_floor,
                    'date': date_formatted,
                    'day': day_name,
                    'time': in_time if in_time else '--',
                    'punch_type': 'IN'
                })
                
                raw_logs.append({
                    'user_id': user_id,
                    'user_name': selected_employee_name,
                    'device_floor': device_floor,
                    'date': date_formatted,
                    'day': day_name,
                    'time': out_time if out_time else '--',
                    'punch_type': 'OUT'
                })
                
                current_date += timedelta(days=1)
    
    conn.close()
    
    return render_template('activity.html',
                         employees=employees,
                         selected_employee=selected_employee,
                         selected_employee_name=selected_employee_name,
                         date_from_activity=date_from,
                         date_to_activity=date_to,
                         raw_logs=raw_logs,
                         status_override=status_override,
                         override_date=override_date,
                         override_employee=override_employee)

@app.route('/export/excel')
@license_required
def export_excel():
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')
    
    if not date_from and not date_to:
        date_to = datetime.now().strftime('%Y-%m-%d')
        date_from = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
    
    data = get_export_data_grouped(date_from, date_to, selected_floor, search_user)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    
    headers = ['ID', 'Floor', 'Date', 'Day', 'Name', 'Chk In', 'Remarks In', 'Chk Out', 'Remarks Out', 'Status']
    ws.append(headers)
    
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1a237e", end_color="1a237e", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
    
    for row in data:
        ws.append([
            row['user_id'],
            row['device_floor'],
            row['date'],
            row['day'],
            row['user_name'],
            row['chk_in'],
            row['remark_in'],
            row['chk_out'],
            row['remark_out'],
            row['status']
        ])
    
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 35)
        ws.column_dimensions[column].width = adjusted_width
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='attendance.xlsx'
    )

@app.route('/export/csv')
@license_required
def export_csv():
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')
    
    if not date_from and not date_to:
        date_to = datetime.now().strftime('%Y-%m-%d')
        date_from = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
    
    data = get_export_data_grouped(date_from, date_to, selected_floor, search_user)
    
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['ID', 'Floor', 'Date', 'Day', 'Name', 'Chk In', 'Remarks In', 'Chk Out', 'Remarks Out', 'Status'])
    
    for row in data:
        writer.writerow([
            row['user_id'],
            row['device_floor'],
            row['date'],
            row['day'],
            row['user_name'],
            row['chk_in'],
            row['remark_in'],
            row['chk_out'],
            row['remark_out'],
            row['status']
        ])
    
    output.seek(0)
    return send_file(
        io.BytesIO(output.getvalue().encode('utf-8-sig')),
        mimetype='text/csv',
        as_attachment=True,
        download_name='attendance.csv'
    )

@app.route('/export/pdf')
@license_required
def export_pdf():
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')
    
    if not date_from and not date_to:
        date_to = datetime.now().strftime('%Y-%m-%d')
        date_from = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
    
    data = get_export_data_grouped(date_from, date_to, selected_floor, search_user)
    
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=20
    )
    
    content = []
    title = Paragraph("Attendance Report", title_style)
    content.append(title)
    content.append(Spacer(1, 10))
    
    table_data = []
    headers = ['ID', 'Floor', 'Date', 'Day', 'Name', 'Chk In', 'Remarks In', 'Chk Out', 'Remarks Out', 'Status']
    table_data.append(headers)
    
    for row in data:
        table_data.append([
            str(row['user_id']),
            str(row['device_floor']),
            str(row['date']),
            str(row['day']),
            str(row['user_name']),
            str(row['chk_in']),
            str(row['remark_in']),
            str(row['chk_out']),
            str(row['remark_out']),
            str(row['status'])
        ])
    
    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a237e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
    ]))
    
    content.append(table)
    doc.build(content)
    buffer.seek(0)
    
    return send_file(
        buffer,
        mimetype='application/pdf',
        as_attachment=True,
        download_name='attendance.pdf'
    )

@app.route('/export/sheets')
@license_required
def export_sheets():
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')
    
    if not date_from and not date_to:
        date_to = datetime.now().strftime('%Y-%m-%d')
        date_from = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
    
    data = get_export_data_grouped(date_from, date_to, selected_floor, search_user)
    
    user_data = {}
    for row in data:
        key = f"{row['user_id']}_{row['user_name']}"
        if key not in user_data:
            user_data[key] = {
                'user_id': row['user_id'],
                'user_name': row['user_name'],
                'records': []
            }
        user_data[key]['records'].append(row)
    
    wb = Workbook()
    
    summary_ws = wb.active
    summary_ws.title = "SUMMARY"
    summary_data = [
        ['Attendance Summary'],
        [''],
        ['Period:', date_from if date_from else 'N/A', 'to', date_to if date_to else 'N/A'],
        ['Total Users:', len(user_data)],
        ['Total Records:', len(data)],
    ]
    for row_idx, row_data in enumerate(summary_data, 1):
        for col_idx, value in enumerate(row_data, 1):
            cell = summary_ws.cell(row=row_idx, column=col_idx, value=value)
            if row_idx == 1:
                cell.font = Font(bold=True, size=14)
    summary_ws.column_dimensions['A'].width = 20
    
    for key, user_info in user_data.items():
        sheet_name = f"{user_info['user_id']}_{user_info['user_name']}"
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:31]
        ws = wb.create_sheet(title=sheet_name)
        
        headers = ['ID', 'Floor', 'Date', 'Day', 'Name', 'Chk In', 'Remarks In', 'Chk Out', 'Remarks Out', 'Status']
        ws.append(headers)
        
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1a237e", end_color="1a237e", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
        
        for row in user_info['records']:
            ws.append([
                row['user_id'],
                row['device_floor'],
                row['date'],
                row['day'],
                row['user_name'],
                row['chk_in'],
                row['remark_in'],
                row['chk_out'],
                row['remark_out'],
                row['status']
            ])
        
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 35)
            ws.column_dimensions[column].width = adjusted_width
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='attendance_sheets.xlsx'
    )

@app.route('/export/employee_excel')
@license_required
def export_employee_excel():
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    employee = request.args.get('employee', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    
    if not employee or not date_from or not date_to:
        flash('Please select employee and date range', 'warning')
        return redirect(url_for('reports'))
    
    parts = employee.split('|')
    if len(parts) != 2:
        flash('Invalid employee selection', 'warning')
        return redirect(url_for('reports'))
    
    user_id = parts[0]
    device_floor = parts[1]
    
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('SELECT name FROM users WHERE user_id = ? AND device_floor = ?', (user_id, device_floor))
    user = cursor.fetchone()
    employee_name = user['name'] if user else 'Unknown'
    
    data = get_export_data_grouped(date_from, date_to, device_floor, user_id)
    data = [row for row in data if row['user_id'] == user_id]
    
    conn.close()
    
    if not data:
        flash('No data found for this employee', 'warning')
        return redirect(url_for('reports'))
    
    wb = Workbook()
    ws = wb.active
    ws.title = f"{employee_name}"
    
    headers = ['Date', 'Day', 'Chk In', 'Remarks In', 'Chk Out', 'Remarks Out', 'Status']
    ws.append(headers)
    
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1a237e", end_color="1a237e", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
    
    for row in data:
        ws.append([
            row['date'],
            row['day'],
            row['chk_in'],
            row['remark_in'],
            row['chk_out'],
            row['remark_out'],
            row['status']
        ])
    
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 35)
        ws.column_dimensions[column].width = adjusted_width
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"{employee_name}_{date_from}_to_{date_to}.xlsx"
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

@app.route('/activity/export')
@license_required
def export_activity():
    return jsonify({'message': 'Activity export'}), 200

@app.context_processor
def utility_processor():
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('SELECT COUNT(*) as total FROM attendance_logs')
        total_logs = cursor.fetchone()['total']
        cursor.execute('SELECT COUNT(*) as total FROM users')
        total_users = cursor.fetchone()['total']
        conn.close()
        return {'total_logs': total_logs, 'total_users': total_users}
    except:
        return {'total_logs': 0, 'total_users': 0}

def get_status_override_for_date(override_date):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT status FROM status_overrides WHERE override_date = ?', (override_date,))
    result = cursor.fetchone()
    conn.close()
    return result['status'] if result else None

def apply_status_override(override_date, status):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('INSERT OR REPLACE INTO status_overrides (override_date, status) VALUES (?, ?)', (override_date, status))
    conn.commit()
    affected = cursor.rowcount
    conn.close()
    return affected

def revert_status_override(override_date):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM status_overrides WHERE override_date = ?', (override_date,))
    affected = cursor.rowcount
    conn.commit()
    conn.close()
    return affected

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
