# fix-change-app
change the functions

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify
import sqlite3
from datetime import datetime, timedelta
import os
import json

app = Flask(__name__)
app.secret_key = 'backup-app-secret-key-2026'
app.config['SESSION_PERMANENT'] = False
app.config['SESSION_COOKIE_HTTPONLY'] = True

DATABASE = 'attendance.db'
ACCESS_CODE = 'alusman@123'
LICENSE_FILE = '.license'
EXPIRY_DAYS = 60
ACTIVATION_KEY = 'azeem@taizco.com'

# ============ LICENSE FUNCTIONS ============

def get_license_data():
    if os.path.exists(LICENSE_FILE):
        try:
            with open(LICENSE_FILE, 'r') as f:
                data = json.load(f)
            return data
        except:
            return None
    return None

def is_license_valid():
    data = get_license_data()
    if not data:
        return False
    try:
        expiry = datetime.strptime(data['expiry'], '%Y-%m-%d %H:%M:%S')
        now = datetime.now()
        if now <= expiry:
            return True
        else:
            return False
    except:
        return False

def license_required(f):
    def decorated_function(*args, **kwargs):
        if not is_license_valid():
            return redirect(url_for('license_expired'))
        return f(*args, **kwargs)
    decorated_function.__name__ = f.__name__
    return decorated_function

# ============ DATABASE FUNCTIONS ============

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

# ============ HELPER FUNCTIONS ============

def group_punches_by_day(logs):
    grouped = {}
    for log in logs:
        key = f"{log['user_id']}|{log['date_only']}|{log['device_floor']}"
        if key not in grouped:
            grouped[key] = {
                'user_id': log['user_id'],
                'user_name': log['user_name'] or 'Unknown',
                'device_floor': log['device_floor'],
                'date': log['date_only'],
                'day': log['day_name'],
                'in_punches': [],
                'out_punches': []
            }
        if log['punch_type'] == 0:
            grouped[key]['in_punches'].append(log['time'])
        else:
            grouped[key]['out_punches'].append(log['time'])
    return grouped

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

def calculate_status(in_time, out_time, day_name, has_in, has_out):
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
    try:
        time_obj = datetime.strptime(time_str, '%H:%M')
        return time_obj.strftime('%I:%M %p').lstrip('0')
    except:
        return time_str

# ============ ROUTES ============

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
            flash('License activated successfully!', 'success')
            return redirect(url_for('login'))
        else:
            flash('Invalid activation key', 'danger')
            return render_template('activate.html', error='Invalid activation key')
    return render_template('activate.html', error=None)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if not is_license_valid():
        return redirect(url_for('license_expired'))
    if 'access_code' in session:
        return redirect(url_for('index'))
    if request.method == 'POST':
        access_code = request.form.get('access_code', '').strip()
        if access_code == ACCESS_CODE:
            session['access_code'] = access_code
            session.permanent = False
            flash('Login successful!', 'success')
            return redirect(url_for('index'))
        else:
            flash('Invalid access code', 'danger')
            return render_template('login.html', error='Invalid access code')
    return render_template('login.html', error=None)

@app.route('/logout')
def logout():
    session.clear()
    flash('Logged out successfully', 'info')
    return redirect(url_for('login'))

# ============ MAIN INDEX ============

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
            a.is_manually_edited,
            log_id = a.id
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

    processed_logs = []
    for log in raw_logs:
        processed_logs.append({
            'user_id': log['user_id'],
            'user_name': log['user_name'] or 'Unknown',
            'device_floor': log['device_floor'],
            'date_only': log['date_only'],
            'day_name': weekday_map.get(log['weekday'], 'Unknown'),
            'time': log['time'],
            'punch_type': log['punch_type'],
            'log_id': log['log_id']
        })

    grouped = group_punches_by_day(processed_logs)

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
            if key in grouped:
                data = grouped[key]
                day_name = data['day']
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                date_formatted = date_obj.strftime('%d-%b-%Y')
                in_times = data['in_punches']
                out_times = data['out_punches']
                in_time = min(in_times) if in_times else None
                out_time = max(out_times) if out_times else None
                has_in = in_time is not None
                has_out = out_time is not None
                chk_in = get_display_time(in_time) if has_in else '--'
                chk_out = get_display_time(out_time) if has_out else '--'
                remark_in = calculate_remark_in(in_time, day_name) if has_in else ('Over Time' if day_name == 'Sunday' else 'Missing In')
                remark_out = calculate_remark_out(out_time, day_name) if has_out else ('Over Time' if day_name == 'Sunday' else 'Missing Out')
                status = calculate_status(in_time, out_time, day_name, has_in, has_out)
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
                    'log_id': data.get('log_id', None)
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
                        'log_id': None
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
                        'log_id': None
                    })

    all_table_data.sort(key=lambda x: datetime.strptime(x['date'], '%d-%b-%Y'), reverse=True)

    total_count = len(all_table_data)
    total_pages = (total_count + per_page - 1) // per_page if total_count > 0 else 1
    start_idx = (page - 1) * per_page
    end_idx = start_idx + per_page
    page_data = all_table_data[start_idx:end_idx]

    conn.close()

    print(f"DEBUG: total_count={total_count}, page_data length={len(page_data)}")
    return render_template('backup_index.html', 
                         logs=page_data,
                         page=page,
                         total_pages=total_pages,
                         total_count=total_count,
                         date_from=date_from,
                         date_to=date_to,
                         selected_floor=selected_floor,
                         search_user=search_user)

# ============ EDIT RECORD ============

@app.route('/save_record', methods=['POST'])
@license_required
def save_record():
    if 'access_code' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.json
    log_id = data.get('id')
    chk_in = data.get('chk_in', '').strip()
    chk_out = data.get('chk_out', '').strip()
    remark_in = data.get('remark_in', '').strip()
    remark_out = data.get('remark_out', '').strip()
    status_override = data.get('status', '').strip()
    
    if not log_id:
        return jsonify({'error': 'No record ID provided'}), 400
    
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('SELECT * FROM attendance_logs WHERE id = ?', (log_id,))
    existing = cursor.fetchone()
    if not existing:
        conn.close()
        return jsonify({'error': 'Record not found'}), 404
    
    updates = []
    params = []
    
    # If admin changed time, recalculate remarks (unless admin provided custom remarks)
    if chk_in:
        try:
            time_obj = datetime.strptime(chk_in, '%I:%M %p')
            chk_in_24 = time_obj.strftime('%H:%M')
            updates.append('remark_in = ?')
            if not remark_in:
                day_name = datetime.strptime(existing['timestamp'], '%Y-%m-%d %H:%M:%S').strftime('%A')
                remark_in = calculate_remark_in(chk_in_24, day_name)
            params.append(remark_in)
        except:
            if remark_in:
                updates.append('remark_in = ?')
                params.append(remark_in)
            else:
                updates.append('remark_in = ?')
                params.append(existing['remark_in'] or '')
    else:
        if remark_in:
            updates.append('remark_in = ?')
            params.append(remark_in)
    
    if chk_out:
        try:
            time_obj = datetime.strptime(chk_out, '%I:%M %p')
            chk_out_24 = time_obj.strftime('%H:%M')
            updates.append('remark_out = ?')
            if not remark_out:
                day_name = datetime.strptime(existing['timestamp'], '%Y-%m-%d %H:%M:%S').strftime('%A')
                remark_out = calculate_remark_out(chk_out_24, day_name)
            params.append(remark_out)
        except:
            if remark_out:
                updates.append('remark_out = ?')
                params.append(remark_out)
            else:
                updates.append('remark_out = ?')
                params.append(existing['remark_out'] or '')
    else:
        if remark_out:
            updates.append('remark_out = ?')
            params.append(remark_out)
    
    if status_override:
        updates.append('status_override = ?')
        params.append(status_override)
    
    updates.append('is_manually_edited = 1')
    updates.append('last_edited_at = CURRENT_TIMESTAMP')
    
    if not updates:
        conn.close()
        return jsonify({'error': 'No changes to save'}), 400
    
    params.append(log_id)
    query = f"UPDATE attendance_logs SET {', '.join(updates)} WHERE id = ?"
    
    try:
        cursor.execute(query, params)
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Record updated successfully'})
    except Exception as e:
        conn.rollback()
        conn.close()
        return jsonify({'error': str(e)}), 500

# ============ CONTEXT PROCESSOR ============

@app.context_processor
def utility_processor():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT COUNT(*) as total FROM attendance_logs')
    total_logs = cursor.fetchone()['total']
    cursor.execute('SELECT COUNT(*) as total FROM users')
    total_users = cursor.fetchone()['total']
    conn.close()
    return {'total_logs': total_logs, 'total_users': total_users}

# ============ MAIN ============

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5015)
