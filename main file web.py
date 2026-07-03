from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file, make_response
import sqlite3
from datetime import datetime, timedelta
import csv
import io
import os
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT

app = Flask(__name__)
app.secret_key = 'web-app-secret-key-2026'
app.config['SESSION_PERMANENT'] = False
app.config['SESSION_COOKIE_HTTPONLY'] = True

DATABASE = 'attendance.db'
ACCESS_CODE = 'alusman@123'
LICENSE_FILE = '.license'
EXPIRY_DAYS = 60
ACTIVATION_KEY = 'azeem@taizco.com'

# ============ LICENSE FUNCTIONS ============

def get_license_data():
    """Read license file"""
    if os.path.exists(LICENSE_FILE):
        try:
            with open(LICENSE_FILE, 'r') as f:
                data = json.load(f)
            return data
        except:
            return None
    return None

def is_license_valid():
    """Check if license is valid"""
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

# ============ LICENSE CHECK DECORATOR ============

def license_required(f):
    """Decorator to check license before accessing routes"""
    def decorated_function(*args, **kwargs):
        if not is_license_valid():
            return redirect(url_for('license_expired'))
        return f(*args, **kwargs)
    decorated_function.__name__ = f.__name__
    return decorated_function

# ============ DATABASE FUNCTIONS ============

def get_db():
    """Get database connection"""
    conn = sqlite3.connect(DATABASE, timeout=20)
    conn.row_factory = sqlite3.Row
    return conn

def get_all_active_users(conn, floor_filter=None):
    """Get all active users, optionally filtered by floor"""
    cursor = conn.cursor()
    query = "SELECT user_id, name, device_floor FROM users WHERE is_active = 1"
    params = []
    if floor_filter:
        query += " AND device_floor = ?"
        params.append(floor_filter)
    query += " ORDER BY name"
    cursor.execute(query, params)
    return cursor.fetchall()

def get_user_name(conn, user_id, floor):
    """Get user name from database"""
    cursor = conn.cursor()
    cursor.execute('''
        SELECT name FROM users
        WHERE user_id = ? AND device_floor = ?
    ''', (user_id, floor))
    result = cursor.fetchone()
    return result['name'] if result else f"User {user_id}"

# ============ HELPER: GROUP PUNCHES BY DAY ============

def group_punches_by_day(logs):
    """ Group punches by user, date, and floor. Returns dict with key: user_id|date|floor """
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
    """Calculate IN remark based on time"""
    if not time_str or time_str == '--':
        return 'Missing In'

    # Sunday special handling
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
    """Calculate OUT remark based on time"""
    if not time_str or time_str == '--':
        return 'Missing Out'

    # Sunday special handling
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
    """Calculate combined status based on IN and OUT presence"""
    
    # Sunday handling
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
    """Convert 24-hour time to 12-hour AM/PM format"""
    if not time_str or time_str == '--':
        return '--'

    try:
        time_obj = datetime.strptime(time_str, '%H:%M')
        return time_obj.strftime('%I:%M %p').lstrip('0')
    except:
        return time_str

def get_sunday_display():
    """Return the Sunday merged display text"""
    return '_**SUNDAY**_ **'

# ============ EXPORT HELPER FUNCTION ============

def get_export_data_grouped(date_from, date_to, selected_floor, search_user):
    """Get grouped data for export in 10-column format"""
    conn = get_db()
    cursor = conn.cursor()

    # Get all active users
    users = get_all_active_users(conn, selected_floor)
    
    # Get logs for the date range
    query = '''
        SELECT 
            a.user_id,
            a.device_floor,
            DATE(a.timestamp) as date_only,
            strftime('%w', a.timestamp) as weekday,
            strftime('%H:%M', a.timestamp) as time,
            a.punch_type,
            u.name as user_name
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
            'punch_type': log['punch_type']
        })

    grouped = group_punches_by_day(processed_logs)

    # Build table data with ALL users for each day
    table_data = []
    
    # Get all dates in range
    if date_from and date_to:
        start = datetime.strptime(date_from, '%Y-%m-%d')
        end = datetime.strptime(date_to, '%Y-%m-%d')
        date_list = []
        current = start
        while current <= end:
            date_list.append(current.strftime('%Y-%m-%d'))
            current += timedelta(days=1)
    else:
        # If no date range, just use today
        date_list = [datetime.now().strftime('%Y-%m-%d')]

    # For each user, create records for each day
    for user in users:
        user_id = user['user_id']
        user_name = user['name']
        device_floor = user['device_floor']
        
        for date_str in date_list:
            # Check if user has punches on this date
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
                    'status': status,
                    'is_sunday': day_name == 'Sunday'
                })
            else:
                # No punches for this user on this date
                day_name = datetime.strptime(date_str, '%Y-%m-%d').strftime('%A')
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                date_formatted = date_obj.strftime('%d-%b-%Y')
                
                if day_name == 'Sunday':
                    # Sunday - show Weekly Off
                    table_data.append({
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
                        'is_sunday': True
                    })
                else:
                    # Working day - show missing
                    table_data.append({
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
                        'is_sunday': False
                    })

    # Sort by date descending
    table_data.sort(key=lambda x: datetime.strptime(x['date'], '%d-%b-%Y'), reverse=True)

    conn.close()
    return table_data

# ============ ROUTES ============

@app.route('/license-expired')
def license_expired():
    """Show license expired page"""
    return render_template('license_expired.html'), 403

@app.route('/activate', methods=['GET', 'POST'])
def activate():
    """Hidden activation page"""
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
    """Login page"""
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
    """Logout and clear session"""
    session.clear()
    flash('Logged out successfully', 'info')
    return redirect(url_for('login'))

# ============ MAIN INDEX ============

@app.route('/')
@license_required
def index():
    """Main attendance view page - 10 column format with ALL users"""
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

    # Get all active users (filtered by floor if selected)
    users = get_all_active_users(conn, selected_floor)
    
    # If search_user is provided, filter users
    if search_user:
        users = [u for u in users if search_user.lower() in u['name'].lower() or search_user in str(u['user_id'])]

    # Get logs for the date range
    query = '''
        SELECT 
            a.user_id,
            a.device_floor,
            DATE(a.timestamp) as date_only,
            strftime('%w', a.timestamp) as weekday,
            strftime('%H:%M', a.timestamp) as time,
            a.punch_type,
            u.name as user_name
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
            'punch_type': log['punch_type']
        })

    grouped = group_punches_by_day(processed_logs)

    # Build table data with ALL users for each day
    all_table_data = []
    
    # Get all dates in range
    start = datetime.strptime(date_from, '%Y-%m-%d')
    end = datetime.strptime(date_to, '%Y-%m-%d')
    date_list = []
    current = start
    while current <= end:
        date_list.append(current.strftime('%Y-%m-%d'))
        current += timedelta(days=1)

    # For each user, create records for each day
    for user in users:
        user_id = user['user_id']
        user_name = user['name']
        device_floor = user['device_floor']
        
        for date_str in date_list:
            # Check if user has punches on this date
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
                    'is_sunday': day_name == 'Sunday'
                })
            else:
                # No punches for this user on this date
                day_name = datetime.strptime(date_str, '%Y-%m-%d').strftime('%A')
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                date_formatted = date_obj.strftime('%d-%b-%Y')
                
                if day_name == 'Sunday':
                    # Sunday - show Weekly Off
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
                        'is_sunday': True
                    })
                else:
                    # Working day - show missing
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
                        'is_sunday': False
                    })

    # Sort by date descending
    all_table_data.sort(key=lambda x: datetime.strptime(x['date'], '%d-%b-%Y'), reverse=True)

    # Pagination
    total_count = len(all_table_data)
    total_pages = (total_count + per_page - 1) // per_page if total_count > 0 else 1
    
    # Get current page data
    start_idx = (page - 1) * per_page
    end_idx = start_idx + per_page
    page_data = all_table_data[start_idx:end_idx]

    conn.close()

    return render_template('index.html', 
                         logs=page_data,
                         page=page,
                         total_pages=total_pages,
                         total_count=total_count,
                         date_from=date_from,
                         date_to=date_to,
                         selected_floor=selected_floor,
                         search_user=search_user)

# ============ USERS ============

@app.route('/users')
@license_required
def users():
    """View all users"""
    if 'access_code' not in session:
        return redirect(url_for('login'))

    search_query = request.args.get('search', '')
    selected_floor = request.args.get('floor', '')
    selected_status = request.args.get('status', '')

    conn = get_db()
    cursor = conn.cursor()

    query = '''
        SELECT u.*, 
               (SELECT COUNT(*) FROM attendance_logs a WHERE a.user_id = u.user_id AND a.device_floor = u.device_floor) as log_count
        FROM users u
        WHERE 1=1
    '''
    params = []

    if search_query:
        query += ' AND (u.user_id LIKE ? OR u.name LIKE ?)'
        search_pattern = f'%{search_query}%'
        params.extend([search_pattern, search_pattern])

    if selected_floor:
        query += ' AND u.device_floor = ?'
        params.append(selected_floor)

    if selected_status:
        query += ' AND u.is_active = ?'
        params.append(selected_status)

    query += ' ORDER BY u.name'
    cursor.execute(query, params)
    users = cursor.fetchall()

    cursor.execute('SELECT COUNT(*) as total FROM users')
    total_users = cursor.fetchone()['total']

    cursor.execute('SELECT COUNT(*) as total FROM users WHERE is_active = 1')
    active_count = cursor.fetchone()['total']

    inactive_count = total_users - active_count

    conn.close()

    return render_template('users.html',
                         users=users,
                         search_query=search_query,
                         selected_floor=selected_floor,
                         selected_status=selected_status,
                         active_count=active_count,
                         inactive_count=inactive_count)

# ============ REPORTS ============

@app.route('/reports')
@license_required
def reports():
    """Reports and statistics page"""
    if 'access_code' not in session:
        return redirect(url_for('login'))

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute('SELECT COUNT(*) as total FROM users WHERE is_active = 1')
    total_users = cursor.fetchone()['total']

    cursor.execute('SELECT COUNT(*) as total FROM users WHERE is_active = 1')
    active_users = cursor.fetchone()['total']

    cursor.execute('SELECT COUNT(*) as total FROM users WHERE is_active = 0')
    inactive_users = cursor.fetchone()['total'] or 0

    cursor.execute('SELECT COUNT(*) as total FROM attendance_logs')
    total_logs = cursor.fetchone()['total']

    cursor.execute('SELECT MIN(timestamp) as min_date, MAX(timestamp) as max_date FROM attendance_logs')
    result = cursor.fetchone()
    date_from = result['min_date'] if result['min_date'] else 'N/A'
    date_to = result['max_date'] if result['max_date'] else 'N/A'

    if date_from != 'N/A' and date_to != 'N/A':
        days_covered = (datetime.strptime(date_to, '%Y-%m-%d %H:%M:%S') - 
                       datetime.strptime(date_from, '%Y-%m-%d %H:%M:%S')).days + 1
    else:
        days_covered = 0

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

    conn.close()

    return render_template('reports.html',
                         active_tab='summary',
                         total_users=total_users,
                         active_users=active_users,
                         inactive_users=inactive_users,
                         total_logs=total_logs,
                         date_from=date_from,
                         date_to=date_to,
                         days_covered=days_covered,
                         floor_users=floor_users,
                         floor_logs=floor_logs,
                         daily_stats=daily_stats,
                         employees=employees,
                         selected_employee='',
                         date_from_activity='',
                         date_to_activity='',
                         activity_logs=[])

# ============ EMPLOYEE ACTIVITY REPORT ============

@app.route('/activity')
@license_required
def activity():
    """Employee Activity Report - shows all punches without grouping"""
    if 'access_code' not in session:
        return redirect(url_for('login'))

    selected_employee = request.args.get('employee', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')

    if not date_from and not date_to:
        date_to = datetime.now().strftime('%Y-%m-%d')
        date_from = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute('SELECT user_id, name, device_floor FROM users WHERE is_active = 1 ORDER BY name')
    employees = cursor.fetchall()

    activity_logs = []
    selected_employee_name = ''

    if selected_employee:
        parts = selected_employee.split('|')
        if len(parts) == 2:
            user_id = parts[0]
            device_floor = parts[1]

            cursor.execute('SELECT name FROM users WHERE user_id = ? AND device_floor = ?', (user_id, device_floor))
            user = cursor.fetchone()
            if user:
                selected_employee_name = user['name']

            query = '''
                SELECT
                    user_id,
                    device_floor,
                    DATE(timestamp) as date,
                    strftime('%w', timestamp) as weekday,
                    strftime('%H:%M', timestamp) as time,
                    punch_type,
                    strftime('%Y-%m-%d', timestamp) as date_only
                FROM attendance_logs
                WHERE user_id = ?
                    AND device_floor = ?
                    AND DATE(timestamp) >= ?
                    AND DATE(timestamp) <= ?
                ORDER BY timestamp ASC
            '''
            cursor.execute(query, (user_id, device_floor, date_from, date_to))
            punches = cursor.fetchall()

            weekday_map = {
                '0': 'Sunday', '1': 'Monday', '2': 'Tuesday', '3': 'Wednesday',
                '4': 'Thursday', '5': 'Friday', '6': 'Saturday'
            }

            for punch in punches:
                day_name = weekday_map.get(punch['weekday'], 'Unknown')
                date_obj = datetime.strptime(punch['date_only'], '%Y-%m-%d')
                date_formatted = date_obj.strftime('%d-%b-%Y')

                time_formatted = punch['time']
                try:
                    time_obj = datetime.strptime(time_formatted, '%H:%M')
                    time_formatted = time_obj.strftime('%I:%M %p').lstrip('0')
                except:
                    pass

                activity_logs.append({
                    'user_id': punch['user_id'],
                    'device_floor': punch['device_floor'],
                    'date': date_formatted,
                    'day': day_name,
                    'user_name': selected_employee_name,
                    'time': time_formatted,
                    'punch_type': punch['punch_type']
                })

    conn.close()

    return render_template('activity.html',
                         employees=employees,
                         selected_employee=selected_employee,
                         date_from_activity=date_from,
                         date_to_activity=date_to,
                         activity_logs=activity_logs)

# ============ EXPORT FUNCTIONS ============

@app.route('/export/excel')
@license_required
def export_excel():
    """Export to Excel (10-column format)"""
    if 'access_code' not in session:
        return redirect(url_for('login'))

    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')

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
    """Export to CSV (10-column format)"""
    if 'access_code' not in session:
        return redirect(url_for('login'))

    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')

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
    """Export to PDF (10-column format)"""
    if 'access_code' not in session:
        return redirect(url_for('login'))

    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')

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

    title = Paragraph(f"🅐∞🅐 USMAN ENTERPRISE - Attendance Report", title_style)
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
    """Export to Excel with multiple sheets (10-column format)"""
    if 'access_code' not in session:
        return redirect(url_for('login'))

    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')

    data = get_export_data_grouped(date_from, date_to, selected_floor, search_user)

    user_data = {}
    for row in data:
        user_key = f"{row['user_id']}_{row['user_name']}"
        if user_key not in user_data:
            user_data[user_key] = {
                'user_id': row['user_id'],
                'user_name': row['user_name'],
                'records': []
            }
        user_data[user_key]['records'].append(row)

    wb = Workbook()

    summary_ws = wb.active
    summary_ws.title = "SUMMARY"

    summary_data = [
        ['USMAN ENTERPRISE - Attendance Summary'],
        [''],
        ['Period:', date_from if date_from else 'N/A', 'to', date_to if date_to else 'N/A'],
        ['Total Users:', len(user_data)],
        ['Total Logs:', len(data)],
        [''],
        ['Users by Floor:'],
    ]

    floor_counts = {}
    for row in data:
        floor = row['device_floor']
        floor_counts[floor] = floor_counts.get(floor, 0) + 1

    for floor, count in floor_counts.items():
        summary_data.append([floor, count])

    summary_data.append([''])
    summary_data.append(['All Rights Reserved ® 2026'])

    for row_idx, row_data in enumerate(summary_data, 1):
        for col_idx, value in enumerate(row_data, 1):
            cell = summary_ws.cell(row=row_idx, column=col_idx, value=value)
            if row_idx == 1:
                cell.font = Font(bold=True, size=14)
            elif row_idx == len(summary_data):
                cell.font = Font(italic=True, size=10)

    summary_ws.column_dimensions['A'].width = 25
    summary_ws.column_dimensions['B'].width = 15
    summary_ws.column_dimensions['C'].width = 15
    summary_ws.column_dimensions['D'].width = 15

    if user_data:
        for user_key, user_info in user_data.items():
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

    if not user_data:
        ws = wb.active
        ws.title = "No Data"
        ws.append(['No attendance records found for the selected filters'])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='attendance_sheets.xlsx'
    )

@app.route('/activity/export')
@license_required
def export_activity():
    """Export Employee Activity Report to Excel"""
    if 'access_code' not in session:
        return redirect(url_for('login'))

    selected_employee = request.args.get('employee', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')

    if not selected_employee or not date_from or not date_to:
        flash('Please select employee and date range', 'error')
        return redirect(url_for('activity'))

    parts = selected_employee.split('|')
    if len(parts) != 2:
        flash('Invalid employee selection', 'error')
        return redirect(url_for('activity'))

    user_id = parts[0]
    device_floor = parts[1]

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute('SELECT name FROM users WHERE user_id = ? AND device_floor = ?', (user_id, device_floor))
    user = cursor.fetchone()
    employee_name = user['name'] if user else 'Unknown'

    query = '''
        SELECT 
            user_id,
            device_floor,
            DATE(timestamp) as date,
            strftime('%w', timestamp) as weekday,
            strftime('%H:%M', timestamp) as time,
            punch_type,
            strftime('%Y-%m-%d', timestamp) as date_only
        FROM attendance_logs
        WHERE user_id = ? 
            AND device_floor = ?
            AND DATE(timestamp) >= ?
            AND DATE(timestamp) <= ?
        ORDER BY timestamp ASC
    '''
    cursor.execute(query, (user_id, device_floor, date_from, date_to))
    punches = cursor.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Activity"

    headers = ['ID', 'Floor', 'Date', 'Day', 'Name', 'Chk In', 'Chk Out']
    ws.append(headers)

    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1a237e", end_color="1a237e", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    weekday_map = {
        '0': 'Sunday', '1': 'Monday', '2': 'Tuesday', '3': 'Wednesday',
        '4': 'Thursday', '5': 'Friday', '6': 'Saturday'
    }

    for punch in punches:
        day_name = weekday_map.get(punch['weekday'], 'Unknown')
        date_obj = datetime.strptime(punch['date_only'], '%Y-%m-%d')
        date_formatted = date_obj.strftime('%d-%b-%Y')

        time_formatted = punch['time']
        try:
            time_obj = datetime.strptime(time_formatted, '%H:%M')
            time_formatted = time_obj.strftime('%I:%M %p').lstrip('0')
        except:
            pass

        chk_in = time_formatted if punch['punch_type'] == 0 else '--'
        chk_out = time_formatted if punch['punch_type'] == 1 else '--'

        ws.append([
            punch['user_id'],
            punch['device_floor'],
            date_formatted,
            day_name,
            employee_name,
            chk_in,
            chk_out
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
        adjusted_width = min(max_length + 2, 30)
        ws.column_dimensions[column].width = adjusted_width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    file_name = f"employee_activity_{employee_name}_{date_from}_to_{date_to}.xlsx"

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=file_name
    )

# ============ CONTEXT PROCESSOR ============

@app.context_processor
def utility_processor():
    """Make variables available to all templates"""
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
    app.run(debug=True, host='0.0.0.0', port=5000)
