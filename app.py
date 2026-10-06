from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import sqlite3, os, io, csv
from datetime import date, datetime
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE_DIR, 'finance.db')
app = Flask(__name__)
app.secret_key = 'personal-finance-tracker-demo-secret'

CATEGORIES = ['Food','Travel','Shopping','Bills','Education','Entertainment','Health','Rent','Salary','Freelance','Other']


def db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    con = db()
    con.executescript('''
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        budget REAL DEFAULT 0,
        savings_goal REAL DEFAULT 0,
        dark_mode INTEGER DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        trans_date TEXT NOT NULL,
        description TEXT NOT NULL,
        category TEXT NOT NULL,
        amount REAL NOT NULL,
        type TEXT NOT NULL CHECK(type IN ('Income','Expense')),
        recurring INTEGER DEFAULT 0,
        FOREIGN KEY(user_id) REFERENCES users(id)
    );
    ''')
    con.commit(); con.close()


def current_user():
    if 'user_id' not in session:
        return None
    con=db(); u=con.execute('SELECT * FROM users WHERE id=?',(session['user_id'],)).fetchone(); con.close(); return u


def login_required():
    return 'user_id' in session


def get_transactions(user_id, start=None, end=None, search='', category='', ttype=''):
    con=db(); q='SELECT * FROM transactions WHERE user_id=?'; args=[user_id]
    if start: q += ' AND trans_date>=?'; args.append(start)
    if end: q += ' AND trans_date<=?'; args.append(end)
    if search: q += ' AND (description LIKE ? OR category LIKE ?)'; args += [f'%{search}%', f'%{search}%']
    if category: q += ' AND category=?'; args.append(category)
    if ttype: q += ' AND type=?'; args.append(ttype)
    q += ' ORDER BY trans_date DESC, id DESC'
    rows=con.execute(q,args).fetchall(); con.close(); return rows


@app.route('/login', methods=['GET','POST'])
def login():
    if request.method=='POST':
        username=request.form['username'].strip(); password=request.form['password']
        con=db(); u=con.execute('SELECT * FROM users WHERE username=?',(username,)).fetchone(); con.close()
        if u and check_password_hash(u['password'], password):
            session['user_id']=u['id']; session['username']=u['username']; return redirect(url_for('dashboard'))
        flash('Invalid username or password.')
    return render_template('login.html')


@app.route('/register', methods=['GET','POST'])
def register():
    if request.method=='POST':
        username=request.form['username'].strip(); password=request.form['password']
        if len(username)<3 or len(password)<4:
            flash('Username must be 3+ characters and password 4+ characters.'); return render_template('register.html')
        con=db()
        try:
            cur=con.execute('INSERT INTO users(username,password) VALUES(?,?)',(username,generate_password_hash(password)))
            con.commit(); uid=cur.lastrowid
        except sqlite3.IntegrityError:
            con.close(); flash('Username already exists.'); return render_template('register.html')
        con.close(); session['user_id']=uid; session['username']=username; return redirect(url_for('dashboard'))
    return render_template('register.html')


@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('login'))


@app.route('/')
def dashboard():
    if not login_required(): return redirect(url_for('login'))
    u=current_user(); tx=get_transactions(u['id'])
    income=sum(float(t['amount']) for t in tx if t['type']=='Income')
    expense=sum(float(t['amount']) for t in tx if t['type']=='Expense')
    balance=income-expense
    month=datetime.now().strftime('%Y-%m')
    mtx=get_transactions(u['id'], month+'-01', month+'-31')
    mincome=sum(float(t['amount']) for t in mtx if t['type']=='Income')
    mexpense=sum(float(t['amount']) for t in mtx if t['type']=='Expense')
    cats={}
    for t in tx:
        if t['type']=='Expense': cats[t['category']]=cats.get(t['category'],0)+float(t['amount'])
    budget=float(u['budget'] or 0); goal=float(u['savings_goal'] or 0)
    budget_used=mexpense/budget*100 if budget else 0
    savings=income-expense
    alerts=[]
    if budget and mexpense>budget: alerts.append(f'Budget exceeded by ₹{mexpense-budget:.2f}')
    if budget and mexpense>=budget*0.8 and mexpense<=budget: alerts.append('You have used 80% or more of this month’s budget.')
    return render_template('index.html', user=u, transactions=tx[:10], income=income, expense=expense, balance=balance,
        mincome=mincome, mexpense=mexpense, cats=cats, budget=budget, budget_used=budget_used, goal=goal, savings=savings,
        alerts=alerts, today=date.today().isoformat(), categories=CATEGORIES)


@app.route('/add', methods=['POST'])
def add():
    if not login_required(): return redirect(url_for('login'))
    try: amount=float(request.form['amount'])
    except: flash('Enter a valid amount.'); return redirect(url_for('dashboard'))
    if amount<=0: flash('Amount must be greater than zero.'); return redirect(url_for('dashboard'))
    con=db(); con.execute('INSERT INTO transactions(user_id,trans_date,description,category,amount,type,recurring) VALUES(?,?,?,?,?,?,?)',
        (session['user_id'],request.form['trans_date'],request.form['description'].strip(),request.form['category'],amount,request.form['type'],1 if request.form.get('recurring') else 0)); con.commit(); con.close()
    flash('Transaction added successfully.'); return redirect(url_for('dashboard'))


@app.route('/delete/<int:tid>', methods=['POST'])
def delete(tid):
    if not login_required(): return redirect(url_for('login'))
    con=db(); con.execute('DELETE FROM transactions WHERE id=? AND user_id=?',(tid,session['user_id'])); con.commit(); con.close(); flash('Transaction deleted.'); return redirect(url_for('history'))


@app.route('/history')
def history():
    if not login_required(): return redirect(url_for('login'))
    filters={k:request.args.get(k,'') for k in ['search','category','type','start','end']}
    tx=get_transactions(session['user_id'],filters['start'],filters['end'],filters['search'],filters['category'],filters['type'])
    return render_template('history.html',transactions=tx,categories=CATEGORIES,filters=filters,user=current_user())


@app.route('/reports')
def reports():
    if not login_required(): return redirect(url_for('login'))
    tx=get_transactions(session['user_id'])
    monthly={}
    for t in tx:
        m=t['trans_date'][:7]; monthly.setdefault(m,{'Income':0,'Expense':0}); monthly[m][t['type']]+=float(t['amount'])
    return render_template('reports.html',user=current_user(),monthly=monthly,transactions=tx,categories=CATEGORIES)


@app.route('/settings', methods=['GET','POST'])
def settings():
    if not login_required(): return redirect(url_for('login'))
    if request.method=='POST':
        budget=float(request.form.get('budget') or 0); goal=float(request.form.get('goal') or 0); dark=1 if request.form.get('dark_mode') else 0
        con=db(); con.execute('UPDATE users SET budget=?, savings_goal=?, dark_mode=? WHERE id=?',(budget,goal,dark,session['user_id'])); con.commit(); con.close(); flash('Settings saved.'); return redirect(url_for('settings'))
    return render_template('settings.html',user=current_user())


def report_rows():
    tx=get_transactions(session['user_id']); return [[t['trans_date'],t['type'],t['category'],t['description'],float(t['amount'])] for t in tx]


@app.route('/download/excel')
def download_excel():
    if not login_required(): return redirect(url_for('login'))
    wb=Workbook(); ws=wb.active; ws.title='Transactions'; ws.append(['Date','Type','Category','Description','Amount'])
    for row in report_rows(): ws.append(row)
    ws.append([]); ws.append(['Generated by Personal Finance Tracker'])
    bio=io.BytesIO(); wb.save(bio); bio.seek(0)
    return send_file(bio,as_attachment=True,download_name='personal_finance_report.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/download/csv')
def download_csv():
    if not login_required(): return redirect(url_for('login'))
    s=io.StringIO(); w=csv.writer(s); w.writerow(['Date','Type','Category','Description','Amount']); w.writerows(report_rows())
    bio=io.BytesIO(s.getvalue().encode('utf-8')); return send_file(bio,as_attachment=True,download_name='personal_finance_report.csv',mimetype='text/csv')


@app.route('/download/pdf')
def download_pdf():
    if not login_required(): return redirect(url_for('login'))
    bio=io.BytesIO(); doc=SimpleDocTemplate(bio,pagesize=A4); styles=getSampleStyleSheet(); story=[Paragraph('Personal Finance Tracker - Report',styles['Title']),Spacer(1,12)]
    rows=[['Date','Type','Category','Description','Amount']] + report_rows()
    table=Table(rows,repeatRows=1); table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.lightgrey),('GRID',(0,0),(-1,-1),0.5,colors.grey),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),8)])); story.append(table); doc.build(story); bio.seek(0)
    return send_file(bio,as_attachment=True,download_name='personal_finance_report.pdf',mimetype='application/pdf')


@app.context_processor
def inject_globals():
    return {'logged_in':login_required(),'session_user':session.get('username')}


init_db()
if __name__=='__main__':
    app.run(debug=True)
