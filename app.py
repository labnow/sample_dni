from flask import Flask, render_template, url_for, request, redirect, send_file
from werkzeug.utils import secure_filename
import sqlite3
from myfunctions import csv2database, json_to_python, pivot_travel, records2excel, dictionary, csv2database, excel_summary, excel_effort_summary, exceL_travel_summary, excel_to_json, check_data
import os
import zipfile
from flask_wtf import FlaskForm
from wtforms import StringField, SubmitField, PasswordField, BooleanField, ValidationError, TextAreaField
from wtforms.validators import DataRequired, EqualTo, Length
from wtforms.widgets import TextArea

app = Flask(__name__)
# app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///test.db'
# db = SQLAlchemy(app)
    
@app.route('/')
def index():
    return render_template('index.html')
@app.route('/admin')
def admin():
    return render_template('index.html')

@app.route('/update/<int:row_num>', methods=['GET', 'POST'])
def update(row_num):
    import json
    with open('all_dict.json', 'r') as j:
        json_data = json.load(j)
        # t_data = json_data['data'] # will update value of this dictionary
        record_to_update = json_data['data'][str(row_num)]
        header = json_data['header']
        header[-4] = 'terms'
        header[-3] = 'stop_date'
        header[-2] = 'invoice_date'
        header[-1] = 'status'
        record_to_update_with_header = dict(zip(header, record_to_update))

    if request.method == "POST":
        record_to_update[-4] = request.form['terms']
        record_to_update[-3] = request.form['sop_date']
        record_to_update[-2] = request.form['invoice_date']
        record_to_update[-1] = request.form['status']

        json_data['data'][str(row_num)] = record_to_update

        with open('all_dict.json', 'w') as json_file:
            json.dump(json_data, json_file, ensure_ascii=False, indent=4)

        return redirect('/show_data/all')
    else:
        return render_template("update2.html", record_to_update_with_header = record_to_update_with_header, id = row_num)

@app.route('/read_excel', methods=['POST', 'GET'])
def read_excel():
    import datetime
    if request.method == 'POST':
        file = request.files['file']
        now = str(datetime.datetime.now())[:19]
        now = now.replace(":","_")
        now = now.replace(" ", "_")
        filename = os.path.join('uploads', now + '.xlsx')
        file.save(filename)
        status_to_return = excel_to_json(filename, request.form['sheet_name'], int(request.form['header_row']), int(request.form['data_start_from_row']), int(request.form['data_end_at_row']), int(request.form['column_start']), int(request.form['column_end']), request.form['base_month'])
        return '<h1>Looks Good!<h1><p>Add records to {}</p>'.format(status_to_return)
    else:
        return render_template('read_excel.html')

@app.route('/show_data/<nt_id>')
def show_data(nt_id):
    import json
    with open('all_dict.json', 'r') as j:
        json_data = json.load(j)

    if json_data:
        # get distinct nt nt_id <-> nt_name pair
        distinct_nt_id = {'all':'_all_'}
        col_num_of_nt_name = json_data['header'].index('Responsible Sales')
        for key, value in json_data['data'].items():
            distinct_nt_id[secure_filename(value[col_num_of_nt_name])] = value[col_num_of_nt_name]
        # get table header
        data_header = json_data['header']
        # get data insight
        data_insight = check_data('all_dict.json', distinct_nt_id[nt_id])
        # get table data
        if nt_id == 'all':
            data_for_jinja = json_data['data']
        else:
            data_for_jinja = dict()
            for key, value in json_data['data'].items():
                if value[col_num_of_nt_name] == distinct_nt_id[nt_id]:
                    data_for_jinja[key] = value
        return render_template('show_data.html', distinct_nt_id=distinct_nt_id, data_for_jinja=data_for_jinja, nt_id=nt_id, data_header=data_header, data_insight=data_insight)
        # return render_template('show_data.html', distinct_nt_id=distinct_nt_id, nt_id=nt_id)

    else:
        return '<h1>Empty json data, please trouble shoot</h1>'

@app.route('/show_data/<engineerName>/<month>', methods=['POST', 'GET'])
def show(engineerName=None, month=None):
    if request.method == 'POST':
        engineerName = request.form['engineers']
        month = request.form['months']
        re_url = '/show_data/'+ engineerName + '/' + month
        try:
            return redirect(re_url)
        except:
            return 'There was an issue adding your task'
    else:
        conn = sqlite3.connect('audi.sqlite')
        cur = conn.cursor()
        engineers = cur.execute('SELECT DISTINCT engineerName FROM (SELECT DISTINCT engineerName FROM effort UNION SELECT DISTINCT engineerName FROM travel AS foo)').fetchall()
        datas = []
        datas2 = []
        if engineerName:
            data_effort = cur.execute('SELECT id, package, date, startTime, endTime, workingHours, overtime, location, worklog FROM effort WHERE engineerName=? AND strftime("%m", date)=? ORDER BY date DESC', (engineerName, month)).fetchall()
            data_travel = cur.execute('SELECT id, date, type, city, description, invoiceType, price FROM travel WHERE engineerName=? AND strftime("%m", date)=? ORDER BY date DESC', (engineerName, month)).fetchall()
            # data_effort_str = int2str(data_effort, 'effort')
            # data_travel_str = int2str(data_travel, 'travel')
        conn.close()
        months = ['01','02','03','04','05','06','07','08','09','10','11','12']
        return render_template('show_data.html', data_effort=data_effort, data_travel=data_travel, engineers=engineers, months=months, engineer=engineerName, month=month)

@app.route('/excel/travel/<engineerName>/<month>')
def excel(engineerName, month):
    columns_travel_interested = ['date','type','city','description','invoiceType','price']
    file_to_return = records2excel('travel', engineerName, month, columns_travel_interested)   
    return send_file('persist/excels/' + file_to_return, mimetype = 'xlsx', download_name= file_to_return, as_attachment = True)

@app.route('/excel/effort/<engineerName>/<month>')
def excel_effort(engineerName, month):
    columns_effort_interested = ['package','date','startTime','endTime','workingHours','overtime','location','worklog']
    file_to_return = records2excel('effort', engineerName, month, columns_effort_interested)   
    return send_file('persist/excels/' + file_to_return, mimetype = 'xlsx', download_name= file_to_return, as_attachment = True)

@app.route('/excel/downloadAll')
def downloadAll():
    columns_effort_interested = ['package','date','startTime','endTime','workingHours','overtime','location','worklog']
    columns_travel_interested = ['date','type','city','description','invoiceType','price']
    # list_genetated_excel = []
    conn = sqlite3.connect('audi.sqlite')
    cur = conn.cursor()
    # all effort excels - (engineerName - month) Combination
    query_sql = 'SELECT DISTINCT engineerName, strftime("%m", date) FROM effort'
    engineerName_month_effort = cur.execute(query_sql).fetchall()
    for item in engineerName_month_effort:
        records2excel('effort', item[0], item[1], columns_effort_interested)

    # all travel excels - (engineerName - month) Combination
    query_sql = 'SELECT DISTINCT engineerName, strftime("%m", date) FROM travel'
    engineerName_month_travel = cur.execute(query_sql).fetchall()
    for item in engineerName_month_travel:
        records2excel('travel', item[0], item[1], columns_travel_interested)

    with zipfile.ZipFile('all_reports.zip', 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk('persist/excels/', topdown=False):
            for file in files:
                zf.write('persist/excels/' + file)
        zf.close()

    return send_file('all_reports.zip', mimetype = 'zip', download_name= 'all_reports.zip', as_attachment = True)

@app.route('/delete/<int:id>')
def delete(id):
    return render_template('404.html')

@app.route('/pivot/travel/<month>', methods=['POST', 'GET'])
def pivot(month):
    if request.method == 'POST':
        file_to_return = pivot_travel(month, 'excel')
        return send_file(file_to_return, mimetype = 'xlsx', download_name= file_to_return, as_attachment = True)
    else:
        filename = pivot_travel(month, 'html')
        return render_template('pivot.html', filename=filename)
    
@app.route('/new_record', methods=['POST', 'GET'])
def new_record():
    if request.method == 'POST':
        try:
            if request.form['form-name'] == 'one-effort':
                myPackage = dictionary('package', request.form['package'], 1)
                print(myPackage, request.form['package'])
                conn = sqlite3.connect('audi.sqlite')
                cur = conn.cursor()
                myQuery = 'INSERT INTO effort (package,date,engineerName,startTime,endTime,workingHours,overtime,location,worklog) VALUES (?,?,?,?,?,?,?,?,?)'
                cur.execute(myQuery, (myPackage, request.form['date'], request.form['engineerName'], request.form['startTime'], request.form['endTime'], request.form['workingHours'], request.form['overtime'], request.form['location'], request.form['worklog']))
                conn.commit()
                conn.close()
                return '<h1>Looks Good!<h1>'
            elif request.form['form-name'] == 'one-travel':
                myType = dictionary('type', request.form['type'], 1)
                myInvoicetype = dictionary('invoiceType', request.form['invoiceType'], 1)
                conn = sqlite3.connect('audi.sqlite')
                cur = conn.cursor()
                myQuery = 'INSERT INTO travel (engineerName,date,type,city,description,invoiceType,price) VALUES (?,?,?,?,?,?,?)'
                cur.execute(myQuery, (request.form['engineerName'], request.form['date'], myType, request.form['city'], request.form['description'], myInvoicetype, request.form['price']))
                conn.commit()
                conn.close()
                return '<h1>Looks Good!<h1>'
            else:
                file = request.files['file']
                filename = os.path.join('persist/uploads', secure_filename(file.filename))
                file.save(filename)
                csv2database(filename, request.form['tableName'])
                return '<h1>Looks Good!<h1>'
        except Exception as e:
            myStr = '<h1>Error add new record!</h1><p>' + str(e) + '</p>'
            return myStr
    else:
        return render_template('new_record.html')

@app.route('/summary/<tableName>/<month>')
def summary(tableName, month):
    excel_summary(tableName, month)
    columns_effort_interested = ['package','date','startTime','endTime','workingHours','overtime','location','worklog']
    file_to_return = records2excel('effort', engineerName, month, columns_effort_interested)   
    return send_file('persist/excels/' + file_to_return, mimetype = 'xlsx', download_name= file_to_return, as_attachment = True)

@app.route('/summary/effort/<month>')
def summary_effort(month):
    if month not in ['06', '07', '08', '09']:
        return '<p>No records found</p>'
    file_to_return = excel_effort_summary(month)
    return send_file(file_to_return, mimetype = 'xlsx', download_name= file_to_return, as_attachment = True)

@app.route('/summary/travel/<month>')
def summary_travel(month):
    if month not in ['06', '07', '08']:
        return '<p>No records found</p>'
    file_to_return = exceL_travel_summary(month)
    return send_file(file_to_return, mimetype = 'xlsx', download_name= file_to_return, as_attachment = True)
a = 'abc'
class RecordForm(FlaskForm):

    def __init__(self):
        print(a)

    # terms = StringField("Name", validators=[DataRequired()])
    # username = StringField("Username", validators=[DataRequired()])
    # email = StringField("Email", validators=[DataRequired()])
    # favorite_color = StringField("Favorite Color")
    # about_author = TextAreaField("About Author")
    # password_hash = PasswordField('Password', validators=[DataRequired(), EqualTo('password_hash2', message='Passwords Must Match!')])
    # password_hash2 = PasswordField('Confirm Password', validators=[DataRequired()])
    # submit = SubmitField("Submit")
    # abcd=self.abc

if __name__ == "__main__":
    # app.run(debug=True)
    # app.run()
    app.run(debug=True, port=3000, host='0.0.0.0')