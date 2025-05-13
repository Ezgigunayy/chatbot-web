from flask import Flask, render_template, request, redirect, session, url_for, flash
import mysql.connector
import pandas as pd
from config import Config
import re
import MySQLdb
from flask_mysqldb import MySQL

app = Flask(__name__)
app.secret_key = "your_secret_key"

app.config['MYSQL_HOST'] = 'localhost'
app.config['MYSQL_USER'] = 'root'
app.config['MYSQL_PASSWORD'] = ''
app.config['MYSQL_DB'] = 'chatbot'

mysql = MySQL(app)

# Verileri Excel'e kaydetme fonksiyonu
def save_to_excel(data):
    try:
        # Mevcut Excel dosyasını yükle (varsa)
        try:
            df_existing = pd.read_excel('chatbot_data.xlsx', engine='openpyxl')
        except FileNotFoundError:
            # Eğer dosya yoksa yeni bir DataFrame oluştur
            df_existing = pd.DataFrame(columns=['Student_number', 'Question', 'Answer', 'Sentiment', 'Topic'])

        # Yeni verileri DataFrame'e ekle
        df_new = pd.DataFrame(data)
        df_all = pd.concat([df_existing, df_new], ignore_index=True)

        # Veriyi Excel dosyasına yaz
        df_all.to_excel('chatbot_data.xlsx', index=False, engine='openpyxl')
    except Exception as e:
        print(f"Veri kaydedilirken hata oluştu: {e}")

@app.route('/', methods=['GET', 'POST'])
def login():
    msg = ''
    if request.method == 'POST' and 'Student_number' in request.form and 'Password' in request.form:
        student_number = request.form['Student_number']
        password = request.form['Password']
        cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
        cursor.execute('SELECT * FROM users WHERE Student_number = %s AND Password = %s', (student_number, password))
        account = cursor.fetchone()
        if account:
            session['loggedin'] = True
            session['Student_number'] = account['Student_number']
            msg = 'Başarıyla giriş yapıldı!'
            return redirect('/chatbot')
        else:
            msg = 'Geçersiz kullanıcı adı veya şifre!'
    return render_template('login.html', msg=msg)

@app.route('/register', methods=['GET', 'POST'])
def register():
    msg = ''
    if request.method == 'POST' and 'Student_number' in request.form and 'Password' in request.form and 'name' in request.form:
        student_number = request.form['Student_number']
        password = request.form['Password']
        name = request.form['name']
        cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
        cursor.execute('SELECT * FROM users WHERE Student_number = %s', (student_number,))
        account = cursor.fetchone()
        if account:
            msg = 'Bu öğrenci numarası zaten kayıtlı!'
        elif not re.match(r'[0-9]+', student_number):
            msg = 'Öğrenci numarası sadece rakamlardan oluşmalıdır!'
        elif not re.match(r'[A-Za-z\s]+', name):
            msg = 'İsim sadece harflerden oluşmalıdır!'
        elif not student_number or not password or not name:
            msg = 'Tüm alanlar doldurulmalıdır!'
        else:
            cursor.execute('INSERT INTO users (Student_number, Password, Name) VALUES (%s, %s, %s)', (student_number, password, name))
            mysql.connection.commit()
            msg = 'Başarıyla kayıt oldunuz!'
            return redirect('/')
    return render_template('register.html', msg=msg)

@app.route('/chatbot', methods=['GET', 'POST'])
def chatbot():
    if 'loggedin' in session:
        msg = ''
        if request.method == 'POST':
            question = request.form.get('question')
            answer = request.form.get('answer')
            sentiment = request.form.get('sentiment')
            topic = request.form.get('topic')

            # Excel'e kaydedilecek veri
            data = {
                'Student_number': [session['Student_number']],
                'Question': [question],
                'Answer': [answer],
                'Sentiment': [sentiment],
                'Topic': [topic]
            }

            # Veriyi Excel dosyasına kaydet
            save_to_excel(data)

            # Veriyi MySQL'e de kaydet (Mevcut kodunuzda olduğu gibi)
            cursor = mysql.connection.cursor()
            cursor.execute(
                'INSERT INTO questions (ogrenci_ID, questions, answers, Feedback, topic) VALUES (%s, %s, %s, %s, %s)',
                (session['Student_number'], question, answer, sentiment, topic),
            )
            mysql.connection.commit()
            msg = 'Soru başarıyla eklendi!'
        return render_template('chatbot.html', msg=msg)
    return redirect('/')

@app.route('/reset_password', methods=['GET', 'POST'])
def reset_password():
    msg = ''
    if request.method == 'POST' and 'Student_number' in request.form and 'new_password' in request.form:
        student_number = request.form['Student_number']
        new_password = request.form['new_password']
        
        cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
        cursor.execute('SELECT * FROM users WHERE Student_number = %s', (student_number,))
        account = cursor.fetchone()
        
        if account:
            cursor.execute('UPDATE users SET Password = %s WHERE Student_number = %s', (new_password, student_number))
            mysql.connection.commit()
            msg = 'Şifreniz başarıyla sıfırlandı!'
        else:
            msg = 'Öğrenci numaranız bulunamadı!'
    
    return render_template('reset_password.html', msg=msg)

@app.route('/logout')
def logout():
    session.pop('loggedin', None)
    session.pop('Student_number', None)
    return redirect('/')


if __name__ == "_main_":
    app.run(debug=True)