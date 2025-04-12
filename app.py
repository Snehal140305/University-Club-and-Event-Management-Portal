from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import qrcode
import os

app = Flask(__name__)
EXCEL_FILE = "club_data.xlsx"

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE) as writer:
            pd.DataFrame(columns=["Club", "Description"]).to_excel(writer, sheet_name="Clubs", index=False)
            pd.DataFrame(columns=["Event", "Club", "Date", "Description"]).to_excel(writer, sheet_name="Events", index=False)
            pd.DataFrame(columns=["Event", "Attendee", "Status"]).to_excel(writer, sheet_name="Attendance", index=False)
            pd.DataFrame(columns=["Title", "Message"]).to_excel(writer, sheet_name="Announcements", index=False)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/dashboard')
def dashboard():
    return render_template('dashboard.html')

@app.route('/create_club', methods=["GET", "POST"])
def create_club():
    if request.method == "POST":
        club = request.form['club']
        desc = request.form['description']
        df = pd.read_excel(EXCEL_FILE, sheet_name="Clubs")
        df.loc[len(df)] = [club, desc]
        with pd.ExcelWriter(EXCEL_FILE, mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name="Clubs", index=False)
        return redirect("/dashboard")
    return render_template("create_club.html")

@app.route('/add_event', methods=["GET", "POST"])
def add_event():
    if request.method == "POST":
        name = request.form['event']
        club = request.form['club']
        date = request.form['date']
        desc = request.form['description']
        df = pd.read_excel(EXCEL_FILE, sheet_name="Events")
        df.loc[len(df)] = [name, club, date, desc]
        with pd.ExcelWriter(EXCEL_FILE, mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name="Events", index=False)

        return redirect("/dashboard")
    return render_template("add_event.html")

@app.route('/mark_attendance', methods=["GET", "POST"])
def mark_attendance():
    if request.method == "POST":
        event = request.form['event']
        attendee = request.form['attendee']
        status = request.form['status']
        df = pd.read_excel(EXCEL_FILE, sheet_name="Attendance")
        df.loc[len(df)] = [event, attendee, status]
        with pd.ExcelWriter(EXCEL_FILE, mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name="Attendance", index=False)

         # Generate QR code for attendance
        qr = qrcode.make(f"Attendance:{attendee,status}")
        qr.save(f"static/qrcodes/{attendee,status}_qr.png")

        return redirect("/dashboard")
    return render_template("mark_attendance.html")

@app.route('/announcement', methods=["GET", "POST"])
def announcement():
    if request.method == "POST":
        title = request.form['title']
        message = request.form['message']
        df = pd.read_excel(EXCEL_FILE, sheet_name="Announcements")
        df.loc[len(df)] = [title, message]
        with pd.ExcelWriter(EXCEL_FILE, mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name="Announcements", index=False)
        return redirect("/dashboard")
    return render_template("announcement.html")

if __name__ == "__main__":
    if not os.path.exists("static/qrcodes"):
        os.makedirs("static/qrcodes")
    init_excel()
    app.run(debug=True)
