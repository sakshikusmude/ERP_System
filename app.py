from flask import Flask, request, render_template, redirect, url_for, flash, session
import pandas as pd
import smtplib
from email.mime.text import MIMEText
import io

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # Needed for session and flash messages

def send_email(to_email, subject, message):
    from_email = "sakshi.22210253@viit.ac.in"
    password = "ssakshikk"

    msg = MIMEText(message)
    msg['Subject'] = subject
    msg['From'] = from_email
    msg['To'] = to_email

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(from_email, password)
            server.sendmail(from_email, to_email, msg.as_string())
        print(f"Reminder email sent successfully to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}. Error: {str(e)}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        student_file = request.files.get('student_file')
        fees_file = request.files.get('fees_file')

        if not student_file or not fees_file:
            flash('Please upload both student and fees files.')
            return redirect(request.url)

        try:
            student_df = pd.read_excel(student_file)
            fees_df = pd.read_excel(fees_file)

            print(f"Student DataFrame columns: {student_df.columns}")
            print(f"Fees DataFrame columns: {fees_df.columns}")

            merged_df = pd.merge(student_df, fees_df, on='PRN')
            merged_df['Remaining fees'] = merged_df['Actual Fees'] - merged_df['Paid fees']
            merged_df['Fees Status'] = merged_df['Remaining fees'].apply(lambda x: 'Not Paid' if x > 0 else 'Paid')

            # Store dataframes in the session
            session['merged_df'] = merged_df.to_html(classes='data', header="true", index=False)

            return redirect(url_for('students'))
        except Exception as e:
            flash(f"Error processing files: {e}")
            return redirect(request.url)
    
    return render_template('upload.html')

@app.route('/students')
def students():
    merged_df_html = session.get('merged_df')
    if not merged_df_html:
        flash('No data to display.')
        return redirect(url_for('upload'))

    return render_template('students.html', tables=merged_df_html)

@app.route('/send_reminders', methods=['POST'])
def send_reminders():
    try:
        # Retrieve dataframe from the session
        merged_df = pd.read_html(session.get('merged_df'))[0]  # Read HTML back to DataFrame

        reminders_sent = False
        for index, row in merged_df.iterrows():
            if row['Fees Status'] == 'Not Paid':
                message = f"Dear {row['Name']},\n\nPlease pay your remaining fees of {row['Remaining fees']}.\n\nRegards,\nFees Reminder System"
                send_email(row['email'], 'Fees Reminder', message)
                reminders_sent = True
        
        if reminders_sent:
            flash('Reminders sent successfully.')
        
        return redirect(url_for('students'))  # Redirect back to students page
    except Exception as e:
        flash(f"Error sending reminders: {e}")
        return redirect(url_for('upload'))



if __name__ == '__main__':
    app.run(debug=True)
