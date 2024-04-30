import pandas as pd
import mysql.connector
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def get_database_credentials():
    host = input("Enter MySQL host: ")
    username = input("Enter MySQL username: ")
    port = input("Enter MySQL port (default is usually 3306): ")
    password = input("Enter MySQL password: ")
    database = input("Enter MySQL database name: ")
    return host, port, username, password, database

def connect_to_database(host, port, username, password, database):
    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=username,
            password=password,
            database=database
        )
        print("Connected to MySQL database successfully!")
        return connection
    except mysql.connector.Error as err:
        print("Error:", err)
        return None


def run_query(connection, query):
    try:
        cursor = connection.cursor()
        cursor.execute(query)
        result = cursor.fetchall()
        return cursor, result
    except mysql.connector.Error as err:
        print("Error:", err)
        return None, None

def save_to_excel(cursor, result):
    if not result:
        print("No data to save.")
        return
    df = pd.DataFrame(result, columns=[x[0] for x in cursor.description])
    df.to_excel("query_result.xlsx", index=False)
    print("Result saved to query_result.xlsx")
    return "query_result.xlsx"

def send_email(filename):
    sender_email = input("Enter your email address: ")
    receiver_email = input("Enter recipient email address: ")
    password = input("Enter your email password: ")

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = "Query Results"

    body = "Please find attached the results of your query."
    msg.attach(MIMEText(body, 'plain'))

    attachment = open(filename, "rb")

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sender_email, password)
    text = msg.as_string()
    server.sendmail(sender_email, receiver_email, text)
    server.quit()

    print("Email sent successfully!")

def main():
    host, port, username, password, database = get_database_credentials()
    connection = connect_to_database(host, port, username, password, database)
    if connection:
        query = input("Enter SQL query: ")
        cursor, result = run_query(connection, query)
        if result:
            filename = save_to_excel(cursor, result)
            send_email(filename)
        connection.close()


if __name__ == "__main__":
    main()
