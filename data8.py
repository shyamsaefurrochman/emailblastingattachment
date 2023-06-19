from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from docx import Document
from tqdm import tqdm
import webbrowser
from tkinter import ttk

# Fungsi untuk mengirim email dengan melampirkan file Word dan mengganti nama di dalam isi Word
def send_email_with_attachment():
    smtp_server = smtp_entry.get()
    smtp_port = int(port_entry.get())
    sender_email = sender_entry.get()
    password = password_entry.get()
    email_column = email_column_entry.get()
    name_column = name_column_entry.get()
    npk_table = npk_table_entry.get()
    excel_file = excel_path_entry.get()
    template_file = template_path_entry.get()
    email_subject = subject_entry.get()
    email_message = message_entry.get("1.0", END)

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            
            df = pd.read_excel(excel_file)
            df = df.head(1000)
            
            success_emails = set()
            failure_count = 0
            
            for index, row in tqdm(df.iterrows(), total=len(df), desc="Sending Emails"):
                receiver_email = row[email_column]
                
                # Cek domain email
                if receiver_email.endswith('@bpjsketenagakerjaan.go.id'):
                    message = row[name_column]
                    npk = row[npk_table]
                    
                    doc = Document(template_file)
                    
                    for paragraph in doc.paragraphs:
                        if '{nama}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{nama}', message)
                        if '{npk}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{npk}', str(npk))
                    
                    modified_filename = f"{message}.docx"
                    doc.save(modified_filename)
                    
                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = receiver_email
                    msg['Subject'] = email_subject

                    # Tambahkan isi email
                    msg.attach(MIMEText(email_message, 'plain'))

                    # Buka file Word yang sudah dimodifikasi dan lampirkan
                    with open(modified_filename, 'rb') as file:
                        attachment = MIMEApplication(file.read(), _subtype='docx')
                        attachment.add_header('Content-Disposition', 'attachment', filename=modified_filename)
                        msg.attach(attachment)

                    try:
                        server.send_message(msg)
                        success_emails.add(receiver_email)
                    except Exception as e:
                        failure_count += 1

            # Tampilkan hasil pengiriman email
            messagebox.showinfo("Pengiriman Email Selesai", f"Total Email berhasil dikirim: {len(success_emails)}\nTotal Email gagal dikirim: {failure_count}")

    except Exception as e:
        messagebox.showerror("Kesalahan", f"Terjadi kesalahan saat mengirim email: {str(e)}")

def browse_excel_file():
    excel_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    excel_path_entry.delete(0, END)
    excel_path_entry.insert(END, excel_path)

def browse_template_file():
    template_path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    template_path_entry.delete(0, END)
    template_path_entry.insert(END, template_path)

def get_third_party_password():
    webbrowser.open('https://myaccount.google.com/security')
    messagebox.showinfo("Informasi", "Untuk menggunakan kata sandi aplikasi pihak ketiga, Anda perlu mengaktifkannya di pengaturan keamanan akun Google Anda. Silakan ikuti langkah-langkah berikut:\n\n1. Buka tautan berikut: https://myaccount.google.com/security\n\n2. Masuk dengan akun Google Anda.\n\n3. Pada bagian \"Pencarian\", cari \"Sandi Aplikasi\".\n\n4. Ikuti petunjuk untuk membuat kata sandi aplikasi pihak ketiga.")
# Fungsi untuk menampilkan pop-up cara kerja program
def show_help_popup():
    help_text = """
    Cara Kerja Program Email Blasting:
    
    1. Isi semua field yang diperlukan, seperti SMTP Server, Email Pengirim, Password, dsb.
    2. Pilih file Excel yang berisi data penerima email.
    3. Pilih template Word yang akan digunakan sebagai isi email.
    4. Tentukan kolom di file Excel yang berisi email, nama, dan tabel NPK.
    5. Isi subjek dan isi email yang akan dikirimkan.
    6. Klik tombol 'Kirim Email' untuk memulai proses pengiriman.
    7. Program akan membaca data dari file Excel, mengubah template Word sesuai dengan data, dan mengirim email dengan lampiran file Word yang telah dimodifikasi.
    8. Setelah proses selesai, akan muncul pop-up dengan informasi hasil pengiriman email.
    
    Pastikan Anda telah mengaktifkan opsi 'Kata Sandi Aplikasi Pihak Ketiga' pada akun Google Anda jika menggunakan SMTP Gmail.
    Untuk informasi lebih lanjut, klik tombol 'Cara Mendapatkan' di bawah kata sandi aplikasi pihak ketiga.
    """
    messagebox.showinfo("Cara Kerja Program", help_text)

# Membuat jendela GUI
window = Tk()
window.title("Email Blasting")

# Menghitung lebar jendela
window_width = 40

# Label dan Input SMTP Server
smtp_label = ttk.Label(window, text="SMTP Server:")
smtp_label.grid(row=0, column=0, padx=30, pady=10, sticky=W)
smtp_entry = ttk.Entry(window)
smtp_entry.grid(row=0, column=0, padx=250, sticky=W)
smtp_entry.insert(END, "smtp.gmail.com")

# Label dan Input SMTP Port
port_label = ttk.Label(window, text="SMTP Port:")
port_label.grid(row=1, column=0, padx=30, pady=10, sticky=W)
port_entry = ttk.Entry(window)
port_entry.grid(row=1, column=0, padx=250, sticky=W)
port_entry.insert(END, "587")

# Label dan Input Email Pengirim
sender_label = ttk.Label(window, text="Email Pengirim:")
sender_label.grid(row=2, column=0, padx=30, pady=10, sticky=W)
sender_entry = ttk.Entry(window)
sender_entry.grid(row=2, column=0, padx=250, sticky=W)

# Label dan Input Password
password_label = ttk.Label(window, text="Password:")
password_label.grid(row=3, column=0, padx=30, pady=10, sticky=W)
password_entry = ttk.Entry(window, show="*")
password_entry.grid(row=3, column=0, padx=250, sticky=W)

# Label dan Tombol untuk Mendapatkan Kata Sandi Aplikasi Pihak Ketiga
third_party_label = ttk.Label(window, text="Kata Sandi Aplikasi Pihak Ketiga:")
third_party_label.grid(row=4, column=0, padx=30, pady=10, sticky=W)
third_party_button = ttk.Button(window, text="Cara Mendapatkan", command=get_third_party_password)
third_party_button.grid(row=4, column=0, padx=250, sticky=W)

# Label dan Tombol untuk Pilih File Excel
excel_label = ttk.Label(window, text="File Excel:")
excel_label.grid(row=5, column=0, padx=30, pady=10, sticky=W)
excel_frame = ttk.Frame(window)
excel_frame.grid(row=5, column=0, padx=250, sticky=W)
excel_path_entry = ttk.Entry(excel_frame)
excel_path_entry.pack(side=LEFT)
browse_excel_button = ttk.Button(excel_frame, text="Browse", command=browse_excel_file)
browse_excel_button.pack(side=LEFT)

# Label dan Tombol untuk Pilih Template Word
template_label = ttk.Label(window, text="Template Word:")
template_label.grid(row=6, column=0, padx=30, pady=10, sticky=W)
template_frame = ttk.Frame(window)
template_frame.grid(row=6, column=0, padx=250, sticky=W)
template_path_entry = ttk.Entry(template_frame)
template_path_entry.pack(side=LEFT)
browse_template_button = ttk.Button(template_frame, text="Browse", command=browse_template_file)
browse_template_button.pack(side=LEFT)

# Label dan Input Kolom Email
email_column_label = ttk.Label(window, text="Kolom Email:")
email_column_label.grid(row=7, column=0, padx=30, pady=10, sticky=W)
email_column_entry = ttk.Entry(window)
email_column_entry.grid(row=7, column=0, padx=250, sticky=W)

# Label dan Input Kolom Nama
name_column_label = ttk.Label(window, text="Kolom Nama:")
name_column_label.grid(row=8, column=0, padx=30, pady=10, sticky=W)
name_column_entry = ttk.Entry(window)
name_column_entry.grid(row=8, column=0, padx=250, sticky=W)

# Label dan Input Tabel NPK
npk_table_label = ttk.Label(window, text="Tabel NPK:")
npk_table_label.grid(row=9, column=0, padx=30, pady=10, sticky=W)
npk_table_entry = ttk.Entry(window)
npk_table_entry.grid(row=9, column=0, padx=250, sticky=W)

# Label dan Input Subjek Email
subject_label = ttk.Label(window, text="Subjek Email:")
subject_label.grid(row=10, column=0, padx=30, pady=10, sticky=W)
subject_entry = ttk.Entry(window, width=window_width)
subject_entry.grid(row=10, column=0, padx=250, sticky=W)

# Label dan Input Isi Email
message_label = ttk.Label(window, text="Isi Email:")
message_label.grid(row=11, column=0, padx=30, pady=20, sticky=W)
message_entry = Text(window, width=window_width, height=10)
message_entry.grid(row=11, column=0, padx=250, pady=20, sticky=W)

# Tombol Kirim Email
send_button = ttk.Button(window, text="Kirim Email", command=send_email_with_attachment)
send_button.grid(row=12, column=0, padx=30, pady=10, sticky=N)

# Tanda tanya kecil di pojok kanan atas
question_mark_label = Label(window, text="?", font=("Arial", 16), fg="blue", cursor="question_arrow")
question_mark_label.place(x=window_width - 30, y=10, anchor=NW)

# Membuat event untuk menampilkan pop-up saat tanda tanya kecil diklik
question_mark_label.bind("<Button-1>", lambda e: show_help_popup())

window.mainloop()