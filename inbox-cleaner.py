import imaplib

mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login("your_email", "your_password")
mail.select("inbox")
_, data = mail.search(None, '(BEFORE 01-Jan-2024)')
for num in data[0].split():
    mail.store(num, '+FLAGS', '\\Deleted')
mail.expunge()
mail.logout()