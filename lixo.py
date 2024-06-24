emails = "Nome:|gisdoc@realvidaseguros"

print(emails.split("|")[0])

for email in emails.split("|"):
    print(email.find("Nome:"))