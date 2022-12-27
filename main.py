import mailbox
from collections import Counter
import re
import glob
import csv

search_path = "data"
blacklist = {
    'privacy@fem.digital',
    'bigius92@gmail.com',
    'caracciolo.biagio1996@gmail.com',
    'bigi96@live.it',
}
mail_re = re.compile(r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)", re.MULTILINE)
class BColors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


contacts = Counter()
details = {}

print(f"Searching mbox's from {search_path}...")
mboxes = glob.glob(f"{search_path}/**/*.mbox", recursive=True)
for mbox_path in mboxes:
    mbox = mailbox.mbox(mbox_path)
    message_count = len(mbox)
    print(f"Mailbox \"{BColors.BOLD}{mbox_path}{BColors.ENDC}\" contains {message_count} messages.")

    for n, message in enumerate(mbox, start=1):
        print(f"{n:<4}/{message_count:>4}", end='')
        date = message['Date']
        subject = message['Subject']
        for header in ('From', 'Sender', 'Reply-To'):
            addr = message.get(header)
            if addr:
                contacts[addr] += 1
                details[addr] = {'date': date, 'subject': subject}

        for part in message.walk():
            if part.get_content_type() in {"text/plain", "text/html"}:
                try:
                    body = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf8')
                except Exception as exc:
                    body = None
                    print(f"Can't parse body on {message['Message-ID']} ({exc}) (charset: {part.get_content_charset()}")
                break
        else:
            body = None
            print(f"No body on {message['Message-ID']}")
        
        if body:
            for addr in mail_re.findall(body):
                addr = addr.strip(".")
                contacts[addr] += 1
                details[addr] = {'date': date, 'subject': subject}
        
        print("\b"*9, end='', flush=True)
    print()

for addr in blacklist:
    try:
        del contacts[addr]
    except KeyError:
        pass

print("Most commond addresses:")
for addr, count in contacts.most_common(5):
    print(f"{addr:<50} • {count:>4}")

with open("results.csv", 'w') as handler:
    csvfile = csv.writer(handler)
    csvfile.writerow(['Indirizzo', '#', 'Ultimo contatto', 'Ultimo soggetto'])
    for addr in contacts:
        csvfile.writerow([addr, contacts[addr], details[addr]['date'], details[addr]['subject']])
