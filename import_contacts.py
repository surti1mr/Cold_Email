"""One-time script to bulk-insert existing Excel tracking data into MySQL."""

import sys
from pathlib import Path

# Make sure we can import database.py from the project folder
sys.path.insert(0, str(Path(__file__).resolve().parent))

import database

database.init_db()

records = [
    # (company_name, person_name, email, type_, status, notes)
    ("Deloitte",        "Greg Boehmer",              "gboehmer@deloitte.com",              "cmu_alumni",  "Coffee Chat Scheduled", "Prep for Call"),
    ("Deloitte",        "Bien, Katie",               "kbien@deloitte.com",                 "cmu_alumni",  "Other",                 "Out of office until sept/oct"),
    ("Deloitte",        "Szostek, Ryan",             "rszostek@deloitte.com",              "cmu_alumni",  "Ghosted",               ""),
    ("Thomas Reuters",  "Stover, Ryan",              "ryan.stover@thomsonreuters.com",      "cmu_alumni",  "Other",                 "Does not work with software engineering team. Can ask for referral."),
    ("Thomas Reuters",  "Sample, Jason R",           "jason.sample@thomsonreuters.com",     "cmu_alumni",  "Other",                 "Told me to visit career page"),
    ("HTCINC",          "Jason Grifka",              "jason.grifka@HTCinc.com",            "cmu_alumni",  "Replied",               "Referred to recruiter Samantha. Send follow-up."),
    ("Thomas Reuters",  "Samantha Lawson",           "samantha.lawson@HTCinc.com",         "cmu_alumni",  "Replied",               "Asked basic information"),
    ("Thomas Reuters",  "Shannon, Jesse (TR Marketing)", "jesse.shannon@thomsonreuters.com","cmu_alumni",  "Other",                 "Told me to add her name in referral. On leave till mid June."),
    ("Thomas Reuters",  "Bendelie, Mark (TR Product)","mark.bendele@thomsonreuters.com",   "cmu_alumni",  "Replied",               "Have given him 5 job references. Send follow-up."),
    ("BOSCHUSA",        "Feely, Catherine",          "catherine.feely@us.bosch.com",       "it_employee", "Other",                 "Out of office with no access to email returning Tuesday 2 June."),
    ("BOSCHUSA",        "Adamkowski, Paul",          "paul.adamkowski@us.bosch.com",       "it_employee", "Coffee Chat Scheduled", "Coffee Chat on June 10th. Prep for Call."),
    ("BOSCHUSA",        "Rummer, Ryan",              "llyan.Rummer@us.bosch.com",          "cmu_alumni",  "Replied",               "Shared my resume to the person recruiting for the role. Send follow-up."),
    ("BOSCHUSA",        "Buerkie, Stefan (RBNA/P)",  "stefan.buerkie@us.bosch.com",        "it_employee", "Other",                 "No longer with Bosch. Reach out to Caitlin Distelrath (Caitlin.Distelrath@us.bosch.com)."),
    ("BOSCHUSA",        "Barchett, Austin (PT/MBP1-NA)","austin.barchett@us.bosch.com",   "it_employee", "Replied",               "ATS will send email reply and ask to apply. Confirming interest in the role. Send follow-up."),
    ("BOSCHUSA",        "david.cook",                "david.cook@us.bosch.com",            "",            "Bounced",               ""),
    ("BOSCHUSA",        "sven.lanwer",               "sven.lanwer@us.bosch.com",           "",            "Delivered",             ""),
]

inserted = 0
for company_name, person_name, email, type_, status, notes in records:
    try:
        database.add_contact(
            company_name=company_name,
            person_name=person_name,
            email=email,
            type_=type_,
            status=status,
            notes=notes,
        )
        print(f"  OK  {person_name} <{email}>")
        inserted += 1
    except Exception as e:
        print(f"  FAIL  {person_name} <{email}> -- {e}")

print(f"\nDone: {inserted}/{len(records)} records inserted.")
