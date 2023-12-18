import pandas
import math

def main():
    data =  pandas.ExcelFile('EmailCampaign.xlsx')
    for sheet in data.sheet_names:
        df = data.parse(sheet)
        for i in range(0, len(df['First Name'])):
            vcard = make_vcard(df['First Name'][i], df['Last Name'][i], df['Company'][i], df['Designation'][i], df['Number 1'][i], df['Number 2'][i], '', df['Email'][i])
            if vcard is not None:
                write_vcard(df['First Name'][i] + ".vcf", vcard)

def make_vcard( first_name, last_name, company, title, mobile, phone, address, email):
    if not math.isnan(mobile):
        if math.isnan(phone):
            return [
                'BEGIN:VCARD',
                'VERSION:2.1',
                f'N:{last_name};{first_name}',
                f'FN:{first_name} {last_name}',
                f'ORG:{company}',
                f'TITLE:{title}',
                f'EMAIL;PREF;INTERNET:{email}',
                f'TEL;WORK;VOICE:{int(mobile)}',
                f'ADR;WORK;PREF:;;{address}',
                f'REV:1',
                'END:VCARD'
            ]
        else:
            return [
                'BEGIN:VCARD',
                'VERSION:2.1',
                f'N:{last_name};{first_name}',
                f'FN:{first_name} {last_name}',
                f'ORG:{company}',
                f'TITLE:{title}',
                f'EMAIL;PREF;INTERNET:{email}',
                f'TEL;WORK;VOICE:{int(mobile)}',
                f'TEL;HOME;VOICE:{int(phone)}',
                f'ADR;WORK;PREF:;;{address}',
                f'REV:1',
                'END:VCARD'
            ]

def write_vcard(f, vcard):
    with open(f, 'w') as f:
        f.writelines([l + '\n' for l in vcard])

if __name__ == "__main__":
    main()