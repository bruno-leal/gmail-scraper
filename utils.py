# ------------------------------------------------------------------
# Contains functions to access a Gmail mailbox, extract e-mails,
# storing them in a pandas dataframe and saving it to a xlsx file.
#
# (C) 2022 Bruno Leal, Sertã, Portugal
# ------------------------------------------------------------------


# Import libraries
import imaplib, email # to access and handle the e-mails
import pandas as pd # to format and save the e-mails data
from datetime import datetime # to handle dates


def get_email_subject(email_obj):
	"""
	Gets subject from e-mail and decodes it's text (considering its specific encoding, if given).

	:param email_obj: object containing e-mail data, as returned by email library (email.message_from_bytes)
	"""
	subject_list = email.header.decode_header(email_obj['subject'])

	temp_subject_list = []
	for subject in subject_list:
		if subject[1]:
			subject = (subject[0].decode(subject[1]))
		elif type(subject[0]) == bytes:
			subject = subject[0].decode('utf-8')
		else:
			subject = subject[0]
		temp_subject_list.append(subject)

	subject = ''.join(temp_subject_list)

	return subject


def get_emails(login_params, mailbox='INBOX', search_criteria='ALL'):
	"""
	Gets all e-mails from specified inbox, filtering by given criteria.

	:param login_params: parameters used to connect to Gmail [should contain 'user', 'password' and 'imap_url']
	:param mailbox: inbox / folder / label from which to obtain the e-mails [default is main / root inbox]
	:param search_criteria: criteria to apply when filtering the e-mails [default = all e-mails] [check https://docs.python.org/3/library/imaplib.html#imaplib.IMAP4.search and https://pypi.org/project/imap-tools/#search-criteria to get info on how to search]
	:return: panda dataframe containing 'date', 'from', 'to', 'subject' and 'body' of each e-mail
	"""
	connection = imaplib.IMAP4_SSL(login_params['imap_url'])

	print('Trying to connect...')
	connection_result = connection.login(login_params['user'], login_params['password'])
	print('Connection result:', connection_result)

	print('Accessing inbox/folder...')
	selection_result = connection.select(mailbox)
	print('Accessing result:', selection_result)

	print('Searching for e-mails corresponding to the given conditions...')
	search_result, email_ids = connection.search(None, search_criteria)
	print('Search result:', search_result, email_ids)

	email_ids_list = email_ids[0].split()

	emails_df = pd.DataFrame(columns=('date', 'from', 'to', 'subject', 'body'))
	
	print('Extracting e-mail data...')
	for id in email_ids_list:
		typ, email_data = connection.fetch(id, '(RFC822)')

		print('\t___________________________E-mail id: {}___________________________'.format(id))

		for response_part in email_data:
			if type(response_part) is tuple:
				msg_obj = email.message_from_bytes(response_part[1])

				email_date = msg_obj['date']
				print("\tDate:", email_date)

				email_from = msg_obj['from']
				print("\tFrom:", email_from)

				email_to = msg_obj['to']
				print("\tTo:", email_to)

				email_subject = get_email_subject(msg_obj)
				print("\tSubject:", email_subject)

				for part in msg_obj.walk():
					if part.get_content_type() == 'text/plain':
						email_body = part.get_payload(decode=True)
						email_body = email_body.decode('ISO-8859-1')
						print('\tBody (first 50 characters):', email_body[0:50])

				emails_df = emails_df.append({ 'date': email_date, 'from': email_from, 'to': email_to, 'subject': email_subject, 'body': email_body}, ignore_index=True)

	emails_df["date"] = pd.to_datetime(emails_df["date"])

	return emails_df.sort_values(by='date')


def save_emails_to_file(df_emails, output_filename='output.xlsx'):
	"""
	Saves pandas dataframe containing e-mails data to xlsx file.

	:param df_emails: pandas dataframe contanint e-mails data
	"""
	# Since, "Excel does not support datetimes with timezones.", we need to strftime into string excluding time zone details and then convert back do datetime.
	df_emails['date'] = df_emails['date'].apply(lambda d: datetime.strftime(d,"%Y-%m-%d %H:%M:%S")) 
	df_emails['date'] = pd.to_datetime(df_emails['date'])

	df_emails.to_excel(output_filename, index=False)