import utils
import configparser


class Scrapper():
	def __init__(self):
		config = configparser.ConfigParser()
		config.read('config.ini')
		config_params = config['Gmail']

		inbox = input("Name of the mailbox from which to extract (default = INBOX): ") or 'INBOX'
		search_criteria = input("Criteria to use on search (default = ALL): ") or 'ALL'
		df_emails = utils.get_emails(config_params, inbox, search_criteria)

		output_filename = input("Output filename (default = output.xlsx):") or 'output.xlsx'
		utils.save_emails_to_file(df_emails, output_filename)


if __name__ == "__main__":
	run = Scrapper()