# Importing Libraries
import re
import enchant  # to load custom dictionary
from enchant.checker import SpellChecker  # for autocorrecting colum using defined custom dictionary
import glob  # for loading multiple Excel files without specifying the names
import pandas as pd  # For reading and modifying

path = 'data'  # directory where data can be found
files = glob.glob(path + '/*.xlsx')  # Specifying the file extension for the file type .xlsx, .csv etc

def start_correction:
    # Loop over all the files and load into a Pandas' data frame
    for file in files:
        temp_df = pd.read_excel(file, sheet_name='Individuals')
        sheet_2 = pd.read_excel(file, sheet_name='Extract')
        if len(temp_df.index) > 95:
            temp_df = temp_df.drop([95, 96])
        else:
            temp_df = temp_df
        # Specifying the column [TemporaryName in this case] to be autocorrected
        # TODO: make it possible to correct several columns at the same time.
        orders = temp_df['TemporaryName'].astype(str)
        d = enchant.request_pwl_dict('./orderlist.txt')  # loading custom dictionary to be used
        checker = SpellChecker(d)  # Specifying to SpellChecker which dictionary to use
        corrected_list = []  # Initializing final list of autocorrected words

        #  Looping through words in column of interested  and correcting them
        for order in orders:
            order = order.replace('-', '_')
            parts = re.split('_+', order)  # Splitting the order name into
            checker.set_text(parts[0])  # Spell Checking the name part of the 'Order'
            suggested = checker.suggest(parts[0])  # Correcting any spelling mistakes
            for suggestion in suggested:
                suggestion = suggestion + '_' + parts[1]  # Recombining the name and number for the individual
                corrected_list.append(suggestion)  # Adding the corrected names to a corrected list

        temp_df['TemporaryName'] = corrected_list  # Replacing errors in the 'TemporaryName' column with the
        # print(len(corrected_list),  ':', corrected_list)

        # Remove the ['.file extension'] (line 1) and add '_corrected' to the file name (line 2)
        new_file = file.split('.')
        new_file_name = (new_file[0] + '_corrected.' + new_file[1]).replace('data', 'corrected/batch_1')

        # Print out each file name, and it's corresponding
        print(f'{new_file_name}: {len(corrected_list)} : {corrected_list}')
        temp_df.rename(columns={'Unnamed: 0': '', 'Unnamed: 1': '', 'Unnamed: 2': ''}, inplace=True)
        # uncomment the line below to save the file with the corrections in the 'corrected' folder:
        path = './' + new_file_name
        writer = pd.ExcelWriter(path, engine='xlsxwriter')
        temp_df.to_excel(writer, sheet_name='Individuals', index=False)
        sheet_2.to_excel(writer, sheet_name='Extract', index=False)
        writer.save()


