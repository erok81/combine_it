import pandas as pd
import numpy as np
import glob
import sys
import os


print('''
 _____                    _      _                _  _    _ 
/  __ \                  | |    (_)              (_)| |  | |
| /  \/  ___   _ __ ___  | |__   _  _ __    ___   _ | |_ | |
| |     / _ \ | '_ ` _ \ | '_ \ | || '_ \  / _ \ | || __|| |
| \__/\| (_) || | | | | || |_) || || | | ||  __/ | || |_ |_|
 \____/ \___/ |_| |_| |_||_.__/ |_||_| |_| \___| |_| \__|(_)
                                                            
 ''')
all_data = pd.DataFrame()

while True:
	file_path = input(r'Enter folder name: ')
	if os.path.exists(file_path):
		break
	else:
		file_path = input(r'Folder not valid. Enter valid folder name: ')

	file_name = input('Enter common file name with extension to merge: ')
	print('All files starting with ' + file_name + ' will be merged')
	new_file_name = input('Enter master file name to be created: ')
	print("You chose " + new_file_name)

	choice = input(f'''Are these choices correct? \n ***common file names: {file_name} \n ***new master file name: {new_file_name} \n y to combine sheets n to re-enter data or q to quit: ''')
	if choice.lower() == 'y':
		break
	elif choice.lower() == 'n':
		pass
	elif choice.lower() == 'q':
		print('Thanks for using Combine It. Have a nice day.')
		raise SystemExit
	else:
		print('I don\'t understand input. Try again.')

for f in glob.glob(os.path.join(file_path,file_name)):
	df = pd.read_excel(f, header=4, skipfooter=3)
	all_data = all_data.append(df,ignore_index=True, sort=False)

master_file = f'{new_file_name}.xlsx'
writer = pd.ExcelWriter(os.path.join(file_path,master_file), engine='xlsxwriter')
all_data.to_excel(writer, sheet_name='Sheet1')
writer.save()
print('data saved as ' + master_file + ' in location' + file_path)

