import os
from striprtf.striprtf import rtf_to_text
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime
import re

def convert_rtf_to_txt(input_folder, output_folder):
    successful_files = []
    unsuccessful_files = []

    for filename in os.listdir(input_folder):
        if filename.endswith(".rtf"):
            rtf_file = os.path.join(input_folder, filename)
            txt_file = os.path.join(output_folder, os.path.splitext(filename)[0] + ".txt")
            try:
                with open(rtf_file, 'r') as rtf:
                    text = rtf_to_text(rtf.read())
                    with open(txt_file, 'w') as txt:
                        txt.write(text)
                successful_files.append(filename)
            except Exception as e:
                unsuccessful_files.append((filename, str(e)))

    print("Successfully parsed and saved " + str(len(successful_files)) + " RTF files.")
    print("Parsing failed for " + str(len(unsuccessful_files)) + " RTF files.")
    
    if len(unsuccessful_files) > 0:
        print("Unsuccessful files:")
        for file, error in unsuccessful_files:
            print(file, error)

def process_txt_files(input_folder):
    data = []
    
    for filename in os.listdir(input_folder):
        if filename.endswith(".txt"):
            txt_file = os.path.join(input_folder, filename)
            with open(txt_file, 'r') as txt:
                soup = BeautifulSoup(txt, 'html.parser')
                
                lines = soup.get_text().split('\n')
                
                # Extract / transform date, newspaper, title according to their position
                date_str = lines[5].strip()
                date = datetime.strptime(date_str, '%B %d, %Y %A').date()
                state = ""
                newspaper = lines[4].strip()
                author = ""
                title = lines[3].strip()
                length = ""
                body = ""
                
                capture_body = False
                for line in lines:
                    # Strip everything after the name in the author line (usually comma or semicolon-delimited)
                    if line.startswith("Byline:"):
                        author_line = line.split("Byline:")[1].strip()
                        match = re.search(r'([^,;]+)', author_line)
                        if match:
                            author = match.group(0).strip()
                    elif line.startswith("Length:"):
                        length = line.split("Length:")[1].strip().split()[0]
                    # Extract the body which is stores in multiple lines (ending with certain marker words)
                    elif line.startswith("Body"):
                        capture_body = True
                        continue
                    elif line.startswith("Notes") or line.startswith("Classification") or line.lower().startswith(author.lower()):
                        break
                    elif capture_body:
                        body += line.strip() + " "
                
                data.append({
                    'Date': date,
                    'State': 'MN',
                    'Newspaper': newspaper,
                    'Author': author,
                    'Title': title,
                    'Length': length,
                    'Body': body,
                    'Energy Type': 'wind',
                })
    
    df = pd.DataFrame(data)
    return df

input_folder = R"C:\Users\lechl\OneDrive - TUM\Hiwi\Jeana\Local US Newspapers\Minneapolis Star Tribune\Articles"
output_folder = R"C:\Users\lechl\OneDrive - TUM\Hiwi\Jeana\Local US Newspapers\Minneapolis Star Tribune\TXT"

convert_rtf_to_txt(input_folder, output_folder)
df = process_txt_files(output_folder)
print(df.head(10))

output_excel_folder = R"C:\Users\lechl\OneDrive - TUM\Hiwi\Jeana\Local US Newspapers\Minneapolis Star Tribune\dataframe.xlsx"
df.to_excel(output_excel_folder, index=False)
print("Data saved successfully to", output_excel_folder)

output_pkl_folder = R"C:\Users\lechl\OneDrive - TUM\Hiwi\Jeana\Local US Newspapers\Minneapolis Star Tribune\dataframe.pkl"
df.to_pickle(output_pkl_folder)
print("Data saved successfully to", output_pkl_folder)

# Testing pickles
#df2 = pd.read_pickle(output_pkl_folder)
#print(df2.head(10))


    