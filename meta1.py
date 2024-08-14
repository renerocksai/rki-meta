# pip install python-docx tqdm pandas 
import os
from docx import Document
from tqdm import tqdm
import pandas as pd
from datetime import datetime, timedelta
import re


supported_extensions = ['.docx']


def get_files(directory):
    all_files = []
    for dirpath, _, filenames in os.walk(directory):
        for filename in filenames:
            basename = os.path.basename(filename)
            if basename.startswith('.') or basename.startswith('~'):
                continue
            extension = os.path.splitext(filename)[1].lower()
            if (extension not in supported_extensions):
                continue
            doc_path = os.path.join(dirpath, filename)
            all_files.append(doc_path)
    all_files.sort()
    return all_files


def process_file(filp):
    doc = Document(filp)
    return (os.path.basename(filp), doc.core_properties)


def process_folder(directory:str):
    all_files = get_files(directory)
    results = []
    print(f'Converting {len(all_files)} files in {directory}...')
    for f in tqdm(all_files):
        result = process_file(f)
        results.append(result)
    return results


if __name__ == '__main__':
    folder = sys.argv[1]
    results = process_folder(folder)
    with open('basis.csv', 'wt') as f:
        f.write('filename,author,created,last_modified_by,modified,revision\n')
        for (doc_name, p) in results:
            f.write(f'"{doc_name}","{p.author}",{p.created},"{p.last_modified_by}",{p.modified},{p.revision}\n')

    # now add extra convenience columns
    df = pd.read_csv('./out.csv')

    # convert strings to dates / times
    df['created'] = pd.to_datetime(df['created'])
    df['modified'] = pd.to_datetime(df['modified'])

    # mod_delta holds the time in days between creation and modification
    df['mod_delta'] = (df['modified'] - df['created']).dt.days

    # convert date string in filename to datetime
    def extract_date_from_filename(filename):
        match = re.search(r'\d{4}-\d{2}-\d{2}', filename)
        if match:
            return pd.to_datetime(match.group(0))
        return None
    df['date_from_filename'] = df['filename'].apply(extract_date_from_filename)

    # Ensure 'created' is in datetime format and only take the date part
    df['created_date'] = pd.to_datetime(df['created']).dt.date
    df['filename_creation_mismatch_days'] = (df['created_date'] - df['date_from_filename'].dt.date)

    # Compare the 'date_from_filename' with 'created_date' and create a boolean column
    df['date_mismatch'] = df['filename_creation_mismatch_days'] > timedelta(days=1)

    # save the csv with the extra convenience columns
    df.to_csv('mod_delta.csv')
