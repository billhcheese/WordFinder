import streamlit as st
import xml.etree.ElementTree as ET
import zipfile
import os
import io
import tempfile
from tempfile import NamedTemporaryFile
from collections import defaultdict
from fuzzywuzzy import fuzz
import re
import pandas as pd
import numpy as np
import csv

#FUNCTIONS TO PROCESS THE WORD DOC AND XML------------------------------------
def unzip_word_document(docx_path, extract_to_folder):
    # Ensure the output folder exists
    if not os.path.exists(extract_to_folder):
        os.makedirs(extract_to_folder)
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to_folder)
        print(f"Word document unzipped successfully to {extract_to_folder}")
    except zipfile.BadZipFile:
        print(f"The file {docx_path} is not a valid ZIP archive.")
    except Exception as e:
        print(f"An error occurred: {e}")

# Function to unzip a DOCX file and process its content
def unzip_docx(docx_file):
    # Create a temporary directory
    temp_dir = tempfile.mkdtemp()

    # In case the temp_dir already exists for some reason, we remove and retry
    if os.path.exists(temp_dir):
        os.rmdir(temp_dir)
        temp_dir = tempfile.mkdtemp()
    
    # Unzip the .docx file (it's essentially a ZIP file)
    st.write("Unzipping the DOCX file...")
    with zipfile.ZipFile(docx_file, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    
    return temp_dir

def parse_xml(file_path):
    """Parse the XML file and return the root element."""
    tree = ET.parse(file_path)
    return tree.getroot()

def extract_matches(root):
    """Extract relevant matches from the XML and return them as a list."""
    combined_matches = []
    for elem in root.iter():
        if elem.tag in [
            r'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t', 
            r'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p', 
            r'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lastRenderedPageBreak'
        ]:
            combined_matches.append(elem)
    return combined_matches

def write_matches_to_log(matches, logger = True, log_file = "xml_parsed_log.txt"):
    """
    Write the matches to a log file, formatted accordingly.
    Bug: This function does not log page breaks if they happen in a table or other irregular word doc elements some times. 
        This is because the document.xml of the word document does not have a in line sequential page break tag on such elements like it does regularly.

    """
    xml_parsed = str()
    page_count = 1
    for match in matches:
        if match.tag == r'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t':
            xml_parsed += f'{match.text}'
        elif match.tag == r'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lastRenderedPageBreak':
            xml_parsed += f'[lastRenderedPageBreak{page_count}]\n'
            page_count += 1
        elif match.tag == r'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
            xml_parsed += (f'[newParagraph]\n')
    
    
    if logger == True:
        page_count = 1
        with open(log_file, 'w') as log:
            for match in matches:
                if match.tag == r'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t':
                    log.write(f'{match.text}')
                elif match.tag == r'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lastRenderedPageBreak':
                    log.write(f'\n------------[lastRenderedPageBreak{page_count}]------------------------------------------------------------\n\n')
                    page_count += 1
                elif match.tag == r'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
                    log.write(f'\n')

    return xml_parsed

def split_text_on_paragraphs(xml_parsed):
    """Splits the input text by [newParagraph] and returns the split text."""
    return re.split(r'(\[newParagraph\])', xml_parsed)

def extract_page_number(part):
    """Extracts page number from a given part if it contains a page break."""
    page_match = re.search(r'\[lastRenderedPageBreak(\d+)\]', part)
    if page_match:
        return int(page_match.group(1)) + 1
    return None

def clean_part(part):
    """Removes unwanted tags like [newParagraph],[lastRenderedPageBreak#] and newline characters."""
    part = re.sub(r'\[lastRenderedPageBreak\d+\]', '', part)  # Remove the page break tag
    part = re.sub(r'\[newParagraph\]', '', part)  # Remove [newParagraph] tag
    part = re.sub(r'\n', '', part)  # Remove newline characters
    return part

def process_sentences(part, page_number, sentence_list, sentence_id, current_sentence):
    """Processes the part into sentences and handles combining short sentences."""
    sentences = part.split('.')
    for sentence in sentences:
        sentence = sentence.strip()
        if sentence:
            if len(sentence.split()) < 5 and current_sentence:
                current_sentence += ' ' + sentence
            else:
                if current_sentence:
                    sentence_list.append({'sent_id': sentence_id, 'sentence': current_sentence.lower(), 'page': page_number, 'matches':[]})
                    #sentence_list.append({'sent_id': sentence_id, 'sentence': current_sentence.lower(), 'page': page_number, 'matches':[], 'found_words':[]})
                    sentence_id += 1
                current_sentence = sentence
    return sentence_list, sentence_id, current_sentence

def add_last_sentence(sentence_list, sentence_id, current_sentence, page_number):
    """Adds the last sentence to the sentence list if there is any."""
    if current_sentence:
        sentence_list.append({'sent_id': sentence_id, 'sentence': current_sentence, 'page': page_number, 'matches':[]})
        #sentence_list.append({'sent_id': sentence_id, 'sentence': current_sentence, 'page': page_number, 'matches':[], 'found_words':[]})
    return sentence_list

def sentence_convert(xml_parsed):
    """Main function to process the XML parsed text."""
    split_text = split_text_on_paragraphs(xml_parsed)
    sentence_list = []
    sentence_id = 1
    page_number = 1
    current_sentence = ""

    for part in split_text:
        page_number = extract_page_number(part) or page_number
        part = clean_part(part)
        
        if part.strip():
            sentence_list, sentence_id, current_sentence = process_sentences(part, page_number, sentence_list, sentence_id, current_sentence)
    
    sentence_list = add_last_sentence(sentence_list, sentence_id, current_sentence, page_number)
    return sentence_list

#FUNCTIONS TO LOAD THE WORD LIST-------------------------------------
def load_word_list(word_file):
    # Function to load the words from a given list
    with open(word_file, 'r') as file:
        return [line.strip() for line in file.readlines() if line.strip()]
        #return [line.strip().lower() for line in file.readlines()]

#FUNCTIONS TO CHECK THE SENTENCES AGAINST THE WORD LIST-------------------------------------
def tokenize_sent(sentence):
    # Tokenize the sentence into words
    sent_words = [sent.strip(".,:;()!?\'\"\\") for sent in sentence.split()]
    return sent_words

def tokenize_word(word_list):
    # Tokenize the word list into words
    '''
        {
        'word_orig': 'clean energy',
        'word_tokens': ['clean','energy'],
        'phrase_type': 'single_word'(or 'multi_word' or 'general'),
        }
    '''
    
    token_items = []
    for phrase in word_list:
        word_tokens = [word.strip(".,:;()!?\'\"\\") for word in phrase.split()]
        
        if len(word_tokens) == 1:
            phrase_type = 'single_word'
        else:
            phrase_type = 'multi_word'

        token_items.append({
            'word_orig': phrase,
            'word_tokens': word_tokens,
            'phrase_type': phrase_type
        })

    return token_items

def check_sentence(sentence_list,word_list): #to compare the sentence to the word list
    sensitivity = 75
    similarity_tracker = {}
    token_word_dict = tokenize_word(word_list)
    
    #process single word comparisons
    for sent_item in sentence_list:
        sentence = sent_item['sentence']
        sent_id = sent_item['sent_id']
        token_sent = tokenize_sent(sentence)
        similarity_tracker[sent_id] = {}
        for sent_word in token_sent:
            similarity_tracker[sent_id][sent_word] = {}
            for word_phrase in token_word_dict:
                for word in word_phrase['word_tokens']:
                    word_ratio = fuzz.ratio(sent_word, word)
                    similarity_tracker[sent_id][sent_word][word]=word_ratio
                    if word_phrase['phrase_type'] == 'single_word' and word_ratio >= sensitivity:
                        sent_item['matches'].append({
                            'match': word,
                            'ratio': word_ratio,
                            'found': sent_word
                            })

    # Iterate through the dictionary for multi-word phrases
    for sent_id, sent_dict in similarity_tracker.items():
        qualified_words = []
        
        #finds word in sentences that are over 75% matched to word(s) on the list
        for word, scores in sent_dict.items():
            # Check if either 'accessible' or 'activism' is over 75
            sent_sensitivity = []
            for list_word, score in scores.items():
                if score > sensitivity:
                    sent_sensitivity.append(list_word)
            if sent_sensitivity:
                qualified_words.append({'found_word':word, 'list_word':sent_sensitivity})
        
        # check if all multi-word phrases are found in words that are matched over 75% and add them to the sentence item matches
        for word_phrase in token_word_dict:
            if word_phrase['phrase_type'] == 'multi_word':
                found_words = [q_word.get('found_word', None) for q_word in qualified_words]
                if all(term in found_words for term in word_phrase['word_tokens']):
                    words_extract = set(word_phrase['word_tokens']) & set(found_words)
                    find_dictionary(sentence_list, 'sent_id', sent_id)['matches'].append({
                        'match': word_phrase['word_orig'],
                        'ratio': None,
                        'found': list(words_extract)
                        })

#utility functions
def max_ignore_none(data):
    filtered_data = [x for x in data if x is not None]
    return max(filtered_data) if filtered_data else None

def find_dictionary(list_of_dictionaries, key, value):
    for dictionary in list_of_dictionaries:
        if dictionary.get(key) == value:
            return dictionary
    return None # or raise an exception if no match is found

#FUNCTIONS TO EXPORT THE DATA-------------------------------------
# Helper function to concatenate lists and strings
def concat_lists_strings(series):
    # Flatten lists and join with commas
    return ', '.join(set((map(str, [item for sublist in series for item in (sublist if isinstance(sublist, list) else [sublist])]))))

# Function to process and collapse sentence list into a DataFrame
def collapse_sentence_data(sentence_list):
    # Create the DataFrame
    df = pd.json_normalize(sentence_list, 'matches', ['sent_id', 'sentence', 'page'], errors='ignore')

    # Grouping by 'sent_id' and applying the aggregation
    collapsed_df = df.groupby('sent_id').agg(
        list_matchs=('match', lambda x: concat_lists_strings(x)),  # Concatenate match strings
        found_words=('found', lambda x: concat_lists_strings(x)),  # Concatenate found strings
        match_certainty=('ratio', lambda x: max(filter(lambda y: y is not None, x), default=None)),  # Take max of ratio, ignore None
        sentence=('sentence', 'first'),  # Take the first sentence for each group
        page_at_or_below=('page', 'first')  # Take the first page for each group
    ).reset_index()

    return collapsed_df

# RUNNING THE MAIN FUNCTION--------------------------------------
def main():
    """Main function to parse the XML, extract matches, and write them to a log."""
    # Streamlit UI
    st.title("Upload DOCX and Download Generated Files")

    uploaded_docx = st.file_uploader("Choose a DOCX file", type=["docx"])
    uploaded_txt = st.file_uploader("Choose a TXT file", type=["txt"])
    
    if uploaded_docx is not None and uploaded_txt is not None:
        # Display the button to process files
        if st.button("Process Files"):
            # Save the DOCX and TXT files temporarily
            with NamedTemporaryFile(delete=False, mode="wb") as docx_tmp:
                docx_tmp.write(uploaded_docx.getvalue())
                word_docx = docx_tmp.name
            
            with NamedTemporaryFile(delete=False, mode="wb") as txt_tmp:
                txt_tmp.write(uploaded_txt.getvalue())
                word_list_docx = txt_tmp.name

            st.write("File successfully uploaded!")
            
            # Unzip the DOCX file
            temp_dir = unzip_docx(word_docx)

            # Process the unzipped DOCX contents (e.g., extract text from the document.xml file)
            document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')

            #unzip_word_document(word_docx)

            #xml_file = 'word/document.xml'
            
            # Parse XML and extract matches
            root = parse_xml(document_xml_path)
            matches = extract_matches(root)
            st.write('xml parsed')
            
            # Write the matches to a log file
            xml_parsed = write_matches_to_log(matches, logger = True)
            st.write('xml written to log')

            # Turn xml into sentence units
            sentence_list = sentence_convert(xml_parsed)
            st.write('sentence list created')
            
            # Load the word list
            word_list = load_word_list(word_list_docx)
            st.write('word list loaded')

            # Check the sentences for matches
            check_sentence(sentence_list, word_list)
            st.write('sentences checked')

            # Create the DataFrame
            collapsed_df = collapse_sentence_data(sentence_list)

            # Save the DataFrame to a CSV file
            csv_file = NamedTemporaryFile(delete=False, mode='w', suffix='.csv', newline='')
            collapsed_df.to_csv(csv_file.name, index=False)
            st.write('collapsed data saved to csv')

            # Display the collapsed DataFrame
            #print(sentence_list)
            #print(collapsed_df)

            # Clean up the temporary directory where the DOCX contents were extracted
            for root, dirs, files in os.walk(temp_dir, topdown=False):
                for name in files:
                    os.remove(os.path.join(root, name))
                for name in dirs:
                    os.rmdir(os.path.join(root, name))
            os.rmdir(temp_dir)

            with open(csv_file.name, 'r') as f:
                st.download_button(
                label="Download Generated Files",
                data=f,
                file_name="collapsed_data.csv",
                mime="text/csv",
                )
            
            # Clean up the temporary files
            os.remove(word_docx)
            os.remove(word_list_docx)
            os.remove(csv_file.name)

# Run the main function
if __name__ == "__main__":
    main()
