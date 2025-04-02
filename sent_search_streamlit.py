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

def write_matches_to_log(matches, logger = False, log_file = "xml_parsed_log.txt"):
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
    
    # Remove whitespace from the top of the page and sidebar
    # the custom CSS lives here:
    hide_default_format = """
    <style>
        .reportview-container .main footer {visibility: hidden;}    
        #MainMenu, header, footer {visibility: hidden;}
        div.stActionButton{visibility: hidden;}
        [class="stAppDeployButton"] {
            display: none;
        }
    </style>

    """
    st.set_page_config(layout="wide",page_icon=":paperclip:",page_title="The Word Finder")
    # inject the CSS
    st.markdown(hide_default_format, unsafe_allow_html=True)
    col1, col2 = st.columns([1,8],gap = "small",vertical_alignment="bottom")
    with col1:
        st.image("clip.png")
    with col2:
        st.title("Welcome to the Word Finder!")

    st.markdown(':grey[This tool allows you to upload a document (must be .docx file) and a list of words or phrases (must be [.txt file](#txt-create) in a [certain format](#txt-format)). The document will be searched to find matches & similar matches to the words or phrases in your uploaded list of words. The results will be saved to a csv file that you can download. More information on the generated .csv matches file is detailed [below](#csv-file-structure).]')

    st.divider()

    st.header("Upload the Document You Want :violet[Searched]")
    uploaded_docx = st.file_uploader(":grey[Must choose a DOCX file]", type=["docx"])
    st.header("Upload the List of Words or Phrases You Want to :orange[Search For]")
    uploaded_txt = st.file_uploader(":grey[Must choose a TXT file. Make sure your TXT file word list is structured correctly. See [Word List Structure Rules](#txt-format) below]", type=["txt"])
    # Streamlit app to display instructions
    
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

            st.write('processing matches... (this may take a few minutes)')

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
                file_name="wordfinder_matches.csv",
                mime="text/csv",
                )
            
            # Clean up the temporary files
            os.remove(word_docx)
            os.remove(word_list_docx)
            os.remove(csv_file.name)
    
    st.divider()

    st.header("FAQ")

    with st.expander("How do I structure my list of words or phrases in my uploaded .txt file?"):
        st.subheader("Word List Structure Rules", anchor='txt-format')
        # Add the instructions as text
        # Instructions for formatting the .txt file
        st.markdown("""
        Make sure that your .txt file has each word or phrase on a new line. Here is the format you should follow:""")

        st.code("""
        red
        flowers that bloom
        blue
        wilting flowers
        """)
        
        st.markdown("""
        _Note: The algorithm currently does not support searching for general terms. For example it will not search for chair or desk if you put furniture as a search term in your .txt file._
        """)      

    with st.expander("Don't know how to create a .txt file? Click here for instructions"):
        st.subheader("Creating a .txt file", anchor='txt-create')
        # Add the instructions as text
        st.write("""
        ### On Windows: Using Notepad

        1. **Open Notepad**:
        - Press `Windows + R` to open the Run dialog.
        - Type `notepad` and press Enter.

        2. **Write your text**:
        - Type the content you want in the `.txt` file.

        3. **Save the file**:
        - Click on `File` in the top-left corner.
        - Select `Save As`.
        - In the "Save as type" dropdown, make sure it says `Text Documents (*.txt)`.
        - Choose a location to save the file and give it a name (e.g., `myfile.txt`).
        - Click `Save`.
        """)

        st.write("""
        ### On Mac: Using TextEdit

        1. **Open TextEdit**:
        - Open Spotlight by pressing `Cmd + Space`.
        - Type `TextEdit` and press Enter to launch the application.

        2. **Write your text**:
        - Click on `File` in the top menu and select `New`.
        - In the new document, type your content.
        - Make sure the file format is set to `Plain Text` by going to `Format` in the top menu and selecting `Make Plain Text`.

        3. **Save the file**:
        - Once you're done, click on `File` and then select `Save...`.
        - In the "Save As" field, enter your desired file name (e.g., `myfile.txt`), it must end in `.txt`.
        - Choose a location to save the file and click `Save`.
        """)

    with st.expander("What is in the generated .csv file?"):
        st.subheader("CSV File Structure", anchor='csv-file-structure')
        st.write("""        
        ### Example of CSV Output

        Here’s an example of how a row might look in the CSV file:

        | sent_id | list_matchs          | found_words          | match_certainty | sentence                             | page_at_or_below |
        |---------|----------------------|----------------------|-----------------|--------------------------------------|-------------------|
        | 5       | red, blue  | red, rod, blue | 100            | Red rods are better than blue ones | 2                 |
        | 24       | wilting flowers          | wilting, flowers           |             | She is a wilting flower.     | 3                 |
        | 300       | flowers that bloom, red       | blowers, that, gloomy, red        |     100        | The red blowers are his that seem gloomy | 5             |

        ### CSV Output Structure
        The CSV file contains the following columns, which are used to track the matching process and its results. Here's what each header represents:

        1. **`sent_id`**:
        represents the **unique identifier** for each sentence or entry in the uploaded document. It is used to differentiate each sentence that is being processed and they are sequentially numbered according to the order in which they appear in the document.

        2. **`list_matchs`**:
        shows **the words/phrases in your uploaded word list** that have a match or similar match in the sentence.

        3. **`found_words`**:
        contains the **specific words or phrases** that were identified in the sentence **from your uploaded document** as a match or similar match to the words/phrases in your uploaded word list.

        4. **`match_certainty`**:
        represents the **certainty level** of the match. It indicates how confident the system is that the words or phrases identified are correct matches. The value ranges from 0-100, with higher values indicating greater certainty. A value of 100 means there is an exact match. Phrases do not recieve a match certainty score.

        5. **`sentence`**:
        shows the sentence from the text that was evaluated. It provides context for the identified matches and allows the user to search for the match sentence in the original document via keyboard shortcut `ctrl + f` or `cmd + f` in that docuemnt.

        6. **`page_at_or_below`**:
        indicates the **page number** at or beyond the sentence was found in the orginal word .docx file. The sentence locating is somewhat imprecise, but the sentence **will not** appear before the listed page number for that sentence. The page locating will tend to become less accurate the deeper into your document you scroll due to word formatting limitations from tables and non-text components in a word document.

        ### How to Interpret the CSV Output:
        - The **`sent_id`** allows you and the system to track unique sentences in your document.
        - The **`list_matchs`** gives a overview of the matched keywords or phrases from your uploaded word list.
        - Thea **`found_words`** lists the exact words that were found in the sentence for you to review.
        - The **`match_certainty`** gives a confidence level on how accurate the match is.
        - The **`sentence`** column provides context, so you can understand where the matched words were found in the text.
        - The **`page_at_or_below`** helps track where in a document or series of pages the sentence is located, if you need to edit the sentence.
        """)

# Run the main function
if __name__ == "__main__":
    main()
