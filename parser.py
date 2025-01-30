import stanza
from docx import Document
from docx.shared import Pt
import os

# Function to process the text
def process_text(filename):
    # Read the content of the uploaded file
    with open(filename, 'r', encoding='utf-8') as file:
        text = file.read()

    # Download English model if not already downloaded
    stanza.download('en')

    # Initialize Stanza pipeline for English
    nlp = stanza.Pipeline('en')

    # Process the text with Stanza
    doc = nlp(text)

    # Create a new Word document
    docx_filename = 'parsed_data.docx'
    document = Document()

    # Add a title
    document.add_heading('Parsed Data', level=1)

    # In CoNLL format
    conll_format = ""
    for sentence in doc.sentences:
        for word in sentence.words:
            conll_format += "{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n".format(
                word.id,       # ID
                word.text,     # Text
                word.lemma,    # Lemma
                word.upos,     # UPOS
                word.xpos,     # XPOS
                word.head,     # Head
                0,   # Deprel
                word.start_char,  # Start Char
                word.end_char,    # End Char
                abs(word.id - word.head)  # Dependency Length
            )
        conll_format += "\n"

    # Save the CoNLL format text to a file
    conll_filename = 'parsed_data.conll'
    with open(conll_filename, 'w', encoding='utf-8') as file:
        file.write(conll_format)

    print(f"CoNLL format data saved to {conll_filename}")
    print(f"Word document saved to {docx_filename}")

# Main function to execute the script
if __name__ == "__main__":
    # Replace 'your_directory_path' with the path to your directory containing text files
    directory_path = 'C:\\Users\\govil\\\OneDrive\\Documents\\cgs_project\\SUD'

    # Check if the directory exists
    if os.path.exists(directory_path):
        # Loop through all files in the directory
        for filename in os.listdir(directory_path):
            # Process only .txt files
            if filename.endswith('.txt'):
                input_filename = os.path.join(directory_path, filename)
                process_text(input_filename)
    else:
        print(f"Directory {directory_path} not found.")