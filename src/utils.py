import os
import sys


def contains_chinese_characters(s):
    """
    Checks if a string contains any Chinese characters.
    """
    for char in s:
        # The following Unicode ranges cover most Chinese characters.
        if '\u4e00' <= char <= '\u9fff' or '\u3400' <= char <= '\u4dbf':
            return True
    return False


def extract_headings_via_word_automation(doc_path):
    """
    Extracts headings and their precise, rendered numbering by automating
    the Microsoft Word application.

    Args:
        doc_path (str): The absolute path to the .docx file.

    Returns:
        list: A list of strings for each heading found, with its number.
              Returns None if Word is not available or an error occurs.
    """
    word = None
    doc = None
    headings = []
    heading_cnt = 0

    if not sys.platform.startswith('win'):
        return  headings, heading_cnt

    # The path must be absolute for Word automation to work reliably
    abs_doc_path = os.path.abspath(doc_path)

    try:
        import win32com.client as win32
        print("--- Extracting headings using MS Word Automation ---")
        # Start the Word application
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False  # Run in the background

        # Open the document
        doc = word.Documents.Open(abs_doc_path)

        # Iterate through all paragraphs in the document
        for para in doc.Paragraphs:
            # Check if the paragraph's style is a heading
            style_name = para.Style.NameLocal
            if style_name.lower().startswith('标题'):
                # para.Range.ListFormat.ListString is the key property.
                # It holds the rendered number/bullet string (e.g., "1.1", "A.", "i.").
                # It is an empty string for non-list paragraphs.
                number_string = para.Range.ListFormat.ListString

                # The text of the paragraph
                text = para.Range.Text.strip()

                if contains_chinese_characters(text):  # Ensure there is text in the paragraph
                    # If there's a number, combine it with the text
                    if number_string:
                        full_heading = f"{number_string} {text}"
                    else:
                        # If Word provides no number string, just use the text
                        full_heading = text

                    headings.append(full_heading)

    except Exception as e:
        print(f"An error occurred during Word automation: {e}")
        print("Please ensure Microsoft Word is installed and pywin32 is working.")
        return None
    finally:
        # IMPORTANT: Always close the document and quit Word
        if doc:
            doc.Close(False)  # False means don't save changes
        if word:
            word.Quit()

    print(f"{len(headings)} Headings extracted by Microsoft Word.")
    return headings, heading_cnt