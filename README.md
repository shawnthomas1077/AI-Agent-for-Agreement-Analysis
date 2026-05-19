# AI-Agent-for-Agreement-Analysis
```python
from docx import Document 
```

* Imports the `Document` class from the `python-docx` library.
* Used to read and work with Microsoft Word (`.docx`) files.

---

```python
import pandas as pd
```

* Imports the `pandas` library.
* `pd` is an alias for pandas.
* Used for handling tabular data using DataFrames.

---

```python
import nltk
```

* Imports the Natural Language Toolkit (`nltk`) library.
* Used for text processing and NLP (Natural Language Processing).

---

```python
from nltk.tokenize import sent_tokenize, word_tokenize
```

* Imports two tokenizer functions:

  * `sent_tokenize()` → splits text into sentences.
  * `word_tokenize()` → splits text into words/tokens.

---

```python
from collections import Counter
```

* Imports `Counter` from Python’s collections module.
* Used to count frequency of words.

---

# Download NLTK Tokenizer Model

```python
nltk.download('punkt')
```

* Downloads the `punkt` tokenizer model.
* Required for sentence and word tokenization.
* Usually downloaded once on a machine.

---

# STEP 1: READ WORD FILE

```python
def read_word(file_path):
```

* Defines a function named `read_word`.
* `file_path` is the path of the Word document.

---

```python
    doc = Document(file_path)
```

* Opens the Word document using `python-docx`.

---

```python
    text = []
```

* Creates an empty list to store paragraph text.

---

```python
    for para in doc.paragraphs:
```

* Loops through all paragraphs in the Word document.

---

```python
        if para.text.strip():
```

* Checks if the paragraph is not empty.
* `strip()` removes extra spaces.

---

```python
            text.append(para.text.strip())
```

* Adds cleaned paragraph text to the `text` list.

---

```python
    return text
```

* Returns the list of paragraphs.

---

# STEP 2: ANALYZE TEXT

```python
def analyze_text(paragraphs):
```

* Defines a function to analyze text paragraphs.

---

```python
    data = []
```

* Creates an empty list to store analysis results.

---

```python
    for i, para in enumerate(paragraphs):
```

* Loops through each paragraph.
* `enumerate()` provides:

  * `i` → index number
  * `para` → paragraph text

---

# Sentence Tokenization

```python
        sentences = sent_tokenize(para)
```

* Splits the paragraph into sentences.

Example:

```python
"This is first. This is second."
```

becomes:

```python
["This is first.", "This is second."]
```

---

# Word Tokenization

```python
        words = word_tokenize(para)
```

* Splits paragraph into words/tokens.

Example:

```python
"Payment is delayed."
```

becomes:

```python
["Payment", "is", "delayed", "."]
```

---

# Count Words

```python
        word_count = len(words)
```

* Counts total words/tokens.

---

# Count Sentences

```python
        sentence_count = len(sentences)
```

* Counts total sentences.

---

# Keyword Extraction

```python
        keywords = [word.lower() for word in words if word.isalnum()]
```

This line:

1. Loops through each word.
2. Converts word to lowercase.
3. Keeps only alphanumeric words.

Example:

```python
["Payment", "Delayed", "."]
```

becomes:

```python
["payment", "delayed"]
```

---

# Find Most Common Keywords

```python
        top_keywords = Counter(keywords).most_common(5)
```

* Counts frequency of words.
* Gets top 5 most common keywords.

Example:

```python
[('payment', 3), ('risk', 2)]
```

---

```python
        top_keywords_str = ", ".join([kw[0] for kw in top_keywords])
```

* Extracts only keyword names.
* Joins them into a comma-separated string.

Example:

```python
"payment, risk"
```

---

# Classification Logic

```python
        category = "General"
```

* Default category is `"General"`.

---

```python
        if "risk" in para.lower():
```

* Checks if paragraph contains the word `"risk"`.

---

```python
            category = "Risk"
```

* Sets category to `"Risk"`.

---

```python
        elif "payment" in para.lower():
```

* Checks if paragraph contains `"payment"`.

---

```python
            category = "Financial"
```

* Sets category to `"Financial"`.

---

```python
        elif "termination" in para.lower():
```

* Checks if paragraph contains `"termination"`.

---

```python
            category = "Legal"
```

* Sets category to `"Legal"`.

---

# Store Analysis Data

```python
        data.append({
```

* Adds analysis result as a dictionary into `data`.

---

```python
            "Paragraph No": i + 1,
```

* Stores paragraph number.
* `i + 1` because indexing starts from 0.

---

```python
            "Text": para,
```

* Stores original paragraph text.

---

```python
            "Sentence Count": sentence_count,
```

* Stores number of sentences.

---

```python
            "Word Count": word_count,
```

* Stores number of words.

---

```python
            "Top Keywords": top_keywords_str,
```

* Stores top keywords.

---

```python
            "Category": category
```

* Stores assigned category.

---

```python
        })
```

* Ends dictionary.

---

# Convert to DataFrame

```python
    return pd.DataFrame(data)
```

* Converts list of dictionaries into a pandas DataFrame.
* Returns structured tabular data.

Example:

| Paragraph No | Category | Word Count |
| ------------ | -------- | ---------- |
| 1            | Risk     | 45         |

---

# STEP 3: EXPORT TO EXCEL

```python
def export_to_excel(df, output_file):
```

* Defines function to export data to Excel.

---

```python
    df.to_excel(output_file, index=False, engine='openpyxl')
```

* Saves DataFrame into Excel file.
* `index=False` removes extra index column.
* `openpyxl` is used as Excel engine.

---

# MAIN AGENT FUNCTION

```python
def ai_agent(word_file, excel_output):
```

* Main controller function.

---

```python
    paragraphs = read_word(word_file)
```

* Reads Word file and gets paragraphs.

---

```python
    analyzed_df = analyze_text(paragraphs)
```

* Analyzes paragraphs.

---

```python
    export_to_excel(analyzed_df, excel_output)
```

* Exports analysis to Excel.

---

```python
    print(f"✅ Analysis complete. File saved at: {excel_output}")
```

* Displays success message.

---

# PROGRAM ENTRY POINT

```python
if __name__ == "__main__":
```

* Checks if the script is being run directly.

---

```python
    ai_agent("input.docx", "output_analysis.xlsx")
```

* Runs the program.
* Input:

  * `input.docx`
* Output:

  * `output_analysis.xlsx`

---

# Overall Flow of the Program

```text
Word File (.docx)
        ↓
Read Paragraphs
        ↓
Analyze Text
   - Sentence Count
   - Word Count
   - Keywords
   - Category
        ↓
Create DataFrame
        ↓
Export to Excel
```

