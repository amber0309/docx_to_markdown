# Docx2Md

## Requirements

This package is tested on Python 3.10.16

Install dependencies using pip with:
```shell
pip -r requirements.txt
```

## Usage

 parse a docx file into Markdown
```python
from src.api import docx_to_markdown
markdown_text = docx_to_markdown('mydoc.docx', './')
```
