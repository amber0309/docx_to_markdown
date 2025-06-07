# Docx2Md

## Requirements

This package is tested on Python 3.10.16

Install basic dependencies using pip with:
```shell
pip -r requirements.txt
```

## Usage

### Basic usage

In basic usage, all text and tables in .docx file will be converted. Images will be saved in an directory in *path_output*. An example is given below:

```python
from src.api import docx_to_markdown
markdown_text = docx_to_markdown('mydoc.docx', './')
```

**NOTE**: Since Linux systems have poor support for emf/wmf format images, it is recommended to perform this step on a Windows system


### Image description generation

In this step, a vision-language model (VLM) is adopted to generate description of all images, which are then inserted in the markdown file to replace the image placeholder of their corresponding images. 

To perform this step, it is required to install the following dependencies:
```
torch==2.3.0
transformers==4.51.3
qwen_vl_utils==0.0.11
flash-attn==2.7.4.post1
```

An example is given below:

```python
from src.api import add_image_descriptions_to_markdown
add_image_descriptions_to_markdown('report.md',
                                    './img_report')
```
