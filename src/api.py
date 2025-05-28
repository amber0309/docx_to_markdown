from src.docx2md import Docx2MdConverter

def docx_to_markdown(file_docx, path_output=None, vlm=None):
    """Convert a docx file to Markdown
    
    Parameters
    ----------
    file_docx: str
        the path of input docx file
    path_output: str, optional
        - if None, save converted Markdown file in path_output
        - if not None, no file would be saved
    vlm: str, optional
        - if None, image info not processed
        - if not None, use specified model to generate image caption

    Returns
    ----------
    markdown_text: str
        a string of converted Markdown
    """
    converter = Docx2MdConverter(file_docx, 
                                 path_output=path_output,
                                 vlm=vlm)
    markdown_text = converter.execute()
    return markdown_text
