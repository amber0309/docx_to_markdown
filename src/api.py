from src.docx2md import Docx2MdConverter
try:
    from src.img2text import add_img_info
except ImportError:
    print(
        '[Warning] One of packages in (torch, transformers, qwen_vl_utils) missing. Images info generation not available')


def docx_to_markdown(
    file_docx: str, 
    path_output: str = None, 
    vlm: str = None
    ) -> str:
    """Convert a docx file to Markdown
    
    Parameters
    ----------
    file_docx: str
        the path of input docx file
    path_output: str, optional
        - if None, no file would be saved
        - if not None, save converted Markdown file in path_output
    vlm: str, optional
        - if None, images will be replaced by placeholder ![('img', 'None')]()
        - if not None, use specified model to generate image description

    Returns
    ----------
    markdown_text: str
        a string of converted Markdown

    Examples
    --------
    content = convert_docx_to_markdown("report.docx")
    # Returns Markdown without saving
    
    content = convert_docx_to_markdown(
        "report.docx", 
        output_dir="converted",
        vlm_model="Qwen/Qwen2.5-VL-7B-Instruct"
    )
    # Converts with image descriptions and saves to 'converted/report.md'
    """
    converter = Docx2MdConverter(file_docx, 
                                 path_output=path_output,
                                 vlm=vlm)
    markdown_text = converter.execute()
    return markdown_text


def add_image_descriptions_to_markdown(
    path_md: str,
    path_imgs: str,
    model_name: str = "Qwen/Qwen2.5-VL-7B-Instruct",
    path_output: str = None
    ) -> str:
    """
    Enhances a Markdown file by generating descriptions for embedded images using a 
    multimodal vision-language model (VLM) and inserts descriptions into the document.

    Steps:
    1. Parses the input Markdown file to identify all embedded images.
    2. For each local image, uses the specified VLM to generate an alt-text description.
    3. Inserts generated descriptions in the caption of their corresponding images in the Markdown content.
    4. Saves the enhanced Markdown to the output directory

    Parameters:
    -----------
    path_md : str
        Path to the source Markdown file (e.g., "/docs/blog.md").
    path_imgs : str
        Base directory for resolving relative image paths in Markdown 
        (e.g., "/static/images" for `![img](posts/hero.png)` → resolves to "/static/images/posts/hero.png").
    model_name : str
        Name of the vision-language model for description generation 
        (e.g., "llava", "blip2", "fuyu-8b").
    path_output : str
        Directory to save the enhanced Markdown file (preserves original filename).

    Returns:
    --------
    markdown_text : str
        The enhanced Markdown content with image descriptions added.

    Example:
    --------
    add_image_descriptions_to_markdown(
        path_md="content/post.md",
        path_imgs="content/images",
        model_name="Qwen/Qwen2.5-VL-7B-Instruct",
        path_output="enhanced_content"
    )
    """
    return add_img_info(path_md, path_imgs, model_name, path_output)
