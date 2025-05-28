import os
import tempfile
from io import BytesIO

from docx import Document
from docx.table import Table as _Table
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from PIL import Image
from transformers import Qwen2_5_VLForConditionalGeneration, AutoProcessor
from qwen_vl_utils import process_vision_info
import torch


class Docx2MdConverter:

    def __init__(self, path_input_file, path_output=None, vlm=None):
        self.path_input = path_input_file
        self.image_counter = 0

        if path_output is None:
            self.output_file = None
            self.path_images = './images'
        else:
            file_name = os.path.basename(os.path.realpath(path_input_file))
            self.output_file = os.path.join(path_output, file_name.split('.')[0] + '.md')
            self.path_images = os.path.join(path_output, 'images')
        os.makedirs(self.path_images, exist_ok=True)

        self.ns = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'v': 'urn:schemas-microsoft-com:vml',
        }

        if vlm is not None and torch.cuda.is_available():
            self.model, self.processor = self._get_vlm(vlm)
        else:
            self.model, self.processor = None, None

    def execute(self):
        doc = Document(self.path_input)
        md_lines = []

        # 1) main text
        for block in doc._element.body.iterchildren():
            self._process_block(block, doc, md_lines)
        # # 2) header and footer
        # for sec in doc.sections:
        #     for block in sec.header._element.iterchildren():
        #         self._process_block(block, doc, md_lines)
        #     for block in sec.footer._element.iterchildren():
        #         self._process_block(block, doc, md_lines)

        # construct Markdown string
        markdown_text = "\n".join(md_lines)
        if self.output_file is not None:
            with open(self.output_file, 'w', encoding='utf-8') as f:
                f.write(markdown_text)

        return markdown_text

    # ----- ----- ----- -----
    # ----- internal funcs
    # ----- ----- ----- -----
    def _get_vlm(self, model_name):
        model = Qwen2_5_VLForConditionalGeneration.from_pretrained(
                model_name, 
                torch_dtype='auto', 
                device_map="balanced_low_0"
                )
        processor = AutoProcessor.from_pretrained(model_name)
        return model, processor

    def _get_image_description(self, path_image):
        messages = [
        {
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "image": path_image,
                },
                {"type": "text", "text": "请用中文详细描述这张图片的内容."},
            ],
        }
        ]

        # Preparation for inference
        text = self.processor.apply_chat_template(
            messages, tokenize=False, add_generation_prompt=True
        )
        image_inputs, video_inputs = process_vision_info(messages)
        inputs = self.processor(
            text=[text],
            images=image_inputs,
            videos=video_inputs,
            padding=True,
            return_tensors="pt",
        )
        inputs = inputs.to(self.model.device)

        # Inference: Generation of the output
        generated_ids = self.model.generate(**inputs, max_new_tokens=2048)
        generated_ids_trimmed = [
            out_ids[len(in_ids) :] for in_ids, out_ids in zip(inputs.input_ids, generated_ids)
        ]
        output_text = self.processor.batch_decode(
            generated_ids_trimmed, skip_special_tokens=True, clean_up_tokenization_spaces=False
        )

        desc = output_text[0]
        # with open(f"image_{self.image_counter-1}.txt", "w") as text_file:
        #     text_file.write(desc)
        return desc
        
    def _process_block(self, block, doc, md_lines):
        """
        递归处理 document.xml / header / footer 中的一个 block (<w:p> 或 <w:tbl>)。
        """
        tag = block.tag.split('}')[1]

        # 段落
        if tag == 'p':
            para = Paragraph(block, doc)
            level = self._parse_heading_level(para.style.name)
            # Heading
            if level > 0:
                text = para.text.strip()
                if text:
                    md_lines.append(f"{'#' * level} {text}")
                    md_lines.append("")
                return
            # 正文
            items = self._extract_paragraph_items(para, doc)
            buf = ""
            for typ, content in items:
                if typ == 'text':
                    buf += content
                else:  # image
                    if buf.strip():
                        md_lines.append(buf)
                        buf = ""
                    md_lines.append(f"![{content}]()")
            if buf.strip():
                md_lines.append(buf)
            md_lines.append("")

        # 表格
        elif tag == 'tbl':
            tbl = _Table(block, doc)
            rows = []
            for row in tbl.rows:
                cells = []
                for cell in row.cells:
                    # 单元格内也按段落→run 扫描
                    cell_items = []
                    for para in cell.paragraphs:
                        cell_items.extend(self._extract_paragraph_items(para, doc))
                    # 拼成单元格文本
                    s = ""
                    for typ, content in cell_items:
                        if typ == 'text':
                            s += content.replace('\n', ' ')
                        else:
                            s += f"![{content}]()"
                    cells.append(s.strip())
                rows.append(cells)

            if rows:
                # Markdown 表头
                header = rows[0]
                md_lines.append('| ' + ' | '.join(header) + ' |')
                md_lines.append('| ' + ' | '.join(['---'] * len(header)) + ' |')
                for row in rows[1:]:
                    md_lines.append('| ' + ' | '.join(row) + ' |')
                md_lines.append("")

    def _extract_run_items(self, run, doc) -> list:
        """
        从一个 run 中按出现顺序提取 [('text', ...), ('image', desc), …]。
        支持 DrawingML (<a:blip>) 和 VML (<v:imagedata>)。
        """
        items = []
        el = run.element

        # DrawingML
        for blip in el.findall('.//a:blip', self.ns):
            rid = blip.get(qn('r:embed'))            
            if not rid or rid not in doc.part.related_parts:
                continue

            part = doc.part.related_parts[rid]
            blob = part.blob
            relpath = self._save_blob_as_png(blob, part.content_type)
            if self.model is not None:
                desc = self._get_image_description(os.path.join(self.path_images, os.path.basename(relpath)))
            else:
                desc = 'None'
            items.append(('image', (relpath, desc)))

        # VML
        for vimg in el.findall('.//v:imagedata', self.ns):
            rid = vimg.get(qn('r:id'))
            if not rid or rid not in doc.part.related_parts:
                continue

            part = doc.part.related_parts[rid]
            blob = part.blob
            relpath = self._save_blob_as_png(blob, part.content_type)
            if self.model is not None:
                desc = self._get_image_description(os.path.join(self.path_images, os.path.basename(relpath)))
            else:
                desc = 'None'
            items.append(('image', (relpath, desc)))

        # plain text
        txt = run.text or ""
        if txt.strip():
            items.append(('text', txt))
        return items

    # ----- image related
    def _save_blob_as_png(self, blob: bytes, content_type: str) -> str:
        """
        普通位图（png/jpg/...）直接用 PIL 打开并保存为 PNG。
        向量图（EMF/WMF）则调用 _convert_vector_to_png。
        返回相对路径。
        """
        # 向量格式
        if content_type in ("image/x-emf", "image/emf", "image/x-wmf", "image/wmf"):
            # EMF/WMF 转 PNG
            ext = "emf" if "emf" in content_type else "wmf"
            return self._convert_vector_to_png(blob, ext)

        fname = f"image_{self.image_counter}.png"
        dst = os.path.join(self.path_images, fname)
        img = Image.open(BytesIO(blob))
        # 如果有 alpha 或者 P 模式的透明度信息，就把它当做 mask 贴到白底上
        if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
            # 转到 RGBA，以便拿到 alpha 通道
            rgba = img.convert("RGBA")
            alpha = rgba.split()[-1]
            # 建一个白底 RGB
            bg = Image.new("RGB", rgba.size, (255,255,255))
            # paste 时用 alpha 作为 mask，透明处留下白底，不透明处贴原图
            bg.paste(rgba, mask=alpha)
            final = bg
        else:
            # 否则直接转 RGB
            final = img.convert("RGB")
        final.save(dst, format="PNG")

        self.image_counter += 1
        return os.path.join(os.path.basename(self.path_images), fname)

    def _convert_vector_to_png(self, blob: bytes, ext: str) -> str:
        """
        把 EMF/WMF blob 写入临时文件，再用 soffice 转成 PNG。
        返回转换后 PNG 的相对路径（相对于 Markdown）。
        """
        with tempfile.TemporaryDirectory() as td:
            # 1) write to vector file
            vec_path = os.path.join(td, f"img.{ext}")
            with open(vec_path, "wb") as f:
                f.write(blob)
            
            # 2) convert to png file
            png_src = os.path.join(td, "img.png")
            im = Image.open(vec_path)
            im.save(os.path.join(td, "img.png"))
            if not os.path.exists(png_src):
                raise RuntimeError("转 EMF→PNG 失败，没有生成 img.png")

            # 3) move to the output path
            fname = f"image_{self.image_counter}.png"
            dst = os.path.join(self.path_images, fname)
            os.replace(png_src, dst)
            self.image_counter += 1

            #  return path relative to the Markdown file
            return os.path.join(os.path.basename(self.path_images), fname)

    # ----- text and table
    def _parse_heading_level(self, style_name: str) -> int:
        """
        从样式名推断 Heading 级别 (“Heading 1”/“标题 2” → 1/2)，否则 0。
        """
        s = (style_name or "").strip().lower()
        if s.startswith("heading"):
            parts = s.split()
            if len(parts) >= 2 and parts[1].isdigit():
                return int(parts[1])
        if s.startswith("标题"):
            num = ''.join(filter(str.isdigit, s))
            if num.isdigit():
                return int(num)
        return 0

    def _extract_paragraph_items(self, para: Paragraph, doc) -> list:
        """
        对一个 Paragraph 的所有 run 做扫描，返回扁平化 items。
        """
        items = []
        for run in para.runs:
            items.extend(self._extract_run_items(run, doc))
        return items

