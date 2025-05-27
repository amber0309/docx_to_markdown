import os

from docx import Document
from docx.table import Table as _Table
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


class Docx2MdConverter:

    def __init__(self, path_input_file, path_output=None):
        self.path_input = path_input_file

        if path_output is None:
            self.output_file = None
        else:
            file_name = os.path.basename(os.path.realpath(path_input_file))
            self.output_file = os.path.join(path_output, file_name.split('.')[0] + '.md')

        self.ns = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'v': 'urn:schemas-microsoft-com:vml',
        }

    def execute(self):
        doc = Document(self.path_input)
        md_lines = []

        # 1) 正文
        for block in doc._element.body.iterchildren():
            self._process_block(block, doc, md_lines)
        # # 2) 页眉页脚（可选）
        # for sec in doc.sections:
        #     for block in sec.header._element.iterchildren():
        #         self._process_block(block, doc, md_lines)
        #     for block in sec.footer._element.iterchildren():
        #         self._process_block(block, doc, md_lines)

        # 写出 Markdown
        markdown_text = "\n".join(md_lines)
        if self.output_file is not None:
            with open(self.output_file, 'w', encoding='utf-8') as f:
                f.write("\n".join(md_lines))

        return markdown_text

    # ----- ----- ----- -----
    # ----- internal funcs
    # ----- ----- ----- -----
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
            if rid and rid in doc.part.related_parts:
                blob = doc.part.related_parts[rid].blob
                desc = (None, '【待解析图片】')  # get_image_description(blob)
                items.append(('image', desc))

        # VML
        for vimg in el.findall('.//v:imagedata', self.ns):
            rid = vimg.get(qn('r:id'))
            if rid and rid in doc.part.related_parts:
                blob = doc.part.related_parts[rid].blob
                desc = (None, '【待解析图片】')  # get_image_description(blob)
                items.append(('image', desc))

        # 纯文字
        txt = run.text or ""
        if txt.strip():
            items.append(('text', txt))
        return items

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

