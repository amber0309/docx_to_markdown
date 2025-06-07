import os
import re

import torch
from transformers import Qwen2_5_VLForConditionalGeneration, AutoProcessor
from qwen_vl_utils import process_vision_info


def replace_image_placeholders(md_content, replacements):
    # 检查数量是否匹配
    placeholder_count = md_content.count("{{NONE}}")
    if placeholder_count != len(replacements):
        raise ValueError(
            f"占位符数量({placeholder_count})与替换列表长度({len(replacements)})不匹配"
        )

    # 逐个替换占位符
    for replacement in replacements:
        md_content = md_content.replace("{{NONE}}", replacement, 1)

    return md_content


def get_sorted_images(directory):
    PATTERN = re.compile(r"^img_(\d+)\.(?:png|jpg|jpeg)$", re.IGNORECASE)
    entries = os.listdir(directory)
    imgs = []
    for name in entries:
        full = os.path.join(directory, name)
        if not os.path.isfile(full):
            continue
        m = PATTERN.match(name)
        if not m:
            continue
        index = int(m.group(1))
        imgs.append((index, full))
    # 按 index 排序，然后只取路径
    imgs.sort(key=lambda x: x[0])
    return [path for _, path in imgs]


def get_img_info(path_imgs, model_name="Qwen/Qwen2.5-VL-7B-Instruct"):
    model = Qwen2_5_VLForConditionalGeneration.from_pretrained(
            model_name, 
            torch_dtype=torch.bfloat16, 
            device_map="balanced_low_0",
            attn_implementation="flash_attention_2"
            )
    processor = AutoProcessor.from_pretrained(model_name)

    imgs_lst = get_sorted_images(path_imgs)

    img_info_lst = []
    for _img in imgs_lst:
        messages = [
        {
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "image": _img,
                },
                {"type": "text", "text": "请用中文详细描述这张图片中各部分的含义，及不同部分之间的关系，并将描述整合在一个段落中"},
            ],
        }
        ]

        # Preparation for inference
        text = processor.apply_chat_template(
            messages, tokenize=False, add_generation_prompt=True
        )
        image_inputs, video_inputs = process_vision_info(messages)
        inputs = processor(
            text=[text],
            images=image_inputs,
            videos=video_inputs,
            padding=True,
            return_tensors="pt",
        )
        inputs = inputs.to(model.device)

        # Inference: Generation of the output
        generated_ids = model.generate(**inputs, max_new_tokens=2048)
        generated_ids_trimmed = [
            out_ids[len(in_ids) :] for in_ids, out_ids in zip(inputs.input_ids, generated_ids)
        ]
        output_text = processor.batch_decode(
            generated_ids_trimmed, skip_special_tokens=True, clean_up_tokenization_spaces=False
        )

        img_info_lst.append(output_text[0])

    return img_info_lst


def add_img_info(path_md, path_imgs, model_name, path_output):
    # read the markdown file
    with open(path_md, "r", encoding="utf-8") as f:
        md_content = f.read()

    placeholder_count = md_content.count("{{NONE}}")
    imgs_lst = get_sorted_images(path_imgs)
    if placeholder_count != len(imgs_lst):
        raise ValueError(
            f"number of placeholder ({placeholder_count}) != number of descriptions ({len(imgs_lst)})"
        )

    # use VLM to generate image descriptions
    img_info_lst = get_img_info(path_imgs, model_name)

    # replace the placeholders in the markdown content
    new_content = replace_image_placeholders(md_content, img_info_lst)
    
    # write the new content to a new markdown file
    fname = os.path.splitext(os.path.basename(path_md))[0]
    if path_output is not None:
        os.makedirs(path_output, exist_ok=True)
        path_new_md = os.path.join(path_output, f"{fname}_img.md")
    else:
        path_new_md = f"{fname}_img.md"
    with open(path_new_md, "w", encoding="utf-8") as f:
        f.write(new_content)

    return new_content


if __name__ == "__main__":
    # import time
    # for i in range(4, 6):
    #     add_img_info(f'/scratch/ywxzml3j/user17/test/chapter_{i}_latex.md', 
    #                 f'/home/ywxzml3j/ywxzml3juser17/imgs/img_ch{i}_latex',
    #                 model_name="Qwen/Qwen2.5-VL-72B-Instruct")
    #     time.sleep(10)

    add_img_info(f'/scratch/ywxzml3j/user17/test/rule_management.md', 
            f'/scratch/ywxzml3j/user17/test/images_rule_management',
            model_name="Qwen/Qwen2.5-VL-72B-Instruct")

