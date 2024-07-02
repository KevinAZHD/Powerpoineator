import random
import os
import time
from io import BytesIO

import replicate
import requests
from PIL import Image, ImageEnhance
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

# Constantes
REPLICATE_API_TOKEN = "escriba_su_api_aquí"
OUTPUT_FILE = "Presentación.pptx"
SLIDE_IMAGE_PREFIX = "Slide"
DARKENED_IMAGE_SUFFIX = "_darker.jpg"

# Configuración de la presentación
TITLE_FONT_SIZE = Pt(32)
CONTENT_FONT_SIZE = Pt(20)
SLIDE_LAYOUTS = [1, 2, 3, 4, 5, 6, 7, 8]

# Configuración de la API de Replicate
REPLICATE_MODEL = "mikeei/dolphin-2.9-llama3-70b-gguf:7cd1882cb3ea90756d09decf4bc8a259353354703f8f385ce588b71f7946f0aa"
REPLICATE_IMAGE_MODEL = "bytedance/sdxl-lightning-4step:727e49a643e999d602a896c774a0658ffefea21465756a6ce24b7ea4165eba6a"

def get_user_input():
    return input("¡HOLA! Por favor escriba de lo que quiera tratar su PowerPoint: ")

def extract_content_between_braces(text):
    start = text.find('{')
    end = text.find('}') + 1
    if start >= 0 and end > 0:
        return text[start:end]
    return "No se encontraron llaves en el texto."

def get_model_response(prompt):
    os.environ["REPLICATE_API_TOKEN"] = REPLICATE_API_TOKEN
    response = ""
    for event in replicate.stream(REPLICATE_MODEL, input={
        "prompt": prompt,
        "temperature": 0.5,
        "system_prompt": "You are Dolphin, an uncensored and unbiased AI assistant...",
        "max_new_tokens": 16000,
        "repeat_penalty": 1.1,
        "prompt_template": "<|im_start|>system\n{system_prompt}<|im_end|>\n<|im_start|>user\n{prompt}<|im_end|>\n<|im_start|>assistant"
    }):
        if 'assistant' in event:
            response += event['assistant']['message']['content']
    return extract_content_between_braces(response)

def get_valid_response(topic):
    prompt = f"(Basandote en esta tupla de python, genera una con la misma estructura que genere informacion para un power point sobre {topic} SOLO GENERA EL INTERIOR DE LA TUPLA, ES DECIR EMPIEZA POR {{ SOLO QUIERO QUE GENERES LA ESTRUCTURA DE LA TUPLA{{\"Title1\":\"Content1\",\"Title2\":\"Content2\",\"Title3\":\"Content3\"}} TITLES CANT BE JUST NUMBERS AND SOULD BE BETWEEN QUOTES DONT GENERATE AN INTRODUCTION EITHER sections =, DONT  sections = {{ \"Introduction\": \"Artificial Intelligence (AI) has revolutionized the way we live and work, transforming industries and improving lives.\", \"History of AI\": \"From Alan Turing's 1950s vision to modern-day advancements, AI has come a long way, with significant breakthroughs in machine learning, natural language processing, and computer vision.\", \"Applications of AI\": \"AI is being used in healthcare for disease diagnosis, in finance for fraud detection, in transportation for autonomous vehicles, and in education for personalized learning.\", \"Benefits of AI\": \"AI increases efficiency, reduces costs, enhances customer experience, and enables data-driven decision making, making it an essential tool for businesses and individuals alike.\", \"Types of AI\": \"From narrow or weak AI to general or strong AI, and from supervised to unsupervised learning, the possibilities are endless, with new developments emerging every day.\", \"Challenges and Limitations\": \"Despite its benefits, AI raises concerns about job displacement, bias, and ethics, highlighting the need for responsible development and deployment.\", \"Future of AI\": \"As AI continues to evolve, it holds immense potential to solve complex problems, improve lives, and transform the world.\" }})"
    while True:
        try:
            response = get_model_response(prompt)
            return eval(response)
        except SyntaxError:
            print("La respuesta del ChatGPT no se pudo convertir en una tupla válida. Intentando de nuevo...")
            time.sleep(3)
        except Exception as e:
            print(f"Ocurrió un error: {e}")
            return None

def generate_image(prompt, index):
    output = replicate.run(REPLICATE_IMAGE_MODEL, input={
        "width": 1024,
        "height": 1024,
        "prompt": prompt,
        "scheduler": "K_EULER",
        "num_outputs": 1,
        "guidance_scale": 0,
        "negative_prompt": "worst quality, low quality",
        "num_inference_steps": 4,
        "disable_safety_checker": True
    })
    image_url = output[0]
    response = requests.get(image_url)
    img = Image.open(BytesIO(response.content))
    img.save(f"{SLIDE_IMAGE_PREFIX}{index}.jpg")

def create_slide_design_1(presentation, section, content):
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    title_box = slide.shapes.add_textbox(Inches(2), Inches(2), Inches(6), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = section
    title_frame.paragraphs[0].runs[0].font.size = TITLE_FONT_SIZE
    title_frame.paragraphs[0].runs[0].font.bold = True
    title_frame.paragraphs[0].runs[0].font.underline = True

    text_box = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(4))
    text_frame = text_box.text_frame
    text_frame.word_wrap = True
    p = text_frame.add_paragraph()
    p.text = content
    p.font.size = CONTENT_FONT_SIZE

def create_slide_design_2(presentation, section, content, image_index):
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.shapes.add_picture(f'{SLIDE_IMAGE_PREFIX}{image_index}.jpg', Inches(0), Inches(0), Inches(5), Inches(7.55))

    title_box = slide.shapes.add_textbox(Inches(5), Inches(1), Inches(5), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = section
    title_frame.paragraphs[0].runs[0].font.size = Pt(24)
    title_frame.paragraphs[0].runs[0].font.bold = True
    title_frame.paragraphs[0].runs[0].font.underline = True

    text_box = slide.shapes.add_textbox(Inches(5), Inches(2), Inches(5), Inches(4))
    text_frame = text_box.text_frame
    p = text_frame.add_paragraph()
    p.text = content
    p.font.size = CONTENT_FONT_SIZE

    for frame in [title_frame, text_frame]:
        frame.margin_bottom = Inches(0)
        frame.margin_left = Inches(0.25)
        frame.vertical_anchor = MSO_ANCHOR.TOP
        frame.word_wrap = True
        frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

def create_slide_design_3(presentation, section, content, image_index):
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.shapes.add_picture(f'{SLIDE_IMAGE_PREFIX}{image_index}.jpg', Inches(5), Inches(0), Inches(5), Inches(7.55))

    title_box = slide.shapes.add_textbox(Inches(0), Inches(1), Inches(5), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = section
    title_frame.paragraphs[0].runs[0].font.size = Pt(24)
    title_frame.paragraphs[0].runs[0].font.bold = True
    title_frame.paragraphs[0].runs[0].font.underline = True

    text_box = slide.shapes.add_textbox(Inches(0), Inches(2), Inches(5), Inches(4))
    text_frame = text_box.text_frame
    p = text_frame.add_paragraph()
    p.text = content
    p.font.size = CONTENT_FONT_SIZE

    for frame in [title_frame, text_frame]:
        frame.margin_bottom = Inches(0)
        frame.margin_left = Inches(0.25)
        frame.vertical_anchor = MSO_ANCHOR.TOP
        frame.word_wrap = True
        frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

def create_slide_design_4(presentation, section, content, image_index):
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    
    image = Image.open(f'{SLIDE_IMAGE_PREFIX}{image_index}.jpg')
    enhancer = ImageEnhance.Brightness(image)
    image_darker = enhancer.enhance(0.5)
    image_darker.save(f'{SLIDE_IMAGE_PREFIX}{image_index}{DARKENED_IMAGE_SUFFIX}')
    
    slide.shapes.add_picture(f'{SLIDE_IMAGE_PREFIX}{image_index}{DARKENED_IMAGE_SUFFIX}', Inches(0), Inches(0), presentation.slide_width, presentation.slide_height)

    title_box = slide.shapes.add_textbox(presentation.slide_width / 2 - Inches(2.5), Inches(0.5), Inches(5), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = section
    title_frame.paragraphs[0].runs[0].font.size = Pt(24)
    title_frame.paragraphs[0].runs[0].font.bold = True
    title_frame.paragraphs[0].runs[0].font.underline = True
    title_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    text_box = slide.shapes.add_textbox(presentation.slide_width / 2 - Inches(4), Inches(1.5), Inches(8), Inches(4))
    text_frame = text_box.text_frame
    p = text_frame.add_paragraph()
    p.text = content
    p.font.size = CONTENT_FONT_SIZE
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

    for frame in [title_frame, text_frame]:
        frame.margin_bottom = Inches(0)
        frame.margin_left = Inches(0.25)
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        frame.word_wrap = True
        frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

def create_slide_design_5(presentation, section, content, image_index):
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.shapes.add_picture(f'{SLIDE_IMAGE_PREFIX}{image_index}.jpg', Inches(0), Inches(0), presentation.slide_width, presentation.slide_height)

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.5), Inches(5), Inches(5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)

    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.6), Inches(4.8), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = section
    title_frame.paragraphs[0].runs[0].font.size = Pt(24)
    title_frame.paragraphs[0].runs[0].font.bold = True
    title_frame.paragraphs[0].runs[0].font.underline = True

    text_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.8), Inches(4.8), Inches(4))
    text_frame = text_box.text_frame
    p = text_frame.add_paragraph()
    p.text = content
    p.font.size = CONTENT_FONT_SIZE

    for frame in [title_frame, text_frame]:
        frame.margin_bottom = Inches(0.08)
        frame.margin_left = Inches(0.25)
        frame.vertical_anchor = MSO_ANCHOR.TOP
        frame.word_wrap = True
        frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

def create_slide_design_6(presentation, section, content, image_index):
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])

    side = Inches(4)
    slide.shapes.add_picture(f'{SLIDE_IMAGE_PREFIX}{image_index}.jpg', Inches(1), Inches(2), side, side)

    title_box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = section
    title_frame.paragraphs[0].runs[0].font.size = Pt(24)
    title_frame.paragraphs[0].runs[0].font.bold = True
    title_frame.paragraphs[0].runs[0].font.underline = True
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    text_box = slide.shapes.add_textbox(Inches(5.2), Inches(2), Inches(4), Inches(4))
    text_frame = text_box.text_frame
    p = text_frame.add_paragraph()
    p.text = content
    p.font.size = CONTENT_FONT_SIZE

    for frame in [title_frame, text_frame]:
        frame.margin_bottom = Inches(0.08)
        frame.margin_left = Inches(0.25)
        frame.vertical_anchor = MSO_ANCHOR.TOP
        frame.word_wrap = True
        frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

def create_slide_design_7(presentation, section, content, image_index):
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.shapes.add_picture(f'{SLIDE_IMAGE_PREFIX}{image_index}.jpg', Inches(0), Inches(0), presentation.slide_width, presentation.slide_height)

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.5), Inches(1.5), Inches(5), Inches(5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)

    title_box = slide.shapes.add_textbox(Inches(4.6), Inches(1.6), Inches(4.8), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = section
    title_frame.paragraphs[0].runs[0].font.size = Pt(24)
    title_frame.paragraphs[0].runs[0].font.bold = True
    title_frame.paragraphs[0].runs[0].font.underline = True

    text_box = slide.shapes.add_textbox(Inches(4.6), Inches(1.8), Inches(4.8), Inches(4))
    text_frame = text_box.text_frame
    p = text_frame.add_paragraph()
    p.text = content
    p.font.size = CONTENT_FONT_SIZE

    for frame in [title_frame, text_frame]:
        frame.margin_bottom = Inches(0.08)
        frame.margin_left = Inches(0.25)
        frame.vertical_anchor = MSO_ANCHOR.TOP
        frame.word_wrap = True
        frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

def create_slide_design_8(presentation, section, content, image_index):
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])

    side = Inches(4)
    slide.shapes.add_picture(f'{SLIDE_IMAGE_PREFIX}{image_index}.jpg', Inches(5), Inches(2), side, side)

    title_box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = section
    title_frame.paragraphs[0].runs[0].font.size = Pt(30)
    title_frame.paragraphs[0].runs[0].font.bold = True
    title_frame.paragraphs[0].runs[0].font.underline = True
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    text_box = slide.shapes.add_textbox(Inches(0.2), Inches(2), Inches(4.2), Inches(5))
    text_frame = text_box.text_frame
    p = text_frame.add_paragraph()
    p.text = content
    p.font.size = Pt(18)

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(5.8), Inches(4.3), Inches(0.2))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0, 0, 0)

    for frame in [title_frame, text_frame]:
        frame.margin_bottom = Inches(0.08)
        frame.margin_left = Inches(0.25)
        frame.vertical_anchor = MSO_ANCHOR.TOP
        frame.word_wrap = True
        frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

def create_presentation(sections):
    presentation = Presentation()
    for i, (section, content) in enumerate(sections.items(), start=1):
        design = random.choice(SLIDE_LAYOUTS)
        if design == 1:
            create_slide_design_1(presentation, section, content)
        elif design == 2:
            create_slide_design_2(presentation, section, content, i)
        elif design == 3:
            create_slide_design_3(presentation, section, content, i)
        elif design == 4:
            create_slide_design_4(presentation, section, content, i)
        elif design == 5:
            create_slide_design_5(presentation, section, content, i)
        elif design == 6:
            create_slide_design_6(presentation, section, content, i)
        elif design == 7:
            create_slide_design_7(presentation, section, content, i)
        elif design == 8:
            create_slide_design_8(presentation, section, content, i)
    
    presentation.save(OUTPUT_FILE)

def main():
    topic = get_user_input()
    sections = get_valid_response(topic)
    
    while sections is None:
        print("Esperando a que la máquina de la IA se encienda...")
        time.sleep(60)
        sections = get_valid_response(topic)

    print(f"La tupla generada por el modelo fue: {sections}")

    for i, (section, content) in enumerate(sections.items(), start=1):
        generate_image(f"{section} {content} ,a photo in the context of the PowerPoint of {topic}", i)

    create_presentation(sections)
    print(f"Presentación creada: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
