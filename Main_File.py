import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE


# --- Helper Function to color the word "Google" ---
def add_colored_google_text(paragraph, text, font_name, font_size, is_bold, default_color_rgb):
    """
    Adds text to a paragraph, coloring the word 'Google' with its logo colors.
    """
    google_colors = {
        'G': RGBColor(0x42, 0x85, 0xF4),  # Blue
        'o': RGBColor(0xDB, 0x44, 0x37),  # Red
        'o2': RGBColor(0xF4, 0xB4, 0x00),  # Yellow
        'g': RGBColor(0x42, 0x85, 0xF4),  # Blue
        'l': RGBColor(0x0F, 0x9D, 0x58),  # Green
        'e': RGBColor(0xDB, 0x44, 0x37)  # Red
    }

    # Split text by the word "Google"
    parts = text.split('Google')
    for i, part in enumerate(parts):
        # Add the part of the text that is not "Google"
        if part:
            run = paragraph.add_run()
            run.text = part
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = is_bold
            run.font.color.rgb = default_color_rgb

        # Add the word "Google" with special coloring, if it's not the last part
        if i < len(parts) - 1:
            # Add 'G'
            run = paragraph.add_run()
            run.text = 'G'
            run.font.color.rgb = google_colors['G']
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = is_bold
            # Add first 'o'
            run = paragraph.add_run()
            run.text = 'o'
            run.font.color.rgb = google_colors['o']
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = is_bold
            # Add second 'o'
            run = paragraph.add_run()
            run.text = 'o'
            run.font.color.rgb = google_colors['o2']
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = is_bold
            # Add 'g'
            run = paragraph.add_run()
            run.text = 'g'
            run.font.color.rgb = google_colors['g']
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = is_bold
            # Add 'l'
            run = paragraph.add_run()
            run.text = 'l'
            run.font.color.rgb = google_colors['l']
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = is_bold
            # Add 'e'
            run = paragraph.add_run()
            run.text = 'e'
            run.font.color.rgb = google_colors['e']
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = is_bold


# --- Presentation Setup ---
prs = pptx.Presentation()
prs.slide_width = Inches(16)
prs.slide_height = Inches(9)
blank_slide_layout = prs.slide_layouts[6]  # Blank layout

# Define colors
BLACK_FONT = RGBColor(0x20, 0x21, 0x24)
WHITE_FONT = RGBColor(0xFF, 0xFF, 0xFF)
FONT_NAME = "Open Sans"

# --- Slide 1: Title Slide ---
slide1 = prs.slides.add_slide(blank_slide_layout)
slide1.background.fill.solid()
slide1.background.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)  # White background for the image to sit on
slide1.shapes.add_picture('title_slide.png', 0, 0, width=prs.slide_width)

# Add Title
title_box = slide1.shapes.add_textbox(Inches(0.5), Inches(0.5), prs.slide_width - Inches(1), Inches(1))
p_title = title_box.text_frame.paragraphs[0]
p_title.alignment = pptx.enum.text.PP_ALIGN.CENTER
add_colored_google_text(p_title, "Why Google Glass Failed", FONT_NAME, 44, True, BLACK_FONT)

# Add Subtitle
subtitle_box = slide1.shapes.add_textbox(Inches(0.5), Inches(1.5), prs.slide_width - Inches(1), Inches(1))
p_subtitle = subtitle_box.text_frame.paragraphs[0]
p_subtitle.alignment = pptx.enum.text.PP_ALIGN.CENTER
p_subtitle.text = "An analysis of the ambitious project's rise and fall"
p_subtitle.font.name = FONT_NAME
p_subtitle.font.size = Pt(24)
p_subtitle.font.color.rgb = BLACK_FONT

# Add Navigation Circles
circle_topics = [
    ("What Was Google Glass?", RGBColor(0x4A, 0x90, 0xE2)),  # Blue
    ("The Hype vs. The Reality", RGBColor(0x4A, 0x90, 0xE2)),  # Blue
    ("The Prohibitive Price Point", RGBColor(0xD0, 0x02, 0x1B)),  # Red
    ("Major Privacy Concerns", RGBColor(0xD0, 0x02, 0x1B)),  # Red
    ("Awkward and Unfashionable Design", RGBColor(0xF5, 0xA6, 0x23)),  # Yellow
    ("Lack of a Clear Purpose", RGBColor(0xF5, 0xA6, 0x23)),  # Yellow
    ("Social Awkwardness and the \"Glasshole\" Effect", RGBColor(0x7E, 0xD3, 0x21)),  # Green
    ("The Enterprise Edition: A Second Life?", RGBColor(0x7E, 0xD3, 0x21))  # Green
]

positions = [
    (2, 3), (4.5, 3), (7, 3), (9.5, 3),
    (2, 5.5), (4.5, 5.5), (7, 5.5), (9.5, 5.5)
]

for i, (topic, color) in enumerate(circle_topics):
    left, top = positions[i]
    shape = slide1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(left), Inches(top), Inches(1.5), Inches(1.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()  # No outline

    tf = shape.text_frame
    p = tf.paragraphs[0]
    p.text = topic
    p.font.name = FONT_NAME
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = WHITE_FONT
    p.alignment = pptx.enum.text.PP_ALIGN.CENTER
    tf.word_wrap = True

# --- Slide 2: What Was Google Glass? ---
slide2 = prs.slides.add_slide(blank_slide_layout)
slide2.shapes.add_picture('blue_bg.png', 0, 0, width=prs.slide_width, height=prs.slide_height)
title2 = slide2.shapes.add_textbox(Inches(1), Inches(0.5), Inches(14), Inches(1))
p2_title = title2.text_frame.paragraphs[0]
add_colored_google_text(p2_title, "What Was Google Glass?", FONT_NAME, 36, True, WHITE_FONT)

content2 = slide2.shapes.add_textbox(Inches(1), Inches(1.5), Inches(7), Inches(6))
tf2 = content2.text_frame
tf2.word_wrap = True
text_lines_2 = [
    "A wearable, voice-controlled Android device.",
    "It featured a small screen, a camera, and a bone-conduction speaker.",
    "It was designed to provide a hands-free, augmented reality experience.",
    "The \"Explorer Edition\" was released to developers and early adopters in 2013 for $1,500."
]
for line in text_lines_2:
    p = tf2.add_paragraph()
    p.text = line
    p.font.name = FONT_NAME
    p.font.size = Pt(22)
    p.font.color.rgb = WHITE_FONT
    p.level = 1

slide2.shapes.add_picture('slide2_components.png', Inches(8.5), Inches(1.75), width=Inches(6))
slide2.shapes.add_picture('slide2_diagram.png', Inches(8.5), Inches(5), width=Inches(6))

# --- Slide 3: The Hype vs. The Reality ---
slide3 = prs.slides.add_slide(blank_slide_layout)
slide3.shapes.add_picture('blue_bg.png', 0, 0, width=prs.slide_width, height=prs.slide_height)
title3 = slide3.shapes.add_textbox(Inches(1), Inches(0.5), Inches(14), Inches(1))
title3.text_frame.paragraphs[0].text = "The Hype vs. The Reality"
title3.text_frame.paragraphs[0].font.name = FONT_NAME
title3.text_frame.paragraphs[0].font.size = Pt(36)
title3.text_frame.paragraphs[0].font.bold = True
title3.text_frame.paragraphs[0].font.color.rgb = WHITE_FONT

content3 = slide3.shapes.add_textbox(Inches(1), Inches(1.5), Inches(7), Inches(6))
tf3 = content3.text_frame
tf3.word_wrap = True
p3_1 = tf3.paragraphs[0]
p3_1.text = "The Hype:"
p3_1.font.bold = True
p3_1.font.name = FONT_NAME
p3_1.font.size = Pt(24)
p3_1.font.color.rgb = WHITE_FONT

hype_lines = [
    "A revolutionary device that would change the way we interact with technology.",
    "Featured in high-fashion magazines and worn by celebrities.",
    "Promised a future of seamless augmented reality."
]
for line in hype_lines:
    p = tf3.add_paragraph()
    p.text = line
    p.font.name = FONT_NAME
    p.font.size = Pt(22)
    p.font.color.rgb = WHITE_FONT
    p.level = 1

p3_2 = tf3.add_paragraph()
p3_2.text = "\nThe Reality:"
p3_2.font.bold = True
p3_2.font.name = FONT_NAME
p3_2.font.size = Pt(24)
p3_2.font.color.rgb = WHITE_FONT

reality_lines = [
    "A clunky, unfinished prototype.",
    "Limited functionality and poor battery life.",
    "A host of technical and social problems."
]
for line in reality_lines:
    p = tf3.add_paragraph()
    p.text = line
    p.font.name = FONT_NAME
    p.font.size = Pt(22)
    p.font.color.rgb = WHITE_FONT
    p.level = 1

slide3.shapes.add_picture('slide3_fashion.png', Inches(8.5), Inches(1.75), width=Inches(6))
slide3.shapes.add_picture('slide3_social.jpg', Inches(8.5), Inches(5), width=Inches(6))

# --- Slide 4: The Prohibitive Price Point ---
slide4 = prs.slides.add_slide(blank_slide_layout)
slide4.shapes.add_picture('red_bg.png', 0, 0, width=prs.slide_width, height=prs.slide_height)
slide4.shapes.add_picture('slide4_price.png', 0, 0, height=prs.slide_height)

title4 = slide4.shapes.add_textbox(Inches(8), Inches(0.5), Inches(7.5), Inches(1))
title4.text_frame.paragraphs[0].text = "The Prohibitive Price Point"
title4.text_frame.paragraphs[0].font.name = FONT_NAME
title4.text_frame.paragraphs[0].font.size = Pt(36)
title4.text_frame.paragraphs[0].font.bold = True
title4.text_frame.paragraphs[0].font.color.rgb = WHITE_FONT

content4 = slide4.shapes.add_textbox(Inches(8), Inches(1.5), Inches(7.5), Inches(6))
tf4 = content4.text_frame
tf4.word_wrap = True
text_lines_4 = [
    "The Explorer Edition cost $1,500.",
    "This was far too expensive for the average consumer, especially for a first-generation device with limited functionality.",
    "The high price created an image of elitism and exclusivity, which alienated many potential users.",
    "The bill of materials was estimated to be around $80, which made the high price tag even more difficult to justify."
]
for line in text_lines_4:
    p = tf4.add_paragraph()
    p.text = line
    p.font.name = FONT_NAME
    p.font.size = Pt(22)
    p.font.color.rgb = WHITE_FONT
    p.level = 1

# --- Slide 5: Major Privacy Concerns ---
slide5 = prs.slides.add_slide(blank_slide_layout)
slide5.shapes.add_picture('red_bg.png', 0, 0, width=prs.slide_width, height=prs.slide_height)
slide5.shapes.add_picture('slide5_concerns.png', 0, 0, height=prs.slide_height)

title5 = slide5.shapes.add_textbox(Inches(8), Inches(0.5), Inches(7.5), Inches(1))
title5.text_frame.paragraphs[0].text = "Major Privacy Concerns"
title5.text_frame.paragraphs[0].font.name = FONT_NAME
title5.text_frame.paragraphs[0].font.size = Pt(36)
title5.text_frame.paragraphs[0].font.bold = True
title5.text_frame.paragraphs[0].font.color.rgb = WHITE_FONT

content5 = slide5.shapes.add_textbox(Inches(8), Inches(1.5), Inches(7.5), Inches(6))
tf5 = content5.text_frame
tf5.word_wrap = True
text_lines_5 = [
    "The built-in camera raised serious privacy questions.",
    "The ability to record video and take pictures discreetly made people uncomfortable.",
    "The term \"Glasshole\" was coined to describe users who were perceived as invading the privacy of others.",
    "Many businesses, such as bars and movie theaters, banned the use of Google Glass on their premises."
]
for line in text_lines_5:
    p = tf5.add_paragraph()
    add_colored_google_text(p, line, FONT_NAME, 22, False, WHITE_FONT)
    p.level = 1

# --- Slide 6: Awkward and Unfashionable Design ---
slide6 = prs.slides.add_slide(blank_slide_layout)
slide6.shapes.add_picture('yellow_bg.png', 0, 0, width=prs.slide_width, height=prs.slide_height)
title6 = slide6.shapes.add_textbox(Inches(1), Inches(0.5), Inches(14), Inches(1))
p6_title = title6.text_frame.paragraphs[0]
p6_title.text = "Awkward and Unfashionable Design"
p6_title.font.name = FONT_NAME
p6_title.font.size = Pt(36)
p6_title.font.bold = True
p6_title.font.color.rgb = WHITE_FONT

content6 = slide6.shapes.add_textbox(Inches(1), Inches(1.5), Inches(14), Inches(3))
tf6 = content6.text_frame
tf6.word_wrap = True
text_lines_6 = [
    "The device was bulky, lopsided, and generally considered to be unattractive.",
    "It was not a device that people felt comfortable wearing in public.",
    "The design screamed \"tech gadget\" rather than \"fashion accessory.\"",
    "Despite collaborations with fashion designers, the fundamental design remained a major turn-off for consumers."
]
for line in text_lines_6:
    p = tf6.add_paragraph()
    p.text = line
    p.font.name = FONT_NAME
    p.font.size = Pt(22)
    p.font.color.rgb = WHITE_FONT
    p.level = 1

slide6.shapes.add_picture('slide6_awkward.jpg', 0, Inches(4.5), width=prs.slide_width)

# --- Slide 7: Lack of a Clear Purpose ---
slide7 = prs.slides.add_slide(blank_slide_layout)
slide7.shapes.add_picture('yellow_bg.png', 0, 0, width=prs.slide_width, height=prs.slide_height)
title7 = slide7.shapes.add_textbox(Inches(1), Inches(0.5), Inches(14), Inches(1))
p7_title = title7.text_frame.paragraphs[0]
p7_title.text = "Lack of a Clear Purpose"
p7_title.font.name = FONT_NAME
p7_title.font.size = Pt(36)
p7_title.font.bold = True
p7_title.font.color.rgb = WHITE_FONT

content7 = slide7.shapes.add_textbox(Inches(1), Inches(1.5), Inches(14), Inches(3.5))
tf7 = content7.text_frame
tf7.word_wrap = True
text_lines_7 = [
    "Google Glass was a solution in search of a problem.",
    "There was no single, compelling use case that made the device a \"must-have.\"",
    "The functionality it offered could be easily replicated by a smartphone.",
    "Without a clear purpose, it was difficult for consumers to justify the high price and social awkwardness."
]
for line in text_lines_7:
    p = tf7.add_paragraph()
    add_colored_google_text(p, line, FONT_NAME, 22, False, WHITE_FONT)
    p.level = 1

slide7.shapes.add_picture('question_glass.png', Inches(6), Inches(5), height=Inches(3.5))

# --- Slide 8: Social Awkwardness ---
slide8 = prs.slides.add_slide(blank_slide_layout)
slide8.shapes.add_picture('green_bg.png', 0, 0, width=prs.slide_width, height=prs.slide_height)
title8 = slide8.shapes.add_textbox(Inches(1), Inches(0.5), Inches(14), Inches(1))
p8_title = title8.text_frame.paragraphs[0]
add_colored_google_text(p8_title, "Social Awkwardness and the \"Glasshole\" Effect", FONT_NAME, 36, True, WHITE_FONT)

content8 = slide8.shapes.add_textbox(Inches(1), Inches(1.5), Inches(14), Inches(3.5))
tf8 = content8.text_frame
tf8.word_wrap = True
text_lines_8 = [
    "Wearing Google Glass in public was a socially awkward experience.",
    "It created a barrier between the user and the people they were interacting with.",
    "The device was seen as a sign of social disengagement and a potential invasion of privacy.",
    "The negative social stigma was a major deterrent for potential users."
]
for line in text_lines_8:
    p = tf8.add_paragraph()
    add_colored_google_text(p, line, FONT_NAME, 22, False, WHITE_FONT)
    p.level = 1

slide8.shapes.add_picture('slide8_glasshole.jpg', Inches(4), Inches(5), width=Inches(8))

# --- Slide 9: The Enterprise Edition ---
slide9 = prs.slides.add_slide(blank_slide_layout)
slide9.shapes.add_picture('green_bg.png', 0, 0, width=prs.slide_width, height=prs.slide_height)
title9 = slide9.shapes.add_textbox(Inches(1), Inches(0.5), Inches(14), Inches(1))
p9_title = title9.text_frame.paragraphs[0]
add_colored_google_text(p9_title, "The Enterprise Edition: A Second Life?", FONT_NAME, 36, True, WHITE_FONT)

content9 = slide9.shapes.add_textbox(Inches(1), Inches(1.5), Inches(7), Inches(6))
tf9 = content9.text_frame
tf9.word_wrap = True
text_lines_9 = [
    "Google eventually discontinued the consumer version of Glass in 2015.",
    "However, the company pivoted to the enterprise market with the \"Glass Enterprise Edition.\"",
    "This version of the device has found success in industries like manufacturing, logistics, and healthcare.",
    "In these contexts, the hands-free functionality and augmented reality features provide real value.",
    "But even those efforts were ultimately discontinued due to the tool's low impact on daily life."
]
for line in text_lines_9:
    p = tf9.add_paragraph()
    add_colored_google_text(p, line, FONT_NAME, 22, False, WHITE_FONT)
    p.level = 1

slide9.shapes.add_picture('slide9_edition.jpg', Inches(8.5), Inches(1.75), width=Inches(6))
slide9.shapes.add_picture('slide9_person.png', Inches(8.5), Inches(5), width=Inches(6))

# --- Slide 10: Lessons Learned ---
slide10 = prs.slides.add_slide(blank_slide_layout)
slide10.shapes.add_picture('conc_slide.png', 0, 0, width=prs.slide_width, height=prs.slide_height)
title10 = slide10.shapes.add_textbox(Inches(1), Inches(0.5), Inches(14), Inches(1))
p10_title = title10.text_frame.paragraphs[0]
p10_title.text = "Lessons Learned"
p10_title.font.name = FONT_NAME
p10_title.font.size = Pt(44)
p10_title.font.bold = True
p10_title.font.color.rgb = BLACK_FONT

content10 = slide10.shapes.add_textbox(Inches(1.5), Inches(1.75), Inches(13), Inches(5))
tf10 = content10.text_frame
tf10.word_wrap = True
text_lines_10 = [
    ("Don't release a prototype as a consumer product:",
     " The Explorer Edition was not ready for the public, and the negative first impression was difficult to overcome."),
    ("Privacy is paramount:",
     " In the age of connected devices, privacy must be a primary consideration in product design."),
    ("Design and social acceptability matter:",
     " For wearable technology, fashion and social norms are just as important as functionality."),
    ("A clear value proposition is essential:",
     " A new product needs to solve a real problem or offer a compelling new experience to succeed.")
]
for bold_part, regular_part in text_lines_10:
    p = tf10.add_paragraph()
    p.font.name = FONT_NAME
    p.font.size = Pt(22)
    p.font.color.rgb = BLACK_FONT
    p.level = 1

    bold_run = p.add_run()
    bold_run.text = bold_part
    bold_run.font.bold = True

    regular_run = p.add_run()
    regular_run.text = regular_part

slide10.shapes.add_picture('logo.png', Inches(7), Inches(7), height=Inches(1.5))

# --- Save Presentation ---
prs.save("Google_Glass_Failure_Presentation.pptx")
print("Presentation 'Google_Glass_Failure_Presentation.pptx' created successfully.")