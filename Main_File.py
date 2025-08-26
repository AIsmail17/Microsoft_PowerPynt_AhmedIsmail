# Libraries
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# Google colors using RGB
GOOGLE_BLUE = RGBColor(66, 133, 244)
GOOGLE_RED = RGBColor(234, 67, 53)
GOOGLE_YELLOW = RGBColor(251, 188, 5)
GOOGLE_GREEN = RGBColor(52, 168, 83)
GOOGLE_NEWS_BLACK = RGBColor(32, 33, 36)

# Google's font
FONT_NAME = "Open Sans"

# Presentation
prs = Presentation()
# Standard 16:9 widescreen format for slides
prs.slide_width = Inches(16)
prs.slide_height = Inches(9)

# Slide layout 6 for title (Blank Slide)
blank_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_slide_layout)

# The image 'S1.png' as BG

left = top = Inches(0)
pic = slide.shapes.add_picture('S1.png', left, top, width=prs.slide_width, height=prs.slide_height)
slide.shapes._spTree.remove(pic._element)
slide.shapes._spTree.insert(2, pic._element)

# The Title
# Dimensions and position for the title
left = Inches(2.0)
top = Inches(0.7)
width = Inches(12.0)
height = Inches(1.5)
txBox = slide.shapes.add_textbox(left, top, width, height)
tf = txBox.text_frame
tf.clear()
tf.margin_bottom = Inches(0)
tf.margin_left = Inches(0)
tf.margin_right = Inches(0)
tf.margin_top = Inches(0)
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER

# The title: "Why " in black because "Google" logo will have different colors
run = p.add_run()
run.text = 'Why '
run.font.name = FONT_NAME
run.font.size = Pt(70)
run.font.bold = True
run.font.color.rgb = GOOGLE_NEWS_BLACK

# "G" inBlue
run = p.add_run()
run.text = 'G'
run.font.name = FONT_NAME
run.font.size = Pt(70)
run.font.bold = False
run.font.color.rgb = GOOGLE_BLUE

# "o" in Red
run = p.add_run()
run.text = 'o'
run.font.name = FONT_NAME
run.font.size = Pt(70)
run.font.bold = False
run.font.color.rgb = GOOGLE_RED

# "o" in Yellow
run = p.add_run()
run.text = 'o'
run.font.name = FONT_NAME
run.font.size = Pt(70)
run.font.bold = False
run.font.color.rgb = GOOGLE_YELLOW

# "g" in Blue
run = p.add_run()
run.text = 'g'
run.font.name = FONT_NAME
run.font.size = Pt(70)
run.font.bold = False
run.font.color.rgb = GOOGLE_BLUE

# "l" in Green
run = p.add_run()
run.text = 'l'
run.font.name = FONT_NAME
run.font.size = Pt(70)
run.font.bold = False
run.font.color.rgb = GOOGLE_GREEN

# "e" in Red
run = p.add_run()
run.text = 'e'
run.font.name = FONT_NAME
run.font.size = Pt(70)
run.font.bold = False
run.font.color.rgb = GOOGLE_RED

# Rest of title
run = p.add_run()
run.text = ' Glass Failed'
run.font.name = FONT_NAME
run.font.size = Pt(70)
run.font.bold = True
run.font.color.rgb = GOOGLE_NEWS_BLACK

# Save the Presentation
file_name = 'google_glass_presentation.pptx'
prs.save(file_name)

print(f"Presentation '{file_name}' created successfully with one slide.")
