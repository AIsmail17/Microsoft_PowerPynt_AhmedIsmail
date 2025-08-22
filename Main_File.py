# main.py
# Before running, make sure you have the python-pptx library installed.
# You can install it using pip in your PyCharm terminal:
# pip install python-pptx

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# --- Google News Core Color Palette ---
# These colors will be used throughout the presentation for a consistent theme.
GOOGLE_BLUE = RGBColor(66, 133, 244)
GOOGLE_RED = RGBColor(219, 68, 55)
GOOGLE_YELLOW = RGBColor(244, 180, 0)
GOOGLE_GREEN = RGBColor(15, 157, 88)
TEXT_COLOR = RGBColor(68, 68, 68)  # Dark Gray for body text
TITLE_COLOR = RGBColor(0, 0, 0)      # Black for titles
BACKGROUND_COLOR = RGBColor(245, 245, 245) # Light Gray for slide background

def set_slide_background(slide, color):
    """Sets the background color of a slide."""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def create_presentation():
    """
    Generates the 'Why Google Glass Failed' presentation.
    """
    prs = Presentation()
    # Set the presentation dimensions for widescreen (16:9)
    prs.slide_width = Inches(16)
    prs.slide_height = Inches(9)

    # --- Slide Content ---
    # A list of dictionaries, where each dictionary represents a slide.
    slides_content = [
        {
            "layout": "title",
            "title": "Why Google Glass Failed: A Post-Mortem",
            "subtitle": "An analysis of the ambitious project's rise and fall.",
            "presenter": "A presentation by Gemini",
            "image_note": "(Image: A picture of someone wearing Google Glass in a futuristic, yet slightly awkward, pose.)"
        },
        {
            "layout": "content",
            "title": "What Was Google Glass?",
            "points": [
                "A wearable, voice-controlled Android device.",
                "It featured a small screen, a camera, and a bone-conduction speaker.",
                "It was designed to provide a hands-free, augmented reality experience.",
                "The \"Explorer Edition\" was released to developers and early adopters in 2013 for $1,500."
            ],
            "image_note": "(Image: A close-up shot of the Google Glass device, highlighting its different components.)"
        },
        {
            "layout": "split",
            "title": "The Hype vs. The Reality",
            "left_title": "The Hype:",
            "left_points": [
                "A revolutionary device that would change the way we interact with technology.",
                "Featured in high-fashion magazines and worn by celebrities.",
                "Promised a future of seamless augmented reality."
            ],
            "right_title": "The Reality:",
            "right_points": [
                "A clunky, unfinished prototype.",
                "Limited functionality and poor battery life.",
                "A host of technical and social problems."
            ],
            "image_note": "(Image: A split-screen image with a glamorous fashion shoot of Google Glass on one side, and a picture of someone looking confused while trying to use it on the other.)"
        },
        {
            "layout": "content",
            "title": "Reason 1: The Prohibitive Price Point",
            "points": [
                "The Explorer Edition cost $1,500.",
                "This was far too expensive for the average consumer, especially for a first-generation device.",
                "The high price created an image of elitism and exclusivity.",
                "The bill of materials was estimated to be around $80, making the price tag harder to justify."
            ],
            "image_note": "(Image: A graphic with a large \"$1,500\" price tag next to a picture of Google Glass.)"
        },
        {
            "layout": "content",
            "title": "Reason 2: Major Privacy Concerns",
            "points": [
                "The built-in camera raised serious privacy questions.",
                "The ability to record video and take pictures discreetly made people uncomfortable.",
                "The term \"Glasshole\" was coined to describe users perceived as invading others' privacy.",
                "Many businesses, such as bars and movie theaters, banned the use of Google Glass."
            ],
            "image_note": "(Image: A \"No Google Glass Allowed\" sign.)"
        },
        {
            "layout": "content",
            "title": "Reason 3: Awkward and Unfashionable Design",
            "points": [
                "The device was bulky, lopsided, and generally considered unattractive.",
                "It was not a device people felt comfortable wearing in public.",
                "The design screamed \"tech gadget\" rather than \"fashion accessory.\"",
                "Despite collaborations with fashion designers, the fundamental design was a major turn-off."
            ],
            "image_note": "(Image: A collage of people wearing Google Glass and looking awkward or out of place.)"
        },
        {
            "layout": "content",
            "title": "Reason 4: Lack of a \"Killer App\" or Clear Purpose",
            "points": [
                "Google Glass was a solution in search of a problem.",
                "There was no single, compelling use case that made the device a \"must-have.\"",
                "The functionality it offered could be easily replicated by a smartphone.",
                "Without a clear purpose, it was hard to justify the high price and social awkwardness."
            ],
            "image_note": "(Image: A question mark superimposed over an image of Google Glass.)"
        },
        {
            "layout": "content",
            "title": "Reason 5: Social Awkwardness and the \"Glasshole\" Effect",
            "points": [
                "Wearing Google Glass in public was a socially awkward experience.",
                "It created a barrier between the user and the people they were interacting with.",
                "The device was seen as a sign of social disengagement and a potential invasion of privacy.",
                "The negative social stigma was a major deterrent for potential users."
            ],
            "image_note": "(Image: A cartoon depicting a \"Glasshole\" ignoring the people around them.)"
        },
        {
            "layout": "content",
            "title": "The Enterprise Edition: A Second Life?",
            "points": [
                "Google eventually discontinued the consumer version of Glass in 2015.",
                "However, the company pivoted to the enterprise market with the \"Glass Enterprise Edition.\"",
                "This version has found success in industries like manufacturing, logistics, and healthcare.",
                "In these contexts, the hands-free functionality provides real value."
            ],
            "image_note": "(Image: A factory worker wearing the Google Glass Enterprise Edition.)"
        },
        {
            "layout": "content",
            "title": "Conclusion: Lessons Learned",
            "points": [
                "Don't release a prototype as a consumer product.",
                "Privacy is paramount in the age of connected devices.",
                "Design and social acceptability matter for wearable technology.",
                "A clear value proposition is essential for a new product to succeed."
            ],
            "image_note": "(Image: The Google logo.)"
        }
    ]

    # --- Slide Generation Loop ---
    for i, content in enumerate(slides_content):
        if content["layout"] == "title":
            # Use the Title Slide layout
            slide_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(slide_layout)
            set_slide_background(slide, BACKGROUND_COLOR)

            # --- Title ---
            title = slide.shapes.title
            title.text = content["title"]
            title.text_frame.paragraphs[0].font.color.rgb = TITLE_COLOR
            title.text_frame.paragraphs[0].font.size = Pt(60)
            title.text_frame.paragraphs[0].font.bold = True

            # --- Subtitle ---
            subtitle = slide.placeholders[1]
            subtitle.text = f"{content['subtitle']}\n\n{content['presenter']}"
            subtitle.text_frame.paragraphs[0].font.color.rgb = TEXT_COLOR
            subtitle.text_frame.paragraphs[0].font.size = Pt(24)
            subtitle.text_frame.paragraphs[2].font.italic = True
            subtitle.text_frame.paragraphs[2].font.size = Pt(20)

        elif content["layout"] == "content":
            # Use the Title and Content layout
            slide_layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(slide_layout)
            set_slide_background(slide, BACKGROUND_COLOR)

            # --- Title ---
            title = slide.shapes.title
            title.text = content["title"]
            title.text_frame.paragraphs[0].font.color.rgb = GOOGLE_BLUE
            title.text_frame.paragraphs[0].font.bold = True

            # --- Body Content ---
            body_shape = slide.shapes.placeholders[1]
            tf = body_shape.text_frame
            tf.clear() # Clear existing text
            for point in content["points"]:
                p = tf.add_paragraph()
                p.text = point
                p.font.color.rgb = TEXT_COLOR
                p.font.size = Pt(28)
                p.level = 0

            # --- Image Placeholder ---
            # Add a box to indicate where the image should go
            txBox = slide.shapes.add_textbox(Inches(9), Inches(2.5), Inches(6), Inches(4))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = "Add Image Here:\n" + content["image_note"]
            p.font.size = Pt(18)
            p.font.italic = True
            p.font.color.rgb = RGBColor(150, 150, 150)
            p.alignment = PP_ALIGN.CENTER

        elif content["layout"] == "split":
            # Custom layout for Hype vs. Reality
            slide_layout = prs.slide_layouts[5] # Blank slide layout
            slide = prs.slides.add_slide(slide_layout)
            set_slide_background(slide, BACKGROUND_COLOR)

            # --- Main Title ---
            title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(15), Inches(1.5))
            title_shape.text = content["title"]
            p = title_shape.text_frame.paragraphs[0]
            p.font.size = Pt(44)
            p.font.bold = True
            p.font.color.rgb = GOOGLE_BLUE
            p.alignment = PP_ALIGN.LEFT

            # --- Left Column (Hype) ---
            left_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(7), Inches(5))
            tf_left = left_box.text_frame
            p_left_title = tf_left.add_paragraph()
            p_left_title.text = content["left_title"]
            p_left_title.font.bold = True
            p_left_title.font.size = Pt(32)
            p_left_title.font.color.rgb = GOOGLE_GREEN
            for point in content["left_points"]:
                p = tf_left.add_paragraph()
                p.text = point
                p.font.size = Pt(24)
                p.font.color.rgb = TEXT_COLOR
                p.level = 1

            # --- Right Column (Reality) ---
            right_box = slide.shapes.add_textbox(Inches(8.5), Inches(2), Inches(7), Inches(5))
            tf_right = right_box.text_frame
            p_right_title = tf_right.add_paragraph()
            p_right_title.text = content["right_title"]
            p_right_title.font.bold = True
            p_right_title.font.size = Pt(32)
            p_right_title.font.color.rgb = GOOGLE_RED
            for point in content["right_points"]:
                p = tf_right.add_paragraph()
                p.text = point
                p.font.size = Pt(24)
                p.font.color.rgb = TEXT_COLOR
                p.level = 1

            # --- Image Placeholder at the bottom ---
            img_note_box = slide.shapes.add_textbox(Inches(0.5), Inches(7.5), Inches(15), Inches(1))
            tf_img = img_note_box.text_frame
            p_img = tf_img.add_paragraph()
            p_img.text = "Add Image Here: " + content["image_note"]
            p_img.font.size = Pt(18)
            p_img.font.italic = True
            p_img.font.color.rgb = RGBColor(150, 150, 150)
            p_img.alignment = PP_ALIGN.CENTER


    # --- Save the presentation ---
    file_path = "Google_Glass_Post-Mortem.pptx"
    prs.save(file_path)
    return file_path

if __name__ == '__main__':
    # Generate the presentation and print the file path
    generated_file = create_presentation()
    print(f"Presentation successfully generated: {generated_file}")

