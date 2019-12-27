import csv

# python -m pip install python-pptx
from pptx import Presentation


# Functions go here


def create_deck(titletxt, subtitletxt, name):
    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = titletxt
    subtitle.text = subtitletxt
    with open(name + '.csv', 'rb') as f:
        reader = csv.reader(f)
        my_list = list(reader)

        for l in my_list:
            blank_slide_layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(blank_slide_layout)
            title = slide.shapes.title
            title.text = titletxt
            shapes = slide.shapes
            body_shape = shapes.placeholders[1]
            tf = body_shape.text_frame
            tf.text = l[0]
            notes_slide = slide.notes_slide
            text_frame = notes_slide.notes_text_frame
            text_frame.text = l[1]

    prs.save(name + '.pptx')


if __name__ == "__main__":
    create_deck("Culture & Organization", "The environment of continuous delivery", "Culture")
    create_deck("Design & Architecture","Structuring your product for success","Design")
    create_deck("Build & Deploy","Building your product artifacts","Build")
    create_deck("Testing & Verification","Product artifact quality","Testing")
    create_deck("Information & Reporting","Measuring success and progress","Information")
