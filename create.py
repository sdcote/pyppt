# python -m pip install python-pptx
import csv

from pptx import Presentation

# lines are interpreted thusly:
# layout, title, content/subtitle, notes
CONCEPT = "Concept"


def create_deck(titletxt, subtitletxt, name):
    prs = Presentation('template.pptx')
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
            if len(l) == 0:
                continue

            layout = l[0]
            if layout is None:
                layout = CONCEPT
            slide_title = l[1]
            slide_body = l[2]
            if len(l) > 3:
                slide_notes = l[3]
            else:
                slide_notes = ""

            blank_slide_layout = prs.slide_layouts.get_by_name(layout)
            slide = prs.slides.add_slide(blank_slide_layout)
            title = slide.shapes.title
            title.text = slide_title
            shapes = slide.shapes
            if layout is not "Section Header":
                body_shape = shapes.placeholders[1]
                tf = body_shape.text_frame
                tf.text = slide_body
            else:
                print("Argh: " + slide_body)

            notes_slide = slide.notes_slide
            text_frame = notes_slide.notes_text_frame
            text_frame.text = slide_notes

    prs.save(name + '.pptx')


if __name__ == "__main__":
    create_deck("Culture & Organization", "The environment of continuous delivery", "Culture")
    create_deck("Design & Architecture", "Structuring your product for success", "Design")
    create_deck("Build & Deploy", "Building your product artifacts", "Build")
    create_deck("Testing & Verification", "Product artifact quality", "Testing")
    create_deck("Information & Reporting", "Measuring success and progress", "Information")
