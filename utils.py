from datetime import datetime, timedelta

def current_monday():
    today = datetime.now()
    days_to_subtract = today.weekday() + 7  
    last_monday = today - timedelta(days=days_to_subtract)
    return last_monday.strftime("%d/%m/%Y")

def previous_monday():
    today = datetime.now()
    days_to_subtract = today.weekday() + 7 +7
    last_monday = today - timedelta(days=days_to_subtract)
    return last_monday.strftime("%d/%m/%Y")


def current_sunday():
    today = datetime.now()
    days_to_subtract = today.weekday() + 7 - 6
    last_sunday = today - timedelta(days=days_to_subtract)
    return last_sunday.strftime("%d/%m/%Y")

def previous_sunday():
    today = datetime.now()
    days_to_subtract = today.weekday() + 7 - 6+7
    last_sunday = today - timedelta(days=days_to_subtract)
    return last_sunday.strftime("%d/%m/%Y")


#work in first slide
def first_slide_work(presentation,current_monday,current_sunday,previous_monday,previous_sunday):
  first_slide = presentation.slides[0]

  for shape in first_slide.shapes:
    if shape.has_text_frame and previous_monday in shape.text_frame.text:
        existing_font = shape.text_frame.paragraphs[0].runs[0].font

        new_paragraph = shape.text_frame.add_paragraph()
        new_paragraph.text = current_monday

        # Apply the font style to the new text
        new_run = new_paragraph.runs[0]
        new_run.font.size = existing_font.size
        new_run.font.name = existing_font.name
        new_run.font.bold = existing_font.bold
        new_run.font.italic = existing_font.italic
        new_run.font.color.theme_color = existing_font.color.theme_color
        new_run.font.color.brightness = existing_font.color.brightness

  for shape in first_slide.shapes:
    if shape.has_text_frame and previous_sunday in shape.text_frame.text:
        existing_font = shape.text_frame.paragraphs[0].runs[0].font

        new_paragraph = shape.text_frame.add_paragraph()
        new_paragraph.text = current_sunday

        # Apply the font style to the new text
        new_run = new_paragraph.runs[0]
        new_run.font.size = existing_font.size
        new_run.font.name = existing_font.name
        new_run.font.bold = existing_font.bold
        new_run.font.italic = existing_font.italic
        new_run.font.color.theme_color = existing_font.color.theme_color
        new_run.font.color.brightness = existing_font.color.brightness





def replace_text_in_shape(shape, old_text, new_text):
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                run.text = run.text.replace(old_text, new_text)



def firstPageChange(presentation,old,new):
    first_slide = presentation.slides[0]
    for shape in first_slide.shapes:
      replace_text_in_shape(shape, old, new)
