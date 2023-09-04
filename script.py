from pptx import Presentation
import argparse

parser = argparse.ArgumentParser(
                    prog='pptx_to_notes',
                    description='Converts a pptx presentation into a abridged word doc (i.e. notes)',
                    epilog='very cool')


def init_presentation(path):
    """Initializes a Presentation object representing a pptx file"""
    try:
        prs = Presentation(path)
    except FileNotFoundError:
        print(
            "No file exists at that file location. Double check your file path to make sure it's correct."
        )
    return prs


def extract_text(prs):
    """Extracts all the text from the given presentation for later"""
    extracted_text = []

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    extracted_text.append(run.text)
    return extracted_text
