import os
import logging
import formatting
import imageConverter
import xlsxConverter
import pandoc
from tqdm import tqdm
import tkinter as tk
from tkinter import filedialog, messagebox

# Configure logging
logging.basicConfig(level=logging.INFO,
                    handlers=[logging.FileHandler("my_log_file.log", mode='w', encoding='utf-8'),
                              logging.StreamHandler()])
        
def process_content(content, output_file):
    logging.info("Removing certain patterns in asciidoc file")
    content = formatting.remove_text_by_patterns(content)

    logging.info("changing all image links to .png")
    content = formatting.replace_image_suffix_to_png(content)

    logging.info("Escaping double angle brackets")
    content = formatting.escape_double_angle_brackets(content)

    logging.info("Styling note boxes")
    content = formatting.recolor_notes(content)

    logging.info("Adding anchors to bibliography")
    keys, content = formatting.add_anchors_to_bibliography(content)

    logging.info("Connecting in-document references to bibliography")
    content = formatting.add_links_to_bibliography(content, keys)

    logging.info("Fixing image captions")
    content = formatting.use_block_tag_for_img_and_move_caption_ahead(content)

    logging.info("Escaping square brackets")
    content = formatting.escape_source_square_brackets(content)

    logging.info("Removing bad ++ patters")
    content = formatting.remove_bad_plus_syntax(content)
    
    logging.info("fixing image file paths")
    content = formatting.fix_image_file_path(content, output_file)

    return content


def write_output(output_file, content):
    logging.info(f"Writing fixed content to the output file: {output_file}")
    with open(output_file, 'w', encoding="utf-8") as file:
        file.write(content)


def fix_asciidoc(input_file, output_file):
    #directory = pathlib.Path(input_file).parent

    logging.info("Read the initial asciidoc file...")
    with open(input_file, 'r', encoding="utf-8") as file:
        content = file.read()

    content = process_content(content, output_file)
    write_output(f"{output_file}/{output_file}.adoc", content)
    os.remove(input_file) 




def main():
    '''
    Main fuction will take two arguments (input_file and output_file) and run a docx to asciidoc conversion of the file, put all media in to its own folder "/media/media", edit the asciidoc file to change all .emf strings to .png strings and convert all emf images to png images"
    
    '''
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    messagebox.showinfo("Select Input File", "Please select the input file (e.g., .docx).")
    input_file = filedialog.askopenfilename(
        title="Select Input File",
        filetypes=[("Word Documents", "*.docx"), ("Excel Spreedsheets", "*.xlsx"), ("All Files", "*.*")]
    )
    if not input_file:
        print("No input file selected.")
        return
    
    file_dir = os.path.dirname(input_file)
    file_name = os.path.basename(input_file)
    file_stem = os.path.splitext(file_name)[0]



    docx_steps = ["Determine if docx or xlsx", "Extract Images", "Initail Convertion", "Formatting", "Convert Images to PNG"]
    xlsx_steps = ["Determine if docx or xlsx", "Converting xlsx"]

    if ".docx" in input_file:
        for step in tqdm(docx_steps, desc="Overall Progress", unit="step"):
            if step == "Determine if docx or xlsx":
                print("Detected docx file")
            if step == "Extract Images":
                print("Extract Images")
                media_folder = f"{file_stem}/extracted_media/"
            if step == "Initail Convertion":
                print("Doing Initail Convertion")
                pandoc.run_pandoc(media_folder, input_file, f"{file_stem}")
            if step == "Formatting":
                print("Formatting as best we can")
                fix_asciidoc(f"{file_stem}/{file_stem}_no_format.adoc", f"{file_stem}")
            if step == "Convert Images to PNG":
                print("Converting images to png")
                imageConverter.convert_images_to_png(media_folder)
    elif ".xlsx" in input_file:
        for step in tqdm(xlsx_steps, desc="Overall Progress", unit="step"):
            if step == "Determine if docx or xlsx":
                print("Detected xlsx file")
            if step == "Extract Images":
                print("Extract Images")
                image_output_dir = f"{file_stem}/extracted_images/"
            if step == "Converting xlsx":
                print("Doing Initail Convertion")
                print("Formatting as best we can")
                print("Converting images to png")
                xlsxConverter.convert_xlsx_to_adoc_with_images(input_file, f"{file_stem}", image_output_dir)
        else:
            print("File not supported: Expected a docx or xlsx file")

if __name__ == "__main__":
    main()
