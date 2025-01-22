
from PIL import Image
import os
import subprocess 
import argparse
import re
import logging

# Configure logging
logging.basicConfig(level=logging.INFO,
                    handlers=[logging.FileHandler("my_log_file.log", mode='w', encoding='utf-8'),
                              logging.StreamHandler()])

def run_pandoc(media_folder, input_file, output_file): 
    """ 
    Run a Windows shell command and save the output to a file. 
    
        Args: command (str): The shell command to run. 
        output_file (str): Path to the file where output will be saved. 
    """ 
    try:
        command = f"/Users/ohuweih/Documents/pandoc-3.6.2-windows-x86_64/pandoc-3.6.2/pandoc.exe -f docx -t asciidoc --default-image-extension '.png' --extract-media={media_folder} -o {output_file[:-15]}/{output_file} {input_file}"
        subprocess.run(command) 
        print(f"Command executed successfully. Output saved to {output_file}") 
    except subprocess.CalledProcessError as e: 
        print(f"Failed to execute command: {e}") 
        
    

def convert_images_to_png(media_folder):
    """
    Converts all emf and wmf images in our media folder to png images
    
    Args:
        media_folder (str): Path to the folder
    """
    
    for filename in os.listdir(f"{media_folder}/media"):
        if 'png' in filename.lower():
            print('png file')
        else:
            file_path = os.path.join(f"{media_folder}/media", filename)
            png_path = os.path.splitext(file_path)[0] + ".png"

            try:
                with Image.open(file_path) as img:
                    img.save(png_path, "PNG")
                    print(f"Converted: {filename} -> {os.path.basename(png_path)}")
                os.remove(file_path)  # Remove the original EMF/WMF file
            except Exception as e:
                print(f"Failed to convert {filename}: {e}")


def escape_double_angle_brackets(content):
    pattern = r"<<(.*?)>>"
    return re.sub(pattern, r"\<<\1>>", content)


def recolor_notes(content):
    notes_patterns = [
        re.compile(r'(?<!foot)(?<!\[)Note:.*'),
        re.compile(r'Note\s+\d+.*'),
        re.compile(r'EXAMPLE\s+\d+:.*'),
        re.compile(r'Please note:.*')
    ]

    matches = set()
    for pattern in notes_patterns:
        matches.update(re.findall(pattern, content))

    # Iterate over the matches and modify the figure tags
    for match in matches:
        content = content.replace(match, f'\n====\n{match.strip()}\n====\n')

    return content


def remove_lines(content, start_line, end_line):
    lines = content.splitlines()
    lines_to_keep = lines[:start_line - 1] + lines[end_line:]
    return '\n'.join(lines_to_keep)


def remove_text_by_patterns(content):
    regular_exp_list = [
        r'Table of Contents\n\n(.*?)(?=\n\n==)',
        r'\[#_Toc\d* \.anchor]####Table \d*\:?\s?',
        r'\[#_Ref\d* \.anchor]####Table \d*\:?\s?',
        r'\[\#_Ref\d* \.anchor\](?:\#{2,4})',
        r'\[\#_Toc\d* \.anchor\](?:\#{2,4})',
        r'\{empty\}'
    ]

    for regex in regular_exp_list:
        # Remove occurrences of the specified regular expression
        content = re.sub(regex, '', content)
    return content


def use_block_tag_for_img_and_move_caption_ahead(content):
    # Define the callback function
    def replacement(match):
        # Use block figure tag
        figure_tag: str = match[1]
        figure_tag = figure_tag.replace("image:", "image::")

        # remove "Figure+num" as it will be automatically done in AsciiDoc
        caption: str = match[2]
        caption = re.sub(r'^Figure \d*:?', '', caption)

        # Move the caption to the beginning of the figure tag
        modified_figure_tag = f'.{caption.strip()}\n{figure_tag}\n'

        return modified_figure_tag

    # Define the regular expression pattern to match figure tags with specific captions
    pattern = r'(image:\S+\[.*?\])\s+?\n?\n?(Figure.*?\n)'

    # Replace the original figure tags with the modified ones
    new_content = re.sub(pattern, replacement, content)
    return new_content


def escape_source_square_brackets(content):
    pattern = r"\[(SOURCE:.*?)\]"
    return re.sub(pattern, r"&#91;\1&#93;", content)


def find_bibliography_section(content):
    # Find the position of "== Bibliography"
    bibliography_pos = content.find("== Bibliography")
    if bibliography_pos == -1:
        logging.warning("Bibliography section not found")
        return None
    return bibliography_pos


def add_anchors_to_bibliography(content):
    bibliography_pos = find_bibliography_section(content)
    if bibliography_pos is None:
        return {}, content

    # Extract the substring starting from the position of "== Bibliography"
    bibliography_text = content[bibliography_pos:]
    biblio_pattern = re.compile(r'^\[(\d+)\](.+)', re.MULTILINE)

    matches = {key: val for key, val in biblio_pattern.findall(bibliography_text)}

    for biblio_tag_num, biblio_tag_text in matches.items():
        anchor = f'[#bib{biblio_tag_num}]'
        modified_text = f'{anchor}\n[{biblio_tag_num}]{biblio_tag_text}'
        bibliography_text = bibliography_text.replace(f'[{biblio_tag_num}]{biblio_tag_text}', modified_text)

    content = content[:bibliography_pos] + bibliography_text
    return matches.keys(), content


def add_links_to_bibliography(content, keys):
    for key in keys:
        in_link_patterns = r'(?<!\[#bib{key}\]\n)\[{key}\]'
        in_link_patterns = in_link_patterns.format(key=key)
        content = re.sub(in_link_patterns, f'link:#bib{key}[[{key}\]]', content)
    return content


def remove_bad_plus_syntax(content):
    pattern = "\+\+"
    return re.sub(pattern, '', content)

def process_content(content):
    logging.info("Removing certain patterns in asciidoc file")
    content = remove_text_by_patterns(content)

    logging.info("Escaping double angle brackets")
    content = escape_double_angle_brackets(content)

    logging.info("Styling note boxes")
    content = recolor_notes(content)

    logging.info("Adding anchors to bibliography")
    keys, content = add_anchors_to_bibliography(content)

    logging.info("Connecting in-document references to bibliography")
    content = add_links_to_bibliography(content, keys)

    logging.info("Fixing image captions")
    content = use_block_tag_for_img_and_move_caption_ahead(content)

    logging.info("Escaping square brackets")
    content = escape_source_square_brackets(content)

    logging.info("Removing bad ++ patters")
    content = remove_bad_plus_syntax(content)

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

    content = process_content(content)
    write_output(f"{output_file[:-5]}/{output_file}", content)
    os.remove(input_file) 


def main():
    '''
    Main fuction will take two arguments (input_file and output_file) and run a docx to asciidoc conversion of the file, put all media in to its own folder "/media/media", edit the asciidoc file to change all .emf strings to .png strings and convert all emf images to png images"
    
    '''
    parser = argparse.ArgumentParser(description="convert docx to adoc, including image support")
    parser.add_argument("-i", "--input", required=True, help="Docx to convert")
    parser.add_argument("-t", "--to", required=True, help="name of file to convert to")

    args = parser.parse_args()
    media_folder = f"{args.to}/extracted_media/"

    run_pandoc(media_folder, args.input, f"{args.to}_no_format.adoc")
    
    fix_asciidoc(f"{args.to}/{args.to}_no_format.adoc", f"{args.to}.adoc")

    convert_images_to_png(media_folder)

if __name__ == "__main__":
    main()
