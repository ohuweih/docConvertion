import os
import subprocess 
import argparse
import logging
import formatting
import imageConverter
import xlsxConverter
import shutil

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

#    pandoc_path = shutil.which("pandoc") or os.path.join(os.getcwd(), "pandoc")

#    if not os.path.exists(pandoc_path):
#        print("Pandoc executable not found. Ensure Pandoc is installed and available in the system PATH.")
#        return
    
    try:
        if not os.path.exists(media_folder): 
            os.makedirs(media_folder)
        command = f"pandoc -f docx -t asciidoc --default-image-extension .png --extract-media={media_folder} -o {output_file[:-15]}/{output_file} {input_file}"

        subprocess.run(command, check=True) 
        print(f"Command executed successfully. Output saved to {output_file}") 

    except subprocess.CalledProcessError as e: 
        print(f"Failed to execute command: {e}") 
        

def process_content(content):
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
    parser.add_argument("-o", "--output", required=True, help="name of file to convert to")

    args = parser.parse_args()


    if ".docx" in args.input:
        media_folder = f"{args.output}/extracted_media/"
        run_pandoc(media_folder, args.input, f"{args.output}_no_format.adoc")
        fix_asciidoc(f"{args.output}/{args.output}_no_format.adoc", f"{args.output}.adoc")
        imageConverter.convert_images_to_png(media_folder)
    elif ".xlsx" in args.input:
        image_output_dir = f"{args.output}/extracted_images/"
        xlsxConverter.convert_xlsx_to_adoc_with_images(args.input, args.output, image_output_dir)
    else:
        print("File not supported: Expected a docx or xlsx file")

if __name__ == "__main__":
    main()
