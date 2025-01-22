from PIL import Image
import os
import subprocess 
import argparse

def run_shell_command(input_file, output_file): 
    """ 
    Run a Windows shell command and save the output to a file. 
    
        Args: command (str): The shell command to run. 
        output_file (str): Path to the file where output will be saved. 
    """ 
    try:
        command = f"/Users/ohuweih/Documents/pandoc-3.6.2-windows-x86_64/pandoc-3.6.2/pandoc.exe -f docx -t asciidoc --default-image-extension '.png' --extract-media=docx_to_adoc_conversion/{output_file[:-5]}/media -o docx_to_adoc_conversion/{output_file[:-5]}/{output_file} {input_file}"
        subprocess.run(command) 
        print(f"Command executed successfully. Output saved to {output_file}") 
    except subprocess.CalledProcessError as e: 
        print(f"Failed to execute command: {e}") 
        
        
def edit_output_file(output_file): 
    """
    Edit a file to replace all occurrences of '.emf' with '.png'. 
    
        Args: output_file (str): Path to the file to edit. 
    """ 
    try: 
        file = f"docx_to_adoc_conversion/{output_file}"
        with open(file, "r", encoding="utf-8") as f:
            content = f.read() 
            # Replace all '.emf' occurrences with '.png'
            updated_content = content.replace(".emf", ".png")
            updated_content = content.replace(".wmf", ".png") 
            
            with open(file, "w", encoding="utf-8") as f:
                f.write(updated_content) 
                print(f"Updated file: {output_file}") 
    except Exception as e: 
        print(f"Failed to edit file {output_file}: {e}") 


def convert_emf_to_png(media_folder):
    """
    Converts all emf and wmf images in our media folder to png images
    
    Args:
        media_folder (str): Path to the folder
        """
    
    for filename in os.listdir(media_folder):
        if filename.lower().endswith(".emf"):
            emf_path = os.path.join(media_folder, filename)
            png_path = os.path.join(media_folder, f"{os.path.splitext(filename)[0]}.png")

            try:
                with Image.open(emf_path) as img:
                    img.save(png_path, "PNG")
                    print(f"Converted: {filename} -> {os.path.basename(png_path)}")
            except Exception as e:
                print(f"Failed to convert {filename}: {e}")

        if filename.lower().endswith(".wmf"):
            wmf_path = os.path.join(media_folder, filename)
            png_path = os.path.join(media_folder, f"{os.path.splitext(filename)[0]}.png")

            try:
                with Image.open(wmf_path) as img:
                    img.save(png_path, "PNG")
                    print(f"Converted: {filename} -> {os.path.basename(png_path)}")
            except Exception as e:
                print(f"Failed to convert {filename}: {e}")


def main():
    '''
    Main fuction will take two arguments (input_file and output_file) and run a docx to asciidoc conversion of the file, put all media in to its own folder "/media/media", edit the asciidoc file to change all .emf strings to .png strings and convert all emf images to png images"
    
    '''
    parser = argparse.ArgumentParser(description="convert docx to adoc, including image support")
    parser.add_argument("-i", "--input", required=True, help="Docx to convert")
    parser.add_argument("-t", "--to", required=True, help="name of file to convert to")

    args = parser.parse_args()
    media_folder = f"./docx_to_adoc_conversion/{args.to[:-5]}/media/media/"

    run_shell_command(args.input, args.to)
    edit_output_file(args.to)
    convert_emf_to_png(media_folder)

if __name__ == "__main__":
    main()
