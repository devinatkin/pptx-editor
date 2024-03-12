import argparse
from zipfile import ZipFile
import os
import tempfile
from PIL import Image
import cairosvg

def shrink_power_point(input_file_path, output_file_path, size = 1024):
    """
    Shrink the images in a PowerPoint presentation to a specified size

    :param input_file_path: Path to the input PowerPoint presentation
    :param output_file_path: Path to save the shrunk PowerPoint presentation
    :param size: Size to shrink the images to (default is 1024)

    :return: Path to the shrunk PowerPoint presentation
    """
    # Open the input file
    with ZipFile(input_file_path, 'r') as input_file:
        # Open the input file
        file_list = input_file.namelist()

        # Grab any images within the media folder
        media_files = [f for f in file_list if 'ppt/media/' in f]

        # Create a temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            # Extract the media files to the temporary directory
            input_file.extractall(temp_dir, file_list)

            # Shrink the images
            for file in media_files:
                file_path = os.path.join(temp_dir, file)

                # Check if the file is an SVG
                if file_path.endswith('.svg'):
                    # Convert the SVG to PNG using CairoSVG
                    png_file_path = file_path.replace('.svg', '.png')
                    cairosvg.svg2png(url=file_path, write_to=png_file_path)
                    file_path = png_file_path

                    file_list.remove(file)
                    file_list.append(file.replace('.svg', '.png'))



                image = Image.open(file_path)
                image.thumbnail((size,size))

                image.save(file_path)

            # Create a new PowerPoint presentation
            print("Creating new PowerPoint presentation")
            with ZipFile(output_file_path, 'w') as output_file:
                # Add the files from the original presentation to the new presentation
                for file in file_list:
                    if file not in media_files:
                        output_file.write(os.path.join(temp_dir, file), file)

                # Add the media files to the new presentation
                for file in media_files:
                    output_file.write(os.path.join(temp_dir, file), file)

    return output_file_path


def main():
    # Create the argument parser
    parser = argparse.ArgumentParser(description='Shrink PowerPoint presentation')

    # Add the input file path argument
    parser.add_argument('input_file', type=str, help='Path to the input PowerPoint presentation')

    # Add the output file path argument
    parser.add_argument('output_file', type=str, help='Path to save the shrunk PowerPoint presentation')

    # Parse the command line arguments
    args = parser.parse_args()

    # Call the shrink_power_point function with the provided input and output file paths
    shrink_power_point(args.input_file, args.output_file)

if __name__ == '__main__':
    main()