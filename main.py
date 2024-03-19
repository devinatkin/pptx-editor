"""
Flask Wrapper for File Shrinking API

Currently takes pptx files and shrinks them to a smaller size by reducing the size of the images within the pptx file.

"""

from flask import Flask, render_template, request, send_file
from shrinkpptx import shrink_power_point
import pptx
from editpptx import process_replacement_dict
import tempfile
import os
app = Flask(__name__)

@app.route("/")
def main_page():
    return render_template("home.html")

@app.route("/privacy")
def privacy():
    return render_template('privacy.html')

@app.route('/shrink_pptx', methods=['POST'])
def shrink_pptx():
    """
    Shrink the images in a PowerPoint presentation to a specified size

    :return: The shrunk PowerPoint presentation
    """
    # Get the file from the request
    file = request.files['file']

    # Create a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        # Save the file to the temporary directory
        file_path = temp_dir + '/input.pptx'
        file.save(file_path)

        # Shrink the PowerPoint presentation
        shrunk_file_path = temp_dir + 'output.pptx'
        shrunk_file_path = shrink_power_point(file_path, shrunk_file_path)

# Return the shrunk PowerPoint presentation
@app.route('/update_pptx', methods=['POST'])
def update_pptx():
    """
    Update the PowerPoint presentation based on the provided JSON dictionary

    :return: The updated PowerPoint presentation
    """
    # Get the file from the request
    file = request.files['file']

    # Get the JSON dictionary from the request
    json_dict = request.get_json()

    # Create a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        # Save the file to the temporary directory
        file_path = temp_dir + '/input.pptx'
        file.save(file_path)

        # Load the PowerPoint presentation
        pptx = pptx.Presentation(file_path)

        # Perform the replacements
        process_replacement_dict(pptx, json_dict['replacements_text'], json_dict['replacements_image'])

        # Save the updated PowerPoint presentation
        updated_file_path = temp_dir + '/output.pptx'
        pptx.save(updated_file_path)

    # Return the updated PowerPoint presentation
    return send_file(updated_file_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
