import os
import shutil
import uuid
import zipfile
from flask import Flask, request, send_file, jsonify, after_this_request
from werkzeug.utils import secure_filename
from lxml import etree # Using lxml for robust XML processing
from google import genai
from google.genai.types import GenerateContentConfig, Schema, Type
import json

# --- Configuration ---
UPLOAD_FOLDER = 'uploads'
TEMP_PROCESSING_FOLDER = 'temp_processing'
MODIFIED_OUTPUT_FOLDER = 'modified_output'
ALLOWED_EXTENSIONS = {'docx'}
client = genai.Client(api_key="")

# --- Flask App Initialization ---
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['TEMP_PROCESSING_FOLDER'] = TEMP_PROCESSING_FOLDER
app.config['MODIFIED_OUTPUT_FOLDER'] = MODIFIED_OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max upload size

# --- Helper Functions ---

def generate_dynamic_properties(num_elements: int) -> dict:
    """
    Generates the 'properties' part of a JSON schema for an object
    with numerically indexed string properties.

    Args:
        num_elements: The number of elements (e.g., if 3, keys will be "0", "1", "2").

    Returns:
        A dictionary suitable for the 'properties' field of a JSON schema.
    """
    properties = {}
    for i in range(num_elements):
        properties[str(i)] = {"type": "STRING"}
    return properties

def generate_full_dynamic_schema(num_elements: int, required_all: bool = True) -> dict:
    """
    Generates a full JSON schema for an object with numerically
    indexed string properties.

    Args:
        num_elements: The number of elements.
        required_all: If True, all generated properties will be marked as required.

    Returns:
        A dictionary representing the complete JSON schema.
    """
    properties = generate_dynamic_properties(num_elements)
    
    schema = {
        "type": "OBJECT",
        "properties": properties,
    }

    if required_all and num_elements > 0:
        schema["required"] = list(properties.keys())
        
    return schema

def allowed_file(filename):
    """Checks if the uploaded file has an allowed extension."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def create_folders():
    """Creates necessary folders if they don't exist."""
    for folder in [UPLOAD_FOLDER, TEMP_PROCESSING_FOLDER, MODIFIED_OUTPUT_FOLDER]:
        if not os.path.exists(folder):
            os.makedirs(folder)


def call_gemini_api(texts: list[str]) -> list[str]:
    """
    Calls the Gemini API to modify CV text segments, using a dictionary approach
    to try and maintain the exact number of elements.

    Args:
        texts: A list of strings, where each string is a segment of a CV.
        job_title: The target job title for tailoring the CV.
        company_name: The target company name.

    Returns:
        A list of strings with the modified CV segments, in the original order.

    Raises:
        ValueError: If the API returns a dictionary with a different number of
                    elements than expected or if keys are missing.
    """
    
    if not texts:
        return []

    # 1. Convert the input list to a dictionary with stringified integer keys
    # This helps maintain order and gives explicit "slots" for the LLM.
    input_cv_dict = {str(i): text for i, text in enumerate(texts)}
    num_elements = len(texts)
    job_title = "Junior web developer"
    company_name = "Google"

    # Prepare the dictionary part of the prompt
    # Ensure proper formatting for the prompt, especially quotes and newlines within f-string
    dict_prompt_lines = [f'"{key}": {json.dumps(value)}' for key, value in input_cv_dict.items()]
    dict_prompt_str = ",\n".join(dict_prompt_lines)
    
    # Construct the prompt
    prompt_content = f"""
    I have a CV represented as a Python dictionary of strings. Each key is a stringified index, and each value is a segment of text from the CV. I need you to tailor this CV for a {job_title} position at {company_name}.

    Input CV (dictionary of string segments, {num_elements} entries):
    {{
    {dict_prompt_str}
    }}

    Your task is to review EACH segment (value associated with each key). Modify its text content if necessary to align with the target role. If a segment does not require modification (e.g., it's a name, a separator, or already optimal), return that segment UNCHANGED as the value for its original key.

    **Critical Constraints:**
    1.  The output **MUST** be a JSON object representing a dictionary.
    2.  The output dictionary **MUST** have exactly {num_elements} entries.
    3.  The output dictionary **MUST** contain all the original keys from the input (i.e., from "0" to "{num_elements - 1}").
    4.  The value associated with each key in the output dictionary must correspond to the modification (or original version) of the value for that same key in the input dictionary.
    5.  **DO NOT** add new keys. **DO NOT** remove keys. **DO NOT** merge segments. Only modify the text content *within* each segment's value.

    Target Job Title: {job_title}
    Target Company: {company_name}

    Output Modified CV (JSON dictionary with keys "0" to "{num_elements - 1}" and their corresponding string values):
    {{
    "0": "...",
    "1": "...",
    ...
    "{num_elements - 1}": "..."
    }}
    """

    # This is a placeholder for your actual API call
    # Replace with your `client.models.generate_content` call
    print("---- PROMPT ----")
    print(prompt_content)
    print("----------------")

    full_schema = generate_full_dynamic_schema(num_elements)

    response = client.models.generate_content(
        model="gemini-2.0-flash", contents = prompt_content, 
        config={
            "response_mime_type": "application/json",
            "response_schema": full_schema,
            # "response_schema": Schema(
            #     type=Type.OBJECT
            # ),
        },
    )

    parsed_dict = json.loads(response.text)
    print("Response: ", response)
    print("Parsed dict: ", parsed_dict)

    # 3. Convert the output dictionary back to a list, ensuring correct order and completeness
    modified_texts_list = []
    for i in range(num_elements):
        key = str(i)
        if key not in parsed_dict:
            raise ValueError(f"API response is missing expected key: '{key}'")
        if not isinstance(parsed_dict[key], str):
            raise ValueError(f"API response for key '{key}' is not a string. Got: {type(parsed_dict[key])}")
        modified_texts_list.append(parsed_dict[key])

    print(f"Reconstructed list length: {len(modified_texts_list)}")
    print(f"Original list length: {num_elements}")

    return modified_texts_list


def process_docx(uploaded_file_path: str, original_filename: str) -> (str | None, str | None):
    """
    Core logic for processing the DOCX file.
    Returns (path_to_modified_docx, error_message)
    """
    request_id = str(uuid.uuid4())
    # Path for this specific request's temporary files
    current_temp_dir = os.path.join(app.config['TEMP_PROCESSING_FOLDER'], request_id)
    # Path for the extracted content of the DOCX
    extracted_content_dir = os.path.join(current_temp_dir, 'extracted_content')
    # Path for the copied DOCX (to be renamed to .zip)
    temp_docx_zip_path = os.path.join(current_temp_dir, f"{request_id}.zip")
    # Path to the target XML file
    document_xml_path = os.path.join(extracted_content_dir, 'word', 'document.xml')
    # Path for the final modified DOCX
    modified_docx_filename = f"modified_{original_filename}"
    output_docx_path = os.path.join(app.config['MODIFIED_OUTPUT_FOLDER'], modified_docx_filename)

    try:
        os.makedirs(current_temp_dir)
        os.makedirs(extracted_content_dir)
        os.makedirs(app.config['MODIFIED_OUTPUT_FOLDER'], exist_ok=True) # Ensure output folder exists

        # 1. Copy uploaded file to temp dir and rename to .zip
        shutil.copy(uploaded_file_path, temp_docx_zip_path)

        # 2. Extract ZIP contents
        with zipfile.ZipFile(temp_docx_zip_path, 'r') as zip_ref:
            zip_ref.extractall(extracted_content_dir)

        if not os.path.exists(document_xml_path):
            return None, "Could not find 'word/document.xml' in the DOCX file. The file might be corrupted or not a valid DOCX."

        # 3. Parse document.xml and extract text from <w:t>
        # Define the WordprocessingML namespace
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        parser = etree.XMLParser(remove_blank_text=False, strip_cdata=False, recover=True)
        tree = etree.parse(document_xml_path, parser)
        root = tree.getroot()

        original_texts = []
        w_t_elements = [] # Store references to the <w:t> elements

        # Iterate through all <w:t> elements in document order
        for wt_element in root.xpath('.//w:t', namespaces=ns):
            # Preserve existing xml:space attribute if present
            # text_content = wt_element.text if wt_element.text is not None else ""
            text_content = "".join(wt_element.itertext()) # Handles mixed content better
            original_texts.append(text_content)
            w_t_elements.append(wt_element)
            
            # Important: Clear the existing content of <w:t> before potentially adding new text.
            # This handles cases where <w:t> might contain other child elements (e.g. <w:tab/>, <w:br/>)
            # which should be preserved. We only want to replace the text nodes.
            for child in list(wt_element): # Iterate over a copy if modifying
                if child.tail: # Preserve text after a child element if any
                    pass # This text is part of the parent's mixed content, not <w:t>'s direct text.
                # wt_element.remove(child) # Don't remove children, just text nodes
            
            # Clear direct text content of <w:t>
            wt_element.text = None 
            # Also clear tail text of its children, as itertext() would have picked it up
            for child in wt_element:
                child.tail = None


        # 4. Send to (simulated) API and get modified texts
        try:
            modified_texts = call_gemini_api(original_texts)
            # modified_texts = original_texts
            print("Original list: ", original_texts)
        except ValueError as e:
            app.logger.error(e)
            return None, f"API processing error: {str(e)}"
        except Exception as e: # Catch other potential API call errors
            app.logger.error(e)
            return None, f"An unexpected error occurred during API simulation: {str(e)}"


        if len(modified_texts) != len(w_t_elements):
            return None, "Mismatch between the number of text elements found and modified texts received."

        # 5. Update <w:t> tags with modified text
        for i, wt_element in enumerate(w_t_elements):
            new_text = modified_texts[i]
            # Set the text. If new_text is None or empty, it will effectively clear the text.
            wt_element.text = new_text
            
            # If the original <w:t> tag had xml:space="preserve", try to keep it.
            # This is important for leading/trailing spaces.
            # lxml handles this fairly well by default if spaces are in the text.
            # However, if the original text was just spaces, and the new text is also just spaces,
            # Word might still collapse them if xml:space="preserve" is not explicitly on the <w:t> tag.
            # The <w:rPr> (run properties) usually handles this, so direct manipulation of xml:space
            # on <w:t> might not always be necessary if the run properties are intact.
            # For simplicity, we are not adding/modifying xml:space here, assuming
            # the surrounding structure and run properties are sufficient.
            # If issues with whitespace arise, this is an area to investigate.

        # 6. Save modified document.xml
        # Use 'xml_declaration=True', 'encoding="UTF-8"', and 'standalone=True' for Word compatibility.
        tree.write(document_xml_path, xml_declaration=True, encoding='UTF-8', standalone=True)

        # 7. Re-create the ZIP archive (new DOCX)
        with zipfile.ZipFile(output_docx_path, 'w', zipfile.ZIP_DEFLATED) as new_zip:
            for root_dir, _, files in os.walk(extracted_content_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    # arcname is the path as it should appear inside the zip file
                    arcname = os.path.relpath(file_path, extracted_content_dir)
                    new_zip.write(file_path, arcname)
        
        return output_docx_path, None

    except Exception as e:
        # Log the full exception for debugging
        app.logger.error(f"Error processing DOCX for request {request_id}: {e}", exc_info=True)
        return None, f"An internal server error occurred: {str(e)}"
    finally:
        # 8. Cleanup temporary files for this request
        if os.path.exists(current_temp_dir):
            shutil.rmtree(current_temp_dir)
        if os.path.exists(uploaded_file_path) and os.path.dirname(uploaded_file_path) == app.config['UPLOAD_FOLDER']:
             # Only delete if it's in the main upload folder, not if it was already a temp file
            try:
                os.remove(uploaded_file_path)
            except OSError as e:
                app.logger.error(f"Error deleting uploaded file {uploaded_file_path}: {e}")


# --- Flask Routes ---
@app.route('/modify_docx', methods=['POST'])
def modify_docx_route():
    """
    Endpoint to upload a DOCX file, modify its content, and return the modified file.
    """
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected for uploading"}), 400

    if file and allowed_file(file.filename):
        original_filename = secure_filename(file.filename)
        # Save uploaded file temporarily with a unique name to avoid conflicts
        # This file will be cleaned up by process_docx's finally block
        temp_upload_id = str(uuid.uuid4())
        uploaded_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{temp_upload_id}_{original_filename}")
        file.save(uploaded_file_path)

        modified_docx_path, error = process_docx(uploaded_file_path, original_filename)

        if error:
            # process_docx should have already cleaned up its own temp files.
            # The uploaded_file_path is handled by process_docx's finally block.
            return jsonify({"error": error}), 500
        
        if modified_docx_path and os.path.exists(modified_docx_path):
            # Schedule the deletion of the modified file after the request is sent
            @after_this_request
            def remove_modified_file(response):
                try:
                    if os.path.exists(modified_docx_path):
                        os.remove(modified_docx_path)
                except Exception as e:
                    app.logger.error(f"Error deleting modified DOCX file {modified_docx_path}: {e}")
                return response
            
            return send_file(
                modified_docx_path,
                as_attachment=True,
                download_name=os.path.basename(modified_docx_path), # Use the generated name
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        else:
            # This case should ideally be caught by the 'error' variable, but as a fallback:
            return jsonify({"error": "Failed to produce modified DOCX file."}), 500
            
    else:
        return jsonify({"error": "Allowed file type is .docx"}), 400

# --- Main Execution ---
if __name__ == '__main__':
    create_folders()
    # For development: app.run(debug=True)
    # For production, use a proper WSGI server like Gunicorn or Waitress.
    app.run(host='0.0.0.0', port=5000, debug=True)
