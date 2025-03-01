import os
import docx
import json
import argparse

def list_config_files(config_dir="config"):
    """ Lists all available JSON config files and extracts product names. """
    if not os.path.exists(config_dir):
        print(f"❌ Config directory '{config_dir}' not found.")
        return []

    config_files = [f for f in os.listdir(config_dir) if f.endswith(".json")]
    config_info = []

    for file in config_files:
        file_path = os.path.join(config_dir, file)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
                product_name = config_data.get("productName", "Unknown Product")
                config_info.append((file, product_name))
        except Exception as e:
            print(f"⚠️  Skipping '{file}': Unable to read or invalid JSON format. ({e})")

    if not config_info:
        print(f"❌ No valid JSON config files found in '{config_dir}'.")
    
    return config_info


def get_user_selected_config(config_info, config_dir="config"):
    """ Prompts the user to select a config file from available options. """
    while True:
        print("\n📌 Available Configurations:")
        for idx, (file, product_name) in enumerate(config_info, start=1):
            print(f"{idx}. {product_name} ({file})")

        choice = input("\nEnter the number of the configuration to use (or 'q' to quit): ")
        if choice.lower() == 'q':
            print("\n🛑 Exiting selection...")
            return None  # Return None if user exits
        try:
            choice = int(choice)
            if 1 <= choice <= len(config_info):
                return os.path.join(config_dir, config_info[choice - 1][0])
            else:
                print("❌ Invalid selection. Please enter a valid number.")
        except ValueError:
            print("❌ Invalid input. Please enter a number.")



def read_data(file_path):
    """ Reads tab-separated input data from a file and converts it to a dictionary. """
    if not os.path.exists(file_path):
        return {}, [f"❌ Data file '{file_path}' not found."]

    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    if len(lines) < 2:
        return {}, [f"❌ Data file '{file_path}' is missing required rows."]

    headers = [h.strip() for h in lines[0].split('\t')]
    values = [v.strip() for v in lines[1].split('\t')]

    return dict(zip(headers, values)), []

def read_config(config_path):
    """ Reads the JSON config file containing template paths and mappings. """
    if not os.path.exists(config_path):
        return {}, [f"❌ Config file '{config_path}' not found."]

    with open(config_path, 'r', encoding='utf-8') as file:
        config = json.load(file)

    required_keys = ["templatePath", "outputPath", "mappings"]
    missing_keys = [key for key in required_keys if key not in config]

    if missing_keys:
        return {}, [f"❌ Config file is missing required keys: {', '.join(missing_keys)}"]

    # Check for duplicate placeholders
    seen_placeholders = set()
    duplicates = []
    for key, placeholder in config["mappings"].items():
        if placeholder in seen_placeholders:
            duplicates.append(placeholder)
        seen_placeholders.add(placeholder)

    if duplicates:
        return {}, [f"❌ Config contains duplicate placeholders: {', '.join(duplicates)}"]

    return config, []

def validate_docx_file(file_path):
    """ Validates if the given .docx file is readable and properly formatted. """
    if not os.path.exists(file_path):
        return False, f"❌ Template file '{file_path}' not found."

    try:
        docx.Document(file_path)
        return True, ""
    except Exception as e:
        return False, f"❌ Failed to read '{file_path}'. The file might be corrupted or not a valid .docx.\n{e}"

def log_mappings(data, config, placeholder_counts, placeholder_line_counts, missing_placeholders, error_messages, verbose=False):
    """ Logs the mappings and placeholder statistics for debugging. """
    print("\n📌 **Placeholder Replacement Log:**")

    max_col_len = max(len(col) for col in data.keys()) if data else 0
    max_placeholder_len = max(len(config["mappings"].get(col, "")) for col in data.keys() if col in config["mappings"]) if data else 0
    max_value_len = max(len(value) for value in data.values()) if data else 0
    max_count_len = max(len(str(count)) for count in placeholder_counts.values()) if placeholder_counts else 0

    col_width = max(max_col_len, len("INPUT"))
    placeholder_width = max(max_placeholder_len, len("PLACEHOLDER"))
    value_width = max(max_value_len, len("VALUE"))
    count_width = max(max_count_len, len("COUNT"))

    print(f"\n{'INPUT'.ljust(col_width)}      {'PLACEHOLDER'.ljust(placeholder_width)}      {'VALUE'.ljust(value_width)}      {'COUNT'.rjust(count_width)}")

    for col, value in data.items():
        if col in config["mappings"]:
            placeholder = config["mappings"][col]
            count = placeholder_counts.get(placeholder, 0)
            print(f"{col.ljust(col_width)}  ->  {placeholder.ljust(placeholder_width)}  ->  {value.ljust(value_width)}  ->  {str(count).ljust(count_width)}")

            if verbose and placeholder in placeholder_line_counts:
                for line_num, line_count in placeholder_line_counts[placeholder].items():
                    print(f"    ⮡  {line_count} Time{'s' if line_count > 1 else ''} at Line {line_num}")
                print("")

    # Collect missing placeholders in the error log
    for placeholder in missing_placeholders:
        error_messages.append(f"❌ ERROR: Placeholder {placeholder} not found in the document.")

    total_changes = sum(placeholder_counts.values())
    # Log expected count mismatch if the field exists
    expected_count = config.get("expectedCount")
    if expected_count is not None and total_changes != expected_count:
        error_messages.append(f"❌ ERROR: Total placeholders replaced ({total_changes}) does not match expected count ({expected_count}).")

    # Log placeholders
    print(f"\n✅ Total placeholder values changed: {total_changes}.")

def replace_text_preserving_format(paragraph, replacements, placeholder_counts, placeholder_line_counts, line_num, missing_placeholders):
    """ Replaces placeholders in a paragraph while preserving formatting. """
    for run in paragraph.runs:
        for key, value in replacements.items():
            if key in run.text:
                run.text = run.text.replace(key, value)
                placeholder_counts[key] += 1
                placeholder_line_counts[key][line_num] = placeholder_line_counts[key].get(line_num, 0) + 1
                missing_placeholders.discard(key)

def replace_placeholders(config, data, error_messages, verbose=False):
    """ Replaces placeholders in a .docx template while preserving formatting and logging occurrences. """
    template_path = os.path.abspath(config["templatePath"])
    output_path = os.path.abspath(config["outputPath"])

    valid, error_message = validate_docx_file(template_path)
    if not valid:
        error_messages.append(error_message)
        return

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    try:
        doc = docx.Document(template_path)
    except Exception as e:
        error_messages.append(f"❌ Unable to open template file '{template_path}'.\n{e}")
        return

    missing_columns = [key for key in config["mappings"] if key not in data]
    for col in missing_columns:
        error_messages.append(f"⚠️  Expected column '{col}' based on config but it was not found in the input data.")

    replacements = {config["mappings"][key]: value for key, value in data.items() if key in config["mappings"]}
    placeholder_counts = {placeholder: 0 for placeholder in replacements}
    placeholder_line_counts = {placeholder: {} for placeholder in replacements}
    missing_placeholders = set(replacements.keys())

    # Replace in paragraphs
    for line_num, para in enumerate(doc.paragraphs, start=1):
        replace_text_preserving_format(para, replacements, placeholder_counts, placeholder_line_counts, line_num, missing_placeholders)

    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for line_num, para in enumerate(cell.paragraphs, start=1):
                    replace_text_preserving_format(para, replacements, placeholder_counts, placeholder_line_counts, line_num, missing_placeholders)

    # Log mappings and save document
    log_mappings(data, config, placeholder_counts, placeholder_line_counts, missing_placeholders, error_messages, verbose)

    try:
        doc.save(output_path)
        print(f"✅ Document successfully saved at: {output_path}")
    except Exception as e:
        error_messages.append(f"❌ Failed to save document to '{output_path}'.\n{e}")


def main():
    parser = argparse.ArgumentParser()

    # Arguent Details
    parser.add_argument("-v", "--verbose", action="store_true", help="Show detailed placeholder changes with line numbers and counts")
    parser.add_argument("-lp","--loop", action="store_true", help="Keep running after completion, allowing re-selection")
    
    args = parser.parse_args()

    data_file = 'input.txt'
    config_dir = "config"
    config_info = list_config_files(config_dir)

    if not config_info:
        print("❌ No valid config files found. Exiting.")
        return

    while True:
        config_file = get_user_selected_config(config_info, config_dir)
        if not config_file:
            return  # Exit if the user chooses to quit


        data, data_errors = read_data(data_file)
        config, config_errors = read_config(config_file)

        error_messages = data_errors + config_errors

        if data and config:
            replace_placeholders(config, data, error_messages, args.verbose)

        if error_messages:
            print("\n📌 **Errors Encountered:** ")
            for error in error_messages:
                print(error)
        
        if not args.loop:
            return

        choice = input("\nPress ENTER to continue... (or 'q' to quit) ")
        if choice.lower() == 'q':
            print("\n🛑 Exiting selection...")
            return

if __name__ == "__main__":
    main()