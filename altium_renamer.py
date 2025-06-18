import os
import re
import sys

def get_project_file(directory):
    """
    Finds the Altium project file (.PrjPcb, .PrjHar, .PrjMbd) in the given directory.
    Ensures there is exactly one project file.
    """
    project_extensions = ('.PrjPcb', '.PrjHar', '.PrjMbd')
    found_project_files = []

    for filename in os.listdir(directory):
        if filename.endswith(project_extensions):
            found_project_files.append(os.path.abspath(os.path.join(directory, filename)))

    if len(found_project_files) == 0:
        return None, "Error: No Altium project file (.PrjPcb, .PrjHar, or .PrjMbd) found in the directory."
    elif len(found_project_files) > 1:
        return None, f"Error: Multiple Altium project files found in the directory. Please ensure only one exists:\n{', '.join([os.path.basename(f) for f in found_project_files])}"
    else:
        return found_project_files[0], None

def extract_parameters_from_project_file(project_file_path):
    """
    Extracts parameter Name and Value from an Altium project file.
    Looks for blocks like:
    [ParameterX]
    Name=Abbreviation
    Value=LPIO
    """
    parameters = {}
    current_parameter_name = None

    try:
        with open(project_file_path, 'r') as f:
            for line in f:
                line = line.strip()

                if re.match(r'\[Parameter\d+\]', line):
                    current_parameter_name = None
                    continue

                name_match = re.match(r'Name=(.+)', line)
                if name_match:
                    current_parameter_name = name_match.group(1).strip()
                    continue

                value_match = re.match(r'Value=(.+)', line)
                if value_match and current_parameter_name:
                    parameters[current_parameter_name] = value_match.group(1).strip()
                    current_parameter_name = None
                    continue

    except FileNotFoundError:
        print(f"Error: Project file not found at '{project_file_path}' during parameter extraction. This should not happen if previous checks passed.")
    except Exception as e:
        print(f"Error reading parameters from project file '{project_file_path}': {e}")

    return parameters

def get_placeholder_parameters_from_filename(filename):
    """
    Extracts all parameter names from a filename that are enclosed in square brackets.
    E.g., "[PCBANumber]_[Abbreviation]_ASSY.PCBDwf" -> ["PCBANumber", "Abbreviation"]
    """
    return re.findall(r'\[(.*?)\]', filename)

def generate_new_filename(old_filename_template, parameters):
    """
    Generates a new filename by replacing bracketed placeholders with their values.
    Returns the new filename and a list of missing parameters.
    """
    missing_params = []
    new_filename = old_filename_template

    placeholders = re.findall(r'(\[.*?\])', old_filename_template)

    for placeholder in placeholders:
        param_name = placeholder[1:-1]
        param_value = parameters.get(param_name)

        if param_value is not None:
            new_filename = new_filename.replace(placeholder, param_value)
        else:
            missing_params.append(param_name)
    
    return new_filename, missing_params

def rename_files_and_update_project(directory):
    """
    Renames files in the directory that match a pattern like [Parameter]_...
    and updates the project file based on parameters extracted from the project file itself.
    """
    # 1. Get the ABSOLUTE path to the project file
    project_file_path, error_message = get_project_file(directory)
    if error_message:
        print(error_message)
        return

    original_project_file_basename = os.path.basename(project_file_path)
    
    # Extract parameters from the project file to get values for its own name if it's templated
    project_params_for_name = extract_parameters_from_project_file(project_file_path)

    project_file_new_basename, missing_project_params = generate_new_filename(
        original_project_file_basename,
        project_params_for_name
    )

    new_project_file_absolute_path = project_file_path # Initialize with current path
    
    if project_file_new_basename != original_project_file_basename:
        # The project file itself needs renaming!
        if missing_project_params:
            print(f"Error: Project file name '{original_project_file_basename}' has missing parameters: {', '.join(missing_project_params)}. Cannot rename project file. Aborting.")
            return
        
        new_project_file_absolute_path = os.path.join(os.path.dirname(project_file_path), project_file_new_basename)
        
        if os.path.exists(new_project_file_absolute_path) and new_project_file_absolute_path != project_file_path:
            print(f"Warning: New project file name '{os.path.basename(new_project_file_absolute_path)}' already exists. Cannot rename project file. Aborting.")
            return
            
        try:
            os.rename(project_file_path, new_project_file_absolute_path)
            print(f"Renamed project file from '{original_project_file_basename}' to '{project_file_new_basename}'.")
            # Update project_file_path to the new path so all subsequent operations use it
            project_file_path = new_project_file_absolute_path
        except OSError as e:
            print(f"Error renaming project file from '{original_project_file_basename}': {e}. Aborting.")
            return

    print(f"Operating on Altium project file: {os.path.basename(project_file_path)}")

    # --- Step 0: Extract all parameters from the project file (now using its potentially new name) ---
    print("\n--- Extracting parameters from project file ---")
    all_extracted_parameters = extract_parameters_from_project_file(project_file_path)

    if not all_extracted_parameters:
        print("Warning: No parameters found in the project file. No files will be renamed based on parameters.")

    # This map will store the ORIGINAL FILENAME AS FOUND ON DISK (which might be the template name)
    # to the NEW ACTUAL FILENAME AFTER RENAMING.
    # E.g., {"[PCBANumber]_[Abbreviation]_ASSY.PCBDwf": "12345_LPIO_ASSY.PCBDwf"}
    # This map will be used to update the DocumentPath references.
    files_renamed_map = {} 

    # --- Step 1: Identify and rename other files ---
    print("\n--- Identifying and renaming other files ---")
    
    current_files_on_disk = os.listdir(directory)
    
    for filename_on_disk in current_files_on_disk:
        full_path_on_disk = os.path.join(directory, filename_on_disk)
        
        if not os.path.isfile(full_path_on_disk):
            continue # Skip directories

        # Skip the project file itself here, as its renaming was handled above
        if full_path_on_disk == project_file_path or \
           (project_file_new_basename != original_project_file_basename and filename_on_disk == original_project_file_basename):
            # This handles the case where project file was renamed (skip new name)
            # or was not renamed but is still the original project file (skip original name)
            continue

        placeholders_in_filename = get_placeholder_parameters_from_filename(filename_on_disk)

        if not placeholders_in_filename:
            continue # Not a file we need to process based on parameters

        new_filename_candidate, missing_params = generate_new_filename(filename_on_disk, all_extracted_parameters)

        if missing_params:
            print(f"Warning: Skipping '{filename_on_disk}'. Missing values for parameters: {', '.join(missing_params)} in project file.")
            continue

        if new_filename_candidate != filename_on_disk:
            old_filepath = full_path_on_disk
            new_filepath = os.path.join(directory, new_filename_candidate)

            if os.path.exists(new_filepath) and new_filepath != old_filepath:
                print(f"Warning: New file '{os.path.basename(new_filepath)}' already exists. Skipping rename for '{filename_on_disk}'.")
                continue

            try:
                os.rename(old_filepath, new_filepath)
                print(f"Renamed '{filename_on_disk}' to '{os.path.basename(new_filepath)}'")
                
                # Store the original filename (which might be the template name)
                # and the new actual name for updating references in the project file.
                files_renamed_map[filename_on_disk] = os.path.basename(new_filepath)
            except OSError as e:
                print(f"Error renaming '{filename_on_disk}': {e}")


    # --- Step 2: Update the project file content (DocumentPath references) ---
    print("\n--- Updating project file content ---")
    
    # Add the project file's own name reference to the map if it was changed.
    # This ensures internal references to itself (if any) are also updated.
    if original_project_file_basename != project_file_new_basename:
        files_renamed_map[original_project_file_basename] = project_file_new_basename

    if not files_renamed_map:
        print("No files were renamed or had their templates resolved, so no DocumentPath updates are necessary.")
        return

    try:
        with open(project_file_path, 'r') as f:
            lines = f.readlines() # Read line by line to modify specific lines

        updated_lines = []
        changes_made_in_content = False

        for line in lines:
            modified_line = line
            # Check if the line starts with 'DocumentPath='
            if line.strip().startswith('DocumentPath='):
                # Extract the path portion (everything after 'DocumentPath=')
                path_part = line.strip().split('=', 1)[1]
                
                # Extract just the filename from the path
                current_filename_in_path = os.path.basename(path_part)
                
                # Look for this filename in our map of renamed files
                new_filename_for_path = files_renamed_map.get(current_filename_in_path)

                if new_filename_for_path and new_filename_for_path != current_filename_in_path:
                    # Replace the old filename with the new filename in the path part
                    # Ensure we preserve the directory structure if any (though usually just filename in PrjXXX)
                    # For simplicity, assuming DocumentPath=filename.ext or DocumentPath=path\filename.ext
                    
                    # More robust replacement using regex to preserve surrounding path
                    # This captures the DocumentPath= prefix and any path before the filename.
                    # It matches the old filename and replaces only that part.
                    # This regex handles cases like:
                    #   DocumentPath=filename.ext
                    #   DocumentPath=SubFolder\filename.ext
                    #   DocumentPath=..\SubFolder\filename.ext
                    
                    # Escape the current_filename_in_path for safe regex use
                    pattern_filename_escaped = re.escape(current_filename_in_path)
                    
                    # Pattern to find 'DocumentPath=' followed by any characters (non-greedy)
                    # then the old filename, then the rest of the line.
                    # We capture the parts we want to keep.
                    
                    # This needs to be carefully constructed. It should find
                    # "DocumentPath=" and then the OLD filename.
                    # Let's rebuild the regex for the replacement.
                    
                    # The value of `path_part` is often just the filename itself,
                    # but it could contain a relative path.
                    # The most straightforward way is to replace `current_filename_in_path`
                    # with `new_filename_for_path` *within* the `path_part` string,
                    # then reconstruct the line.
                    
                    new_path_part = path_part.replace(current_filename_in_path, new_filename_for_path)
                    modified_line = f"DocumentPath={new_path_part}\n" # Reconstruct the line
                    
                    # Only mark changes if the line actually changed
                    if modified_line.strip() != line.strip():
                        changes_made_in_content = True
                        print(f"  Updated DocumentPath: '{line.strip()}' -> '{modified_line.strip()}'")
            
            updated_lines.append(modified_line)

        if changes_made_in_content:
            with open(project_file_path, 'w') as f:
                f.writelines(updated_lines)
            print(f"Successfully wrote updated content to '{os.path.basename(project_file_path)}'.")
        else:
            print(f"No DocumentPath changes were needed in '{os.path.basename(project_file_path)}'.")

    except Exception as e:
        print(f"Error updating project file content '{os.path.basename(project_file_path)}': {e}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("This script is designed to be run from the Windows Explorer context menu.")
        print("Please right-click on the desired project directory and select the 'Rename Altium Project Files' option.")
        # Optionally, you could exit here:
        # sys.exit(1)
    else:
        # sys.argv[1] will be the path to the directory that was right-clicked
        target_directory = sys.argv[1]
        
        if not os.path.isdir(target_directory):
            print(f"Error: Directory '{target_directory}' not found or is not a valid directory.")
            # sys.exit(1)
        else:
            # Ensure target_directory is always an absolute path
            absolute_target_directory = os.path.abspath(target_directory)
            rename_files_and_update_project(absolute_target_directory)
