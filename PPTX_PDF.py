import os
import sys
import comtypes.client

def convert_folder_to_pdf(input_folder, output_folder):
    # 1. Normalize paths to absolute paths
    input_folder = os.path.abspath(input_folder)
    output_folder = os.path.abspath(output_folder)

    # 2. Check if input exists
    if not os.path.exists(input_folder):
        print(f"Error: Input folder not found: {input_folder}")
        return

    # 3. Create output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Created output folder: {output_folder}")

    # 4. Initialize PowerPoint Application
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1  # PowerPoint must be visible for COM to work occasionally

    print(f"Scanning '{input_folder}' for .pptx files...")
    
    count = 0
    
    try:
        # 5. Iterate over files
        for filename in os.listdir(input_folder):
            if filename.lower().endswith(".pptx"):
                input_path = os.path.join(input_folder, filename)
                
                # Create output filename (replace .pptx with .pdf)
                output_filename = os.path.splitext(filename)[0] + ".pdf"
                output_path = os.path.join(output_folder, output_filename)

                print(f"Converting: {filename} -> {output_filename}")

                # 6. Open Presentation
                deck = powerpoint.Presentations.Open(input_path)
                
                # 7. Save as PDF (32 is the format type for PDF)
                deck.SaveAs(output_path, 32)
                
                # 8. Close the specific presentation
                deck.Close()
                count += 1
                
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # 9. Quit PowerPoint application to free resources
        powerpoint.Quit()
        print(f"Done. Converted {count} files.")

if __name__ == "__main__":
    # --- CONFIGURATION ---
    # You can change these paths to whatever you need
    SOURCE_DIR = "./project_pptx"        # Folder containing your .pptx files
    DEST_DIR = "./project"            # Folder where PDFs will be saved
    # ---------------------

    convert_folder_to_pdf(SOURCE_DIR, DEST_DIR)