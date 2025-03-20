import os
import win32com.client
def get_folder_path(prompt):
    while True:
        folder_path = input(prompt).strip().strip('"')  # Remove extra spaces and quotes
        if os.path.isdir(folder_path):
            return folder_path
        else:
            print("Invalid path. Please enter a valid folder path.")
# Ask user for input and output folders
input_folder = get_folder_path("Enter the path of the folder containing PPTX files: ")
output_folder = get_folder_path("Enter the path where PDFs should be saved: ")
if not os.path.exists(output_folder):
    os.makedirs(output_folder)
# Initialize PowerPoint
pptApp = win32com.client.Dispatch("PowerPoint.Application")
pptApp.Visible = 1  # Make visible to avoid permission issues
# Convert all PPTX files
for file in os.listdir(input_folder):
    if file.endswith(".pptx"):
        pptPath = os.path.join(input_folder, file)
        pdfPath = os.path.join(output_folder, file.replace(".pptx", ".pdf"))
        print(f"Converting: {file} → {os.path.basename(pdfPath)}")
        presentation = pptApp.Presentations.Open(pptPath, WithWindow=False)
        presentation.SaveAs(pdfPath, 32)  # 32 is PDF format
        presentation.Close()
pptApp.Quit()
print("All files have been converted successfully!")