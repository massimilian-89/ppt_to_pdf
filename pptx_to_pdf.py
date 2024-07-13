import os
import argparse
import comtypes.client

def pptx_to_pdf(input_file, output_file):
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    presentation = powerpoint.Presentations.Open(input_file)
    presentation.SaveAs(output_file, 32)  # 32 is the PowerPoint constant for pdf format
    presentation.Close()
    powerpoint.Quit()

def convert_single_file(pptx_path):
    pdf_path = pptx_path.replace('.pptx', '.pdf')
    pptx_to_pdf(pptx_path, pdf_path)
    print(f"Converted {pptx_path} to {pdf_path}")

def convert_all_files_in_directory(directory_path):
    for filename in os.listdir(directory_path):
        if filename.endswith('.pptx'):
            pptx_path = os.path.join(directory_path, filename)
            convert_single_file(pptx_path)

def convert_all_files_in_current_directory():
    directory_path = os.getcwd()
    convert_all_files_in_directory(directory_path)

def main():
    parser = argparse.ArgumentParser(description="Convert PPTX files to PDF")
    parser.add_argument("option", type=int, choices=[1, 2, 3], help="1: Convert a single PPTX file, 2: Convert all PPTX files in a directory, 3: Convert all PPTX files in the current directory")
    parser.add_argument("path", type=str, nargs='?', help="Path to the PPTX file or directory (required for options 1 and 2)")

    args = parser.parse_args()

    if args.option == 1:
        if args.path and os.path.isfile(args.path):
            convert_single_file(args.path)
        else:
            print("Invalid file path. Please try again.")
    elif args.option == 2:
        if args.path and os.path.isdir(args.path):
            convert_all_files_in_directory(args.path)
        else:
            print("Invalid directory path. Please try again.")
    elif args.option == 3:
        convert_all_files_in_current_directory()

if __name__ == "__main__":
    main()
