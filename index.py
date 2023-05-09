import os
import img2pdf
import docx2pdf
import comtypes.client
import PyPDF2



# print the program header
def print_header():
    print("===============================================")
    print("   Welcome to the Ultimate File Converter v1.0  ")
    print("===============================================")
    print()

# convert image to pdf
def image_to_pdf(input_file, output_file):
    with open(output_file, "wb") as f:
        f.write(img2pdf.convert(input_file))
    print("\033[32mConversion complete!\033[0m") # green text

# convert docx to pdf
def docx_to_pdf(input_file, output_file):
    docx2pdf.convert(input_file, output_file)
    print("\033[32mConversion complete!\033[0m") # green text

# convert ppt to pdf
def ppt_to_pdf(input_file, output_file):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    ppt = powerpoint.Presentations.Open(input_file)
    ppt.SaveAs(output_file, FileFormat=32)  # format type for pdf
    ppt.Close()

    powerpoint.Quit()

    print("\033[32mConversion complete!\033[0m") # green text

# main program
def main():
    # print header
    print_header()
    print("Choose the conversion type:")
    print("\033[33m1. Image To Pdf\033[0m") # yellow text
    print("\033[33m2. Word To Pdf\033[0m") # yellow text
    print("\033[33m3. PPT To Pdf\033[0m") # yellow text
    print("\033[33m0. Exit\033[0m") # yellow text

    userChoose = float(input("\033[35mInsert your choice: \033[0m")) # magenta text

    if userChoose == 1:
        print("You selected: Image to PDF")
        print("Please enter the image file name (supported formats: jpg, jpeg, png)")
        if __name__ == "__main__":
            # input
            input_file = input("\033[35mInsert name of the input file: \033[0m")
            input_file_path = os.path.join(".", "input", input_file)

            # output
            output_file = input("\033[35mInsert name of the output file: \033[0m") + ".pdf"
            output_file_path = os.path.join(".", "output", output_file)

            # get current working directory
            cwd = os.getcwd()

            # construct full path to input and output files
            input_file_path = os.path.join(cwd, input_file_path)
            output_file_path = os.path.join(cwd, output_file_path)

            # check if input file is a supported image format
            if os.path.splitext(input_file)[1].lower() in [".jpg", ".jpeg", ".png"]:
                image_to_pdf(input_file_path, output_file_path)
            else:
                print("\033[31mThe format file is not supported.\033[0m") # red text

    elif userChoose == 2:
        print("You selected: Word to PDF")
        print("Please enter the Word file name (supported format: docx)")
        if __name__ == "__main__":
            # input
            input_file = input("\033[35mInput file name: \033[0m") + ".docx"
            input_file_path = os.path.join(".", "input", input_file)

            # output
            output_file = input("\033[35mOutput file name: \033[0m") + ".pdf"
            output_file_path = os.path.join(".", "output", output_file)

            # get current working directory
            cwd = os.getcwd()

            # construct full path to input and output files
            input_file_path = os.path.join(cwd, input_file_path)
            output_file_path = os.path.join(cwd, output_file_path)

            # check if input file is a supported word format
            if os.path.splitext(input_file)[1].lower() in [".docx"]:
                docx_to_pdf(input_file_path, output_file_path)
            else:
                print("\033[31mThe format file is not supported.\033[0m") # red text

    elif userChoose == 3:
        print("You selected: PPT to PDF")
        print("Please enter the PPT file name (supported format: pptx)")
        if __name__ == "__main__":
            # input
            input_file = input("\033[35mInput file name: \033[0m") + ".pptx"
            input_file_path = os.path.join(".", "input", input_file)

            # output
            output_file = input("\033[35mOutput file name: \033[0m") + ".pdf"
            output_file_path = os.path.join(".", "output", output_file)

            # get current working directory
            cwd = os.getcwd()

            # construct full path to input and output files
            input_file_path = os.path.join(cwd, input_file_path)
            output_file_path = os.path.join(cwd, output_file_path)

            # check if input file is a supported PowerPoint format
            if os.path.splitext(input_file)[1].lower() in [".pptx"]:
                ppt_to_pdf(input_file_path, output_file_path)
            else:
                print("\033[31mThe file format is not supported.\033[0m") # red text
    
    
                  
    elif userChoose == 0:
            print("Goodbye!")
            exit()
    
    else:
        print("\033[31mInvalid choice. Please try again.\033[0m") # red text
    
    # return to the main menu
    input("Press enter to return to the main menu...")
    main()
if __name__ == "__main__":
    main()