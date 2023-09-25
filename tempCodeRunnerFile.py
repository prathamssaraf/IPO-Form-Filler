from flask import Flask, flash, request, render_template, send_file
from PIL import Image
from pdf2image import convert_from_path
import os

from form_filler_functions import printName, printAddress, printNumber, printEmail, printPAN, \
    printCDSL, printQuantity, printAmount, printWords, printAccNumber, printPrice, printBankName,printTick,printTickcategory,printTickcategory_with_huf_check

app = Flask(__name__)

app.secret_key = "prathamsaraf"


@app.route('/')
def home():
    return render_template('index.html')

@app.route('/instructions.html')
def instructions():
    return render_template('instructions.html')


@app.route('/convert', methods=['POST'])

def convert():
    # pdf_files
    # pdf_files.clear()
    # Get the uploaded PDF files from the form
    pdf_files = request.files.getlist('pdf_files')

    excel_file = request.files['excel_file']
    excel_path = "formdata.xlsx"
    excel_file.save(excel_path)

    # Initialize a list to store the paths of converted PDFs
    converted_pdf_paths = []

    # Process each uploaded PDF file
    for i, pdf_file in enumerate(pdf_files):
        # Save the uploaded file to the server
        pdf_path = f"form ({i}).pdf"
        pdf_file.save(pdf_path)

        # Get the number of iterations for this PDF
    n = int(request.form['number'])

        # Convert and overlay images for each iteration
    convert_and_overlay_images(n)
    pdf_path = "output.pdf"
    flash("Conversion completed successfully!")
    pdf_files.clear()
    return send_file(pdf_path, as_attachment=True)


def convert_and_overlay_images(n):
    for i in range(0, n):
        pdf_file = f'form ({i}).pdf'
        images = convert_from_path(pdf_file, dpi=200)
        image_path = f'form_{i}.png'
        images[0].save(image_path, 'PNG')
        output_image_path = image_path
        output_image_path = printName(i, output_image_path)
        output_image_path = printAddress(i, output_image_path)
        output_image_path = printNumber(i, output_image_path)
        output_image_path = printEmail(i, output_image_path)
        output_image_path = printPAN(i, output_image_path)
        output_image_path = printCDSL(i, output_image_path)
        output_image_path = printQuantity(i, output_image_path)
        output_image_path = printAmount(i, output_image_path)
        output_image_path = printWords(i, output_image_path)
        output_image_path = printAccNumber(i, output_image_path)
        output_image_path = printPrice(i, output_image_path)
        output_image_path = printTick(i,output_image_path)
        output_image_path = printBankName(i, output_image_path)
        output_image_path = printTickcategory(i,output_image_path)
        output_image_path = printTickcategory_with_huf_check(i,output_image_path)
        print("Output image saved as:", output_image_path.format(i=i))
    images = [
        Image.open(f)
        for f in ["form_{}.png".format(i) for i in range(n)]
    ]
    pdf_path = "output.pdf"
    images[0].save(pdf_path, "PDF", resolution=100.0, save_all=True, append_images=images[1:])

    for i in range(0, n):
        os.remove(f'form ({i}).pdf')
        os.remove(f'form_{i}.png')
    os.remove("formdata.xlsx")

if __name__ == '__main__':
    app.run(debug=True)
