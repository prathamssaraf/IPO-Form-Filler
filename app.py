from flask import Flask, request, render_template
from PIL import Image
from pdf2image import convert_from_path
from flask import send_file



from form_filler_functions import printName, printAddress, printNumber, printEmail, printPAN, \
    printCDSL, printQuantity, printAmount, printWords, printAccNumber, printPrice, printBankName

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    n = int(request.form['number'])
    convert_and_overlay_images(n)
    pdf_path = "output.pdf"
    return send_file(pdf_path, as_attachment=True)

def convert_and_overlay_images(n):
    for i in range(0, n+1):
        pdf_file = f'form ({i+1}).pdf'
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
        output_image_path = printBankName(i, output_image_path)
        print("Output image saved as:", output_image_path.format(i=i))
    images = [
        Image.open(f)
        for f in ["form_{}.png".format(i) for i in range(n)]
    ]
    pdf_path = "output.pdf"
    images[0].save(pdf_path, "PDF", resolution=100.0, save_all=True, append_images=images[1:])


if __name__ == '__main__':
    app.run(debug=True)
