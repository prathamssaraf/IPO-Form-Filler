from PIL import Image, ImageDraw, ImageFont
import pandas as pd

def printName(i, image_path):
    # Open the image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path)

    # Get the value from the specified cell
    name = str(df.loc[i, 'Mr./ Ms.'])

    # Split the name if it exceeds 15 characters
    if len(name) > 15:
        text1 = name[:15]
        text2 = name[15:]
    else:
        text1 = name
        text2 = ""

    # Specify the font and font size
    font = ImageFont.truetype('assets/arial.ttf', 35)  # Change 'assets/arial.ttf' to the path of your desired font file

    # Specify the position to write the text
    text_position = (953, 275)  # Change the coordinates to the desired position

    # Specify the text color
    text_color = (0, 0, 0)  # Black color (RGB values)

    # Calculate the spacing between characters
    letter_spacing = 10  # Adjust the spacing value to increase or decrease the spacing between characters

    # Write text1 on the image
    x, y = text_position
    for char in text1:
        draw.text((x, y), char, font=font, fill=text_color)
        char_width, _ = draw.textsize(char, font=font)
        x += char_width + 2.4 * letter_spacing
    
    
    text_position = (865, 315)
    x, y = text_position
    for char in text2:
        draw.text((x, y), char, font=font, fill=text_color)
        char_width, _ = draw.textsize(char, font=font)
        x += char_width + 2.4 * letter_spacing


   
    # Specify the font and position for text2
    font = ImageFont.truetype('assets/arial.ttf', 20)
    text_position = (260, 1838)
    x, y = text_position
    draw.text((x, y), name, font=font, fill=text_color)

    font = ImageFont.truetype('assets/arial.ttf', 24)
    text_position = (1130, 1985)
    x, y = text_position
    draw.text((x, y), name, font=font, fill=text_color)

    # Save the modified image
    output_path = f'form_{i}.png'
    image.save(output_path)

    # Close the image
    image.close()

    return output_path


def printEmail(i, image_path):
    # Open the image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path)

    # Get the value from the specified cell
    email = df.loc[i, 'Email']

    # Specify the font and font size
    font = ImageFont.truetype('assets/arial.ttf', 35)  # Change 'assets/arial.ttf' to the path of your desired font file

    # Specify the text color
    text_color = (0, 0, 0)  # Black color (RGB values)

    # Calculate the spacing between characters
    letter_spacing = 10

    font = ImageFont.truetype('assets/arial.ttf', 22)
    text_position = (1230, 385)
    x, y = text_position
    draw.text((x, y), email, font=font, fill=text_color)

    acc = str(int(df.loc[i, 'Account No.']))
    text_position = (700, 1875)
    x, y = text_position
    draw.text((x, y), email, font=font, fill=text_color)

    # Save the modified image
    output_path = image_path
    image.save(output_path)

    # Close the image
    image.close()

    return output_path


def printAddress(i, image_path):
    # Open the image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path)

    # Get the value from the specified cell
    address = df.loc[i, 'Address']

    # Specify the font and font size
    font = ImageFont.truetype('assets/arial.ttf', 20)  # Change 'assets/arial.ttf' to the path of your desired font file

    # Specify the position to write the text
    text_position = (950, 355)  # Change the coordinates to the desired position

    # Specify the text color
    text_color = (0, 0, 0)  # Black color (RGB values)

    # Calculate the spacing between characters
    letter_spacing = 10  # Adjust the spacing value to increase or decrease the spacing between characters

    # Write the address on the image
    draw.text(text_position, address, font=font, fill=text_color)

    # Save the modified image
    output_path = image_path
    image.save(output_path)

    # Close the image
    image.close()

    return output_path

def printNumber(i, image_path):
    # Open the image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path)

    # Get the value from the specified cell
    number = str(int(df.loc[i, 'Tel. No(with STD code)/Mobile']))

    # Specify the font and font size
    font = ImageFont.truetype('assets/arial.ttf', 35)  # Change 'assets/arial.ttf' to the path of your desired font file

    # Specify the position to write the text
    text_position = (1180, 415)  # Change the coordinates to the desired position

    # Specify the text color
    text_color = (0, 0, 0)  # Black color (RGB values)

    # Calculate the spacing between characters
    letter_spacing = 10  # Adjust the spacing value to increase or decrease the spacing between characters

    # Iterate over each character in the text
    x, y = text_position
    for char in number:
        # Write the character on the image
        draw.text((x, y), char, font=font, fill=text_color)
        
        # Calculate the width of the character
        char_width, _ = draw.textsize(char, font=font)
        
        # Update the x-coordinate for the next character, considering the spacing
        x += char_width + 2.55*letter_spacing

    # Save the modified image
    output_path = image_path
    image.save(output_path)

    # Close the image
    image.close()

    return output_path


def printPAN(i, image_path):
    # Open the image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path)

    # Get the value from the specified cell
    pan_number = str(df.loc[i, 'PAN OF SOLE/FIRST BIDDER'])

    # Specify the font and font size
    font = ImageFont.truetype('assets/arial.ttf', 45)  # Change 'assets/arial.ttf' to the path of your desired font file

    # Specify the position to write the text
    text_position = (848, 505)  # Change the coordinates to the desired position

    # Specify the text color
    text_color = (0, 0, 0)  # Black color (RGB values)

    # Calculate the spacing between characters
    letter_spacing = 10  # Adjust the spacing value to increase or decrease the spacing between characters

    # Iterate over each character in the text
    x, y = text_position
    for char in pan_number:
        # Write the character on the image
        draw.text((x, y), char, font=font, fill=text_color)
        
        # Calculate the width of the character
        char_width, _ = draw.textsize(char, font=font)
        
        # Update the x-coordinate for the next character, considering the spacing
        x += char_width + 5.2*letter_spacing

    # Specify the position and iterate again for the second set of PAN number
    text_position = (1090, 1655)
    x, y = text_position
    for char in pan_number:
        # Write the character on the image
        draw.text((x, y), char, font=font, fill=text_color)
        
        # Calculate the width of the character
        char_width, _ = draw.textsize(char, font=font)
        
        # Update the x-coordinate for the next character, considering the spacing
        x += char_width + 3*letter_spacing

    # Save the modified image
    output_path = image_path
    image.save(output_path)

    # Close the image
    image.close()

    return output_path


def printCDSL(i, image_path):
    # Open the image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path)

    # Get the value from the specified cell
    cdsl_number = str(int(df.loc[i, 'CDSL']))

    # Specify the font and font size
    font = ImageFont.truetype('assets/arial.ttf', 50)  # Change 'assets/arial.ttf' to the path of your desired font file

    # Specify the position to write the text
    text_position = (40, 600)  # Change the coordinates to the desired position

    # Specify the text color
    text_color = (0, 0, 0)  # Black color (RGB values)

    # Calculate the spacing between characters
    letter_spacing = 10  # Adjust the spacing value to increase or decrease the spacing between characters

    # Iterate over each character in the text
    x, y = text_position
    for char in cdsl_number:
        # Write the character on the image
        draw.text((x, y), char, font=font, fill=text_color)
        
        # Calculate the width of the character
        char_width, _ = draw.textsize(char, font=font)
        
        # Update the x-coordinate for the next character, considering the spacing
        x += char_width + 5*letter_spacing

    # Specify the position and iterate again for the second set of CDSL number
    text_position = (100, 1655) 
    x, y = text_position
    for char in cdsl_number:
        # Write the character on the image
        draw.text((x, y), char, font=font, fill=text_color)
        
        # Calculate the width of the character
        char_width, _ = draw.textsize(char, font=font)
        
        # Update the x-coordinate for the next character, considering the spacing
        x += char_width + 3*letter_spacing

    # Save the modified image
    output_path = image_path
    image.save(output_path)

    # Close the image
    image.close()

    return output_path


def printAccNumber(i, image_path):
    # Open the image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path, dtype={'Account No.': str})

    # Get the value from the specified cell
    account_number = (df.loc[i, 'Account No.'])

    # Specify the font and font sizeac
    font = ImageFont.truetype('assets/arial.ttf', 28)  # Change 'assets/arial.ttf' to the path of your desired font file

    # Specify the position to write the text
    text_position = (160, 1059)  # Change the coordinates to the desired position

    # Specify the text color
    text_color = (0, 0, 0)  # Black color (RGB values)

    # Calculate the spacing between characters
    letter_spacing = 10  # Adjust the spacing value to increase or decrease the spacing between characters

    # Iterate over each character in the text
    x, y = text_position
    for char in account_number:
        # Write the character on the image
        draw.text((x, y), char, font=font, fill=text_color)
        
        # Calculate the width of the character
        char_width, _ = draw.textsize(char, font=font)
        
        # Update the x-coordinate for the next character, considering the spacing
        x += char_width + 2.5*letter_spacing

    # Specify the font and position for the additional account number
    font = ImageFont.truetype('assets/arial.ttf', 25)
    text_position = (700, 1750)
    x, y = text_position
    draw.text((x, y), account_number, font=font, fill=text_color)

    # Specify the font and position for the second additional account number
    font = ImageFont.truetype('assets/arial.ttf', 22)
    text_position = (350, 2115)
    x, y = text_position
    draw.text((x, y), account_number, font=font, fill=text_color)

    # Save the modified image
    output_path = image_path
    image.save(output_path)

    # Close the image
    image.close()

    return output_path


def printBankName(i, image_path):
    # Open the image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path)

    # Get the value from the specified cell
    bank_name = df.loc[i, 'Bank Name']

    # Specify the font and font size
    font = ImageFont.truetype('assets/arial.ttf', 25)  # Change 'assets/arial.ttf' to the path of your desired font file

    # Specify the position to write the text
    text_position = (260, 1100)  # Change the coordinates to the desired position

    # Specify the text color
    text_color = (0, 0, 0)  # Black color (RGB values)

    # Calculate the spacing between characters
    letter_spacing = 10  # Adjust the spacing value to increase or decrease the spacing between characters

    # Iterate over each character in the text
    x, y = text_position

    draw.text((x, y), bank_name, font=font, fill=text_color)

    text_position = (260, 1795)
    x, y = text_position
    draw.text((x, y), bank_name, font=font, fill=text_color)

    font = ImageFont.truetype('assets/arial.ttf', 22)
    text_position = (300, 2150)
    x, y = text_position
    draw.text((x, y), bank_name, font=font, fill=text_color)

    # Save the modified image
    output_path = image_path
    image.save(output_path)

    # Close the image
    image.close()

    return output_path


def printWords(i, image_path):
    # Open the image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path)

    # Get the value from the specified cell
    words = str(df.loc[i, 'Words'])

    # Specify the font and font size
    font = ImageFont.truetype('assets/arial.ttf', 25)  # Change 'assets/arial.ttf' to the path of your desired font file

    # Specify the position to write the text
    text_position = (800, 1010)  # Change the coordinates to the desired position

    # Specify the text color
    text_color = (0, 0, 0)  # Black color (RGB values)

    # Calculate the spacing between characters
    letter_spacing = 10  # Adjust the spacing value to increase or decrease the spacing between characters

    x, y = text_position
    draw.text((x, y), words, font=font, fill=text_color)

    # Save the modified image
    output_path = image_path
    image.save(output_path)

    # Close the image
    image.close()

    return output_path


def printTick(i, image_path):
    # Open the main image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path, dtype={'Amount': str})

    # Get the value from the specified cell
    amount = (df.loc[i, 'Amount'])

    amount = int(amount)

    
    tick_position = (1083, 860)
        
    # Load the tick image
    tick_image = Image.open('assets/tick.png')

    # Resize the tick image if necessary
    tick_size = (20, 20)  # Change the size of the tick image as needed
    tick_image = tick_image.resize(tick_size)

    # Paste the tick image onto the main image at the specified position
    if amount < 200000:
        image.paste(tick_image, tick_position, tick_image)
        image.paste(tick_image, tick_position, tick_image)
        image.paste(tick_image, tick_position, tick_image)
        image.paste(tick_image, tick_position, tick_image)
    
    # Save the modified image
    output_path = image_path
    image.save(output_path)

    # Close the images
    image.close()
    tick_image.close()

    return output_path

def printTickcategory(i, image_path):
    # Open the main image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path, dtype={'Amount': str})

    # Get the value from the specified cell
    amount = (df.loc[i, 'Amount'])

    amount = int(amount)

    
    tick_position = (1083, 860)


    if amount < 200000:
          tick_position = (1148, 765)
    else:
        tick_position = (1148, 855)
    # Load the tick image
    tick_image = Image.open('assets/tick.png')

    # Resize the tick image if necessary
    tick_size = (20, 20)  # Change the size of the tick image as needed
    tick_image = tick_image.resize(tick_size)

    # Paste the tick image onto the main image at the specified position
    
    image.paste(tick_image, tick_position, tick_image)
    image.paste(tick_image, tick_position, tick_image)
    image.paste(tick_image, tick_position, tick_image)
    image.paste(tick_image, tick_position, tick_image)

    tick_position = (716, 573)
    image.paste(tick_image, tick_position, tick_image)
    image.paste(tick_image, tick_position, tick_image)
    image.paste(tick_image, tick_position, tick_image)
    image.paste(tick_image, tick_position, tick_image)

    # Save the modified image
    output_path = image_path
    image.save(output_path)
   
    # Close the images
    image.close()
    tick_image.close()

    return output_path

def printTickcategory_with_huf_check(i, image_path):
    # Open the main image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path)

    # Get the value from the specified cell
    name = str(df.loc[i, 'Mr./ Ms.'])

    tick_position = (1083, 860)

    # Check if the last three letters of the name are "HUF"
    if name[-3:].upper() == "HUF":
        tick_position = (1310, 630)
    else:
        tick_position = (1310, 600)

    # Load the tick image
    tick_image = Image.open('assets/tick.png')

    # Resize the tick image if necessary
    tick_size = (20, 20)  # Change the size of the tick image as needed
    tick_image = tick_image.resize(tick_size)

    # Paste the tick image onto the main image at the specified position
    image.paste(tick_image, tick_position, tick_image)
    image.paste(tick_image, tick_position, tick_image)
    image.paste(tick_image, tick_position, tick_image)
    image.paste(tick_image, tick_position, tick_image)

    # Save the modified image
    output_path = image_path
    image.save(output_path)

    # Close the images
    image.close()
    tick_image.close()

    return output_path

def printAmount(i, image_path):
    # Open the image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file

    df = pd.read_excel(excel_path, dtype={'Amount': str})

    # Get the value from the specified cell
    amount = (df.loc[i, 'Amount'])
    # Specify the font and font size
    font = ImageFont.truetype('assets/arial.ttf', 25)  # Change 'assets/arial.ttf' to the path of your desired font file

    # Specify the position to write the text
    text_position = (320, 1010)  # Change the coordinates to the desired position

    # Specify the text color
    text_color = (0, 0, 0)  # Black color (RGB values)

    # Calculate the spacing between characters
    letter_spacing = 10  # Adjust the spacing value to increase or decrease the spacing between characters

    # Iterate over each character in the text
    x, y = text_position
    for char in amount:
        # Write the character on the image
        draw.text((x, y), char, font=font, fill=text_color)

        # Calculate the width of the character
        char_width, _ = draw.textsize(char, font=font)

        # Update the x-coordinate for the next character, considering the spacing
        x += char_width + 2.2 * letter_spacing

    font = ImageFont.truetype('assets/arial.ttf', 24)
    text_position = (300, 1750)
    x, y = text_position
    draw.text((x, y), "₹ " + amount, font=font, fill=text_color)

    font = ImageFont.truetype('assets/arial.ttf', 24)
    text_position = (370, 2070)
    x, y = text_position
    draw.text((x, y), "₹ " + amount, font=font, fill=text_color)

    # Save the modified image
    output_path = image_path
    image.save(output_path)

    # Close the image
    image.close()

    return output_path


def printQuantity(i, image_path):
    # Open the image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path)

    # Get the value from the specified cell
    quantity = str(int(df.loc[i, 'Quantity']))

    # Specify the font and font size
    font = ImageFont.truetype('assets/arial.ttf', 30)  # Change 'assets/arial.ttf' to the path of your desired font file

    # Specify the position to write the text
      # Change the coordinates to the desired position
    num_units = len(quantity)
    if num_units == 1:
        text_position = (520, 850)
    elif num_units == 2:
        text_position = (470, 850)  # Adjust for 2 units
    elif num_units == 3:
        text_position = (420, 850)  # Adjust for 3 units
    elif num_units == 4:
        text_position = (370, 850)
    elif num_units == 5:
        text_position = (320, 850)


    # Specify the text color
    text_color = (0, 0, 0)  # Black color (RGB values)

    # Calculate the spacing between characters
    letter_spacing = 10  # Adjust the spacing value to increase or decrease the spacing between characters

    # Calculate the total width of the text
    

    # Iterate over each character in the text
    x, y = text_position
    for char in quantity:
        # Write the character on the image
        draw.text((x, y), char, font=font, fill=text_color)

        # Calculate the width of the character
        char_width, _ = draw.textsize(char, font=font)

        # Update the x-coordinate for the next character, considering the spacing
        x += char_width + 3*letter_spacing

    quantity = str(int(df.loc[i, 'Quantity']))
    font = ImageFont.truetype('assets/arial.ttf', 24)
    text_position = (300, 1997)
    x, y = text_position
    draw.text((x, y), quantity, font=font, fill=text_color)

    # Save the modified image
    output_path = image_path
    image.save(output_path)

    # Close the image
    image.close()

    return output_path


def printPrice(i, image_path):
    # Open the image
    image = Image.open(image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(image)

    excel_path = 'formdata.xlsx'  # Path to the Excel file
    df = pd.read_excel(excel_path)

    # Get the value from the specified cell
    price = str(int(df.loc[i, 'Price']))

    # Specify the font and font size
    font = ImageFont.truetype('assets/arial.ttf', 40)  # Change 'assets/arial.ttf' to the path of your desired font file

    # Specify the text color
    text_color = (0, 0, 0)  # Black color (RGB values)

    font = ImageFont.truetype('assets/arial.ttf', 24)
    text_position = (570, 860)
    x, y = text_position
    draw.text((x, y), "₹" + price, font=font, fill=text_color)

    font = ImageFont.truetype('assets/arial.ttf', 24)
    text_position = (300, 2035)
    x, y = text_position
    draw.text((x, y), "₹" + price, font=font, fill=text_color)

    # Save the modified image
    output_path = image_path
    image.save(output_path)

    # Close the image
    image.close()

    return output_path


