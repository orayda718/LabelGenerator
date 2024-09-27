"""
Label Generator Script
Author: Orayda Shagifa
Description:
This script reads data from an Excel file and generates labels as images. 
Each label contains formatted details such as price, lot number, car name, VIN, miles, and a code. 
The labels are saved in a folder named 'labels_<current date and time>'. 
Additionally, a logo is placed on the bottom-left corner of each label.
Made for A1 Big Star Motors LTD.
"""

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
from datetime import datetime

# Load the spreadsheet containing data
filePath = 'SourceData.xlsx'  # Path to the Excel file
data = pd.read_excel(filePath)

# Replace any NaN values with empty strings
data = data.fillna("")

# Set up font paths and sizes for the label text
fontPath = 'fonts/Roboto/Roboto-Bold.ttf'
fontLarge = ImageFont.truetype(fontPath, 180)  # Large font for price
fontMedium = ImageFont.truetype(fontPath, 85)   # Medium font for other details
fontSmall = ImageFont.truetype(fontPath, 60)    # Small font for code

# Canvas dimensions (size of the label image)
canvasWidth, canvasHeight = 1056, 816

# Load and resize the logo
logoPath = 'logo.png'  # Path to your logo file
logo = Image.open(logoPath)
logo = logo.resize((150, 150))  # Resize the logo to 150x150 pixels

# Get current date and time for the folder name
currentTime = datetime.now().strftime("%Y-%m-%d_%H:%M")

# Create a folder to save the generated labels
outputFolder = f"labels_{currentTime}"
os.makedirs(outputFolder, exist_ok=True)

# Function to generate labels for each row of data
def createLabel(row):
    # Create a blank white canvas for the label
    img = Image.new('RGB', (canvasWidth, canvasHeight), color='white')
    draw = ImageDraw.Draw(img)

    # Function to center text horizontally on the label
    def centerText(drawObj, text, yPos, font):
        bbox = drawObj.textbbox((0, 0), text, font=font)
        textWidth = bbox[2] - bbox[0]  # Calculate text width
        xPos = (canvasWidth - textWidth) // 2  # Center the text
        drawObj.text((xPos, yPos), text, font=font, fill='black')
        return bbox[3] - bbox[1]  # Return text height for next position

    # Function to right-align text (for the code at the bottom right)
    def rightAlignText(drawObj, text, yPos, font):
        bbox = drawObj.textbbox((0, 0), text, font=font)
        textWidth = bbox[2] - bbox[0]  # Calculate text width
        xPos = canvasWidth - textWidth - 30  # Align the text 30px from the right
        drawObj.text((xPos, yPos - bbox[3] + bbox[1] - 30), text, font=font, fill='black')
        return bbox[3] - bbox[1]  # Return text height

    # Initialize the vertical position for text
    verticalPos = 100

    # Function to update the vertical position for centered text
    def addCenteredText(font, text):
        nonlocal verticalPos
        textHeight = centerText(draw, text, verticalPos, font)
        verticalPos += textHeight + 50  # Add padding between lines

    # Add details to the label
    formattedPrice = f"${float(row['Price']):,.0f}"  # Format price with commas
    addCenteredText(fontLarge, formattedPrice)
    addCenteredText(fontMedium, f"Lot {row['Lot']}")
    addCenteredText(fontSmall, f"{row['Car']}")
    addCenteredText(fontMedium, f"{row['Vin']}")
    addCenteredText(fontMedium, f"{row['Miles']}")

    # Add right-aligned code at the bottom right corner
    rightAlignText(draw, f"00{row['Code']}", canvasHeight, fontSmall)

    # Paste the logo at the bottom-left corner
    img.paste(logo, (10, canvasHeight - logo.size[1] - 10))  # 10px margin from the bottom and left

    # Save the label image in the output folder
    img.save(os.path.join(outputFolder, f"label_{row['Vin']}.png"))

# Generate labels for each row
for index, row in data.iterrows():
    createLabel(row) # Pass each row from the DataFrame to the createLabel function
    print("Creating label " + str(index + 2) + "/" + str(data.shape[0] + 1) + "...")

print("Done creating " + str(data.shape[0] + 1) + " labels!")