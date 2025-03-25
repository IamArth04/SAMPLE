import cv2  # For image processing
import pytesseract  # For OCR text extraction
import pandas as pd  # For working with Excel
import os  # For file operations
import glob  # For handling multiple images

# ✅ Set the correct Tesseract OCR path
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ✅ Define the folder containing images and the output Excel file
image_folder = "images/"
excel_file = "service_orders.xlsx"

# ✅ Function to preprocess images (grayscale + adaptive thresholding)
def preprocess_image(image_path):
    image = cv2.imread(image_path)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  # Convert to grayscale
    gray = cv2.GaussianBlur(gray, (5, 5), 0)  # Apply Gaussian blur to reduce noise
    
    # Apply Adaptive Thresholding
    processed = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 2
    )
    return processed

# ✅ Function to extract text from an image
def extract_text_from_image(image_path):
    processed_image = preprocess_image(image_path)  # Preprocess the image
    extracted_text = pytesseract.image_to_string(processed_image, config='--oem 3 --psm 6')

    data_dict = {}  # Dictionary to store key-value pairs
    lines = extracted_text.split("\n")  # Split text into lines

    for line in lines:
        if ":" in line:  # Check for key-value pairs
            key, value = line.split(":", 1)  # Split at first ':'
            data_dict[key.strip()] = value.strip()  # Store in dictionary

    return data_dict  # Return extracted data

# ✅ Process all images in the folder and store data in Excel
def process_images():
    all_data = []  # List to store extracted data

    for img_path in glob.glob(os.path.join(image_folder, "*.jpg")):  # Get all .jpg images
        extracted_data = extract_text_from_image(img_path)  # Extract text
        if extracted_data:
            all_data.append(extracted_data)  # Add data to the list

    if all_data:  # If we have extracted data
        df = pd.DataFrame(all_data)  # Convert list to DataFrame

        if os.path.exists(excel_file):  # Check if Excel file exists
            existing_df = pd.read_excel(excel_file)  # Load existing data
            df = pd.concat([existing_df, df], ignore_index=True)  # Append new data

        df.to_excel(excel_file, index=False)  # Save to Excel
        print(f"✅ {len(all_data)} images processed and saved to '{excel_file}'!")
    else:
        print("❌ No valid data extracted from images.")

# ✅ Function to display the main menu
def bot_menu():
    print("\n📌 Service Order BOT Menu:")
    print("1. Process Service Order Images")
    print("2. View All Records")
    print("3. Search for a Specific Service Order ID")
    print("4. Exit")

# ✅ Function to view all stored records
def view_records():
    if os.path.exists(excel_file):  # Check if the file exists
        df = pd.read_excel(excel_file)  # Load data
        print("\n📄 Service Orders Data:")
        print(df)  # Display records
    else:
        print("❌ No records found. Please process images first.")

# ✅ Function to search for a specific Service Order ID
def search_order():
    if os.path.exists(excel_file):  # Check if the file exists
        df = pd.read_excel(excel_file)  # Load data
        search_id = input("🔍 Enter Service Order ID: ").strip()  # Get user input

        if search_id in df.iloc[:, 0].astype(str).values:  # Check first column
            result = df[df.iloc[:, 0].astype(str) == search_id]  # Filter results
            print("\n📌 Matching Record Found:")
            print(result)
        else:
            print("❌ No record found for this Service Order ID.")
    else:
        print("❌ No records found. Please process images first.")

# ✅ Run the Bot
while True:
    bot_menu()
    choice = input("👉 Choose an option: ")

    if choice == "1":
        process_images()
    elif choice == "2":
        view_records()
    elif choice == "3":
        search_order()
    elif choice == "4":
        print("🚀 Exiting BOT. Goodbye!")
        break
    else:
        print("❌ Invalid choice. Please try again.")




# import cv2  # For image processing
# import pytesseract  # For OCR text extraction
# import pandas as pd  # For working with Excel
# import os  # For file operations
# import glob  # For handling multiple images

# # ✅ Set the correct Tesseract OCR path
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# # ✅ Define the folder containing images and the output Excel file
# image_folder = "images/"
# excel_file = "service_orders.xlsx"

# # ✅ Function to extract text from an image
# def extract_text_from_image(image_path):
#     image = cv2.imread(image_path)  # Read the image
#     gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  # Convert to grayscale
#     extracted_text = pytesseract.image_to_string(gray)  # Extract text

#     data_dict = {}  # Dictionary to store key-value pairs
#     lines = extracted_text.split("\n")  # Split text into lines

#     for line in lines:
#         if ":" in line:  # Check for key-value pairs
#             key, value = line.split(":", 1)  # Split at first ':'
#             data_dict[key.strip()] = value.strip()  # Store in dictionary

#     return data_dict  # Return extracted data

# # ✅ Process all images in the folder and store data in Excel
# def process_images():
#     all_data = []  # List to store extracted data

#     for img_path in glob.glob(os.path.join(image_folder, "*.jpg")):  # Get all .jpg images
#         extracted_data = extract_text_from_image(img_path)  # Extract text
#         if extracted_data:
#             all_data.append(extracted_data)  # Add data to the list

#     if all_data:  # If we have extracted data
#         df = pd.DataFrame(all_data)  # Convert list to DataFrame

#         if os.path.exists(excel_file):  # Check if Excel file exists
#             existing_df = pd.read_excel(excel_file)  # Load existing data
#             df = pd.concat([existing_df, df], ignore_index=True)  # Append new data

#         df.to_excel(excel_file, index=False)  # Save to Excel
#         print(f"✅ {len(all_data)} images processed and saved to '{excel_file}'!")
#     else:
#         print("❌ No valid data extracted from images.")

# # ✅ Function to display the main menu
# def bot_menu():
#     print("\n📌 Service Order BOT Menu:")
#     print("1. Process Service Order Images")
#     print("2. View All Records")
#     print("3. Search for a Specific Service Order ID")
#     print("4. Exit")

# # ✅ Function to view all stored records
# def view_records():
#     if os.path.exists(excel_file):  # Check if the file exists
#         df = pd.read_excel(excel_file)  # Load data
#         print("\n📄 Service Orders Data:")
#         print(df)  # Display records
#     else:
#         print("❌ No records found. Please process images first.")

# # ✅ Function to search for a specific Service Order ID
# def search_order():
#     if os.path.exists(excel_file):  # Check if the file exists
#         df = pd.read_excel(excel_file)  # Load data
#         search_id = input("🔍 Enter Service Order ID: ").strip()  # Get user input

#         if search_id in df.iloc[:, 0].astype(str).values:  # Check first column
#             result = df[df.iloc[:, 0].astype(str) == search_id]  # Filter results
#             print("\n📌 Matching Record Found:")
#             print(result)
#         else:
#             print("❌ No record found for this Service Order ID.")
#     else:
#         print("❌ No records found. Please process images first.")

# # ✅ Run the Bot
# while True:
#     bot_menu()
#     choice = input("👉 Choose an option: ")

#     if choice == "1":
#         process_images()
#     elif choice == "2":
#         view_records()
#     elif choice == "3":
#         search_order()
#     elif choice == "4":
#         print("🚀 Exiting BOT. Goodbye!")
#         break
#     else:
#         print("❌ Invalid choice. Please try again.")
