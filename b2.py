import cv2
import pytesseract
import pandas as pd
import os
import glob

# Set Tesseract path - UPDATE THIS TO YOUR TESSERACT INSTALLATION PATH
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Configuration
image_folder = "images/"
excel_file = "service_orders.xlsx"

def preprocess_image(image_path):
    """Optimized preprocessing for service order documents"""
    try:
        # Read image
        img = cv2.imread(image_path)
        if img is None:
            print(f"Error: Could not read image at {image_path}")
            return None
        
        # Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # Apply mild blur to reduce noise while preserving edges
        gray = cv2.medianBlur(gray, 3)
        
        # Thresholding - inverted for dark text on light background
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        
        return thresh
    except Exception as e:
        print(f"Preprocessing error for {image_path}: {str(e)}")
        return None

def extract_text_from_image(image_path):
    """Specialized extraction for service order key-value pairs"""
    try:
        processed_img = preprocess_image(image_path)
        if processed_img is None:
            return {}
        
        # Custom OCR configuration for form-like documents
        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(processed_img, config=custom_config)
        
        data = {}
        current_key = None
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        for line in lines:
            # Skip checkbox symbols
            line = line.replace("✔", "").strip()
            
            # Check for key-value pairs
            if ':' in line:
                key_part, value_part = line.split(':', 1)
                current_key = key_part.strip()
                data[current_key] = value_part.strip()
            elif current_key and line:
                # Handle multi-line values
                data[current_key] += ' ' + line
            elif line and not any(c.isalpha() for c in line):
                # Handle standalone values (IDs, dates)
                if current_key:
                    data[current_key] += ' ' + line
                elif 'Service Order ID' not in data:
                    data['Service Order ID'] = line
        
        # Post-process common OCR errors specific to your documents
        data = {k: (v.replace('_', ' ') 
                   .replace('Ovminq', 'Owning')
                   .replace('Dalo', 'Date')
                   .replace('...', '')) 
               for k, v in data.items()}
        
        return data
    except Exception as e:
        print(f"Extraction error for {image_path}: {str(e)}")
        return {}

def process_images():
    """Process all images in the designated folder"""
    all_data = []
    
    # Get all JPG and PNG files in the images folder
    image_files = (glob.glob(os.path.join(image_folder, "*.jpg")) + 
                  glob.glob(os.path.join(image_folder, "*.png")))
    
    if not image_files:
        print(f"❌ No images found in {image_folder}")
        return
    
    for img_path in image_files:
        print(f"\nProcessing {os.path.basename(img_path)}...")
        extracted_data = extract_text_from_image(img_path)
        
        if extracted_data:
            print("Extracted data:")
            for key, value in extracted_data.items():
                print(f"  {key}: {value}")
            all_data.append(extracted_data)
        else:
            print(f"⚠️ No data extracted from {os.path.basename(img_path)}")
    
    if all_data:
        # Create or update Excel file
        df = pd.DataFrame(all_data)
        
        if os.path.exists(excel_file):
            existing_df = pd.read_excel(excel_file)
            updated_df = pd.concat([existing_df, df], ignore_index=True)
            updated_df.to_excel(excel_file, index=False)
        else:
            df.to_excel(excel_file, index=False)
        
        print(f"\n✅ Successfully processed {len(all_data)} images. Data saved to {excel_file}")
    else:
        print("\n❌ No valid data extracted from any images")

def view_records():
    """Display all records from the Excel file"""
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
        if not df.empty:
            print("\n📄 All Service Orders:")
            print(df.to_string(index=False))
        else:
            print("\nℹ️ No records found in the database")
    else:
        print("\n❌ Database file not found. Please process images first")

def search_order():
    """Search for a specific service order by ID"""
    if not os.path.exists(excel_file):
        print("\n❌ Database file not found. Please process images first")
        return
    
    df = pd.read_excel(excel_file)
    if df.empty:
        print("\nℹ️ No records available to search")
        return
    
    search_id = input("\n🔍 Enter Service Order ID to search: ").strip()
    if not search_id:
        print("❌ Please enter a valid Service Order ID")
        return
    
    # Search in all columns that might contain IDs
    results = df[df.apply(lambda row: search_id in ' '.join(row.astype(str).values), axis=1)]
    
    if not results.empty:
        print("\n📌 Matching Record(s):")
        print(results.to_string(index=False))
    else:
        print(f"\n❌ No records found for Service Order ID: {search_id}")

def bot_menu():
    """Display the main menu"""
    print("\n" + "="*40)
    print("📋 Service Order Processing Bot")
    print("="*40)
    print("1. Process Service Order Images")
    print("2. View All Records")
    print("3. Search for a Service Order")
    print("4. Exit")
    print("="*40)

if __name__ == "__main__":
    print("\n🚀 Service Order Processing Bot Initialized")
    
    # Create images folder if it doesn't exist
    if not os.path.exists(image_folder):
        os.makedirs(image_folder)
        print(f"📁 Created images folder at: {os.path.abspath(image_folder)}")
    
    while True:
        bot_menu()
        choice = input("\n👉 Select an option (1-4): ").strip()
        
        if choice == "1":
            process_images()
        elif choice == "2":
            view_records()
        elif choice == "3":
            search_order()
        elif choice == "4":
            print("\n🚀 Exiting the bot. Goodbye!")
            break
        else:
            print("\n❌ Invalid option. Please choose 1-4")