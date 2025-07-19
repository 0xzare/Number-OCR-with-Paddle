import pandas as pd
import os
import glob
from datetime import datetime
from paddleocr import PaddleOCR

ocr = PaddleOCR(lang="en", device="gpu:0", use_angle_cls=True)


def find_numeric_elements(arr):
    ten_digit = None
    fifteen_digit = None

    for item in arr:
        # Check if the item has exactly 10 or 15 characters and consists only of digits
        if len(item) == 10 and item.isdigit():
            ten_digit = item
        elif len(item) == 15 and item.isdigit():
            fifteen_digit = item

        # If both elements are found, break early for efficiency
        if ten_digit is not None and fifteen_digit is not None:
            break

    return ten_digit, fifteen_digit


def process_single_image(image_path):
    """پردازش یک تصویر و استخراج شماره تلفن و سریال"""
    try:
        ocr_img = ocr.ocr(image_path, cls=True)
        result = []
        for res in ocr_img[0]:
            box, (text, confidence) = res
            result.append(text)
        phone_number, serial_number = find_numeric_elements(result)

        return phone_number, serial_number, "Success"

    except Exception as e:
        print(f"Error processing {image_path}: {str(e)}")
        return None, None, f"Error: {str(e)}"


def process_images_from_folder(folder_path, output_excel_path=None):
    """Process all images in folder and save results to Excel"""

    # فرمت‌های تصویری پشتیبانی شده
    image_extensions = ["*.jpg", "*.jpeg", "*.png", "*.bmp", "*.tiff", "*.tif"]

    # پیدا کردن تمام فایل‌های تصویری
    image_files = []
    for extension in image_extensions:
        image_files.extend(glob.glob(os.path.join(folder_path, extension)))

    if not image_files:
        print("No image files found in the folder!")
        return

    print(f"Found {len(image_files)} images.")

    # لیست برای ذخیره نتایج
    results = []

    # Process each image
    for i, image_path in enumerate(image_files, 1):
        print(
            f"Processing image {i}/{len(image_files)}: {os.path.basename(image_path)}"
        )

        phone, serial, status = process_single_image(image_path)

        results.append(
            {
                "Filename": os.path.basename(image_path),
                "Full Path": image_path,
                "Phone Number": phone if phone else "",
                "Serial Number": serial if serial else "",
                "Status": status,
                "Processing Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        )

        # Display result
        print(f"  - Phone Number: {phone if phone else 'Not found'}")
        print(f"  - Serial Number: {serial if serial else 'Not found'}")
        print(f"  - Status: {status}")
        print("-" * 50)

    # ایجاد DataFrame
    df = pd.DataFrame(results)

    # تعیین نام فایل خروجی
    if output_excel_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_excel_path = f"ocr_results_{timestamp}.xlsx"

    # Save to Excel
    try:
        with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="OCR Results", index=False)

            # Adjust column widths
            worksheet = writer.sheets["OCR Results"]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"\nResults successfully saved to file: {output_excel_path}")

        # Display summary statistics
        successful_extractions = df[df["Status"] == "Success"].shape[0]
        phone_found = df[df["Phone Number"] != ""].shape[0]
        serial_found = df[df["Serial Number"] != ""].shape[0]

        print(f"\n📊 Summary Statistics:")
        print(f"- Total images: {len(image_files)}")
        print(f"- Successfully processed: {successful_extractions}")
        print(f"- Phone numbers found: {phone_found}")
        print(f"- Serial numbers found: {serial_found}")

    except Exception as e:
        print(f"Error saving Excel file: {str(e)}")


def main():
    """Main program function"""

    # Path to folder containing images
    folder_path = input(
        "Enter the path to folder containing images (or Enter for current folder): "
    ).strip()

    if not folder_path:
        folder_path = "."  # Current folder

    if not os.path.exists(folder_path):
        print("The specified folder does not exist!")
        return

    # Output Excel file path (optional)
    output_path = input(
        "Enter output Excel filename (or Enter for automatic name): "
    ).strip()

    if not output_path:
        output_path = None

    # Start processing
    print(f"\nStarting image processing from folder: {os.path.abspath(folder_path)}")
    import time

    time1 = time.time()
    process_images_from_folder(folder_path, output_path)
    time2 = time.time()
    print("Calculation Time:", time2 - time1)


if __name__ == "__main__":
    main()

# For direct usage in code:
# process_images_from_folder("path/to/your/images/folder", "results.xlsx")
