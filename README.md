# eBay 3D Seller Template Generator

## Description
The eBay 3D Seller Template Generator is a desktop application designed to automate the process of creating product listings for eBay, particularly tailored for vendors who need to merge data from multiple sources into a single eBay-compatible upload format. The application supports loading multiple types of files (CSV, XLSX, ZIP), enabling users to compile comprehensive listings that include SKUs, titles, descriptions, categories, images, and other necessary eBay listing attributes.

## Features
- **Multi-File Input**: Accepts vendor items files, competitor files, PIES files, and category ID mappings.
- **Support for Multiple Formats**: Handles CSV and XLSX files directly and can extract from ZIP archives containing multiple XLSX files.
- **Attribute Aggregation**: Combines attributes from different segments like product attributes, PIES descriptions, and more into a single, cohesive listing.
- **Image Handling**: Supports up to 10 images per listing, automatically mapping and aligning with the respective SKU.
- **Customizable Output**: Generates an Excel file formatted specifically for eBay's 3D seller import requirements.

## Prerequisites
Before you start using the eBay 3D Seller Template Generator, ensure you have the following installed on your system:
- Python 3.8 or higher
- Pandas library
- OpenPyXL library (for handling Excel files)
- Tkinter library (for the GUI)

## Installation
1. **Clone the Repository**
   ```bash
   git clone https://github.com/lightningcraft-0201/Py-eBay-Template-Automation.git
   ```
2. **Navigate to the Project Directory**
   ```bash
   cd eBayTemplateGenerator
   ```
3. **Install Required Python Packages**
   ```bash
   pip install pandas openpyxl
   ```

## Usage
To use the eBay 3D Seller Template Generator, follow these steps:
1. **Start the Application**
   - Run `python app.py` from your command line to open the GUI.
2. **Load Files**
   - Use the 'Browse' buttons to load the required files for each section:
     - Vendor Items File (CSV)
     - Competitor Output Files (CSV)
     - PIES File (XLSX or ZIP)
     - Category ID File (CSV)
     - Image List File (CSV)
3. **Set the Brand Name**
   - Enter the brand name as it should appear in the eBay listings.
4. **Generate the Template**
   - Click 'Generate' to process the files and create the eBay listing template. The application will prompt you to save the output Excel file.

## Contributing
Contributions to the eBay 3D Seller Template Generator are welcome! Please fork the repository and submit a pull request with your enhancements. For major changes, please open an issue first to discuss what you would like to change.