# Generates JSON library to import into Fusion, and excel data for Harvey, Helical, Kodiak, Haas and GARR using URL from JSON Library

import json
import os
import pandas as pd
import requests
from bs4 import BeautifulSoup, Comment

# Function to search for the tool in the JSON files and log the details
def search_in_json_files(barcode, json_files, tools_log):
    try:
        # Read existing data from New_Tool_Library.json if it exists
        with open("Inset File Path and Name of JSON file to import into fusion", 'r', encoding='utf-8-sig') as f:
            existing_data = json.load(f).get('data', [])
    except FileNotFoundError:
        existing_data = []

    found_tools = []
    # Check each JSON file for the barcode
    for json_file in json_files:
        with open(json_file, 'r', encoding='utf-8-sig') as f:
            data = json.load(f)
        
        for tool in data.get('data', []):
            if tool.get('product-id') == barcode:
                found_tools.append((json_file, tool))

    # Handle multiple or single entries found
    if len(found_tools) > 1:
        print("Multiple entries found. Please select the manufacturer:")
        for i, (file_path, tool) in enumerate(found_tools):
            manufacturer_name = os.path.basename(file_path).split('-')[0]
            print(f"{i + 1}: {manufacturer_name}")
        
        selection = int(input("Enter the number of the manufacturer you wish to select: ")) - 1
        if 0 <= selection < len(found_tools):
            selected_tool_info = found_tools[selection][1]
            manufacturer = os.path.basename(found_tools[selection][0]).split('-')[0]
            existing_data.append(selected_tool_info)
            tools_log[barcode] = {
                'manufacturer': manufacturer,
                'product-link': selected_tool_info.get('product-link', 'No link available')
            }
    elif found_tools:
        selected_tool_info = found_tools[0][1]
        manufacturer = os.path.basename(found_tools[0][0]).split('-')[0]
        existing_data.append(selected_tool_info)
        tools_log[barcode] = {
            'manufacturer': manufacturer,
            'product-link': selected_tool_info.get('product-link', 'No link available')
        }
    else:
        manufacturer = input(f"Enter the manufacturer for barcode {barcode}: ").strip()
        tools_log[barcode] = {'manufacturer': manufacturer}

    # Write the updated data back to New_Tool_Library.json
    if existing_data:
        with open("/Users/henrywright/Desktop/Shop Model/JSON - Tool Libraries/New_Tool_Library.json", 'w', encoding='utf-8') as f:
            json.dump({'data': existing_data}, f)

    return manufacturer




if __name__ == "__main__":
    json_files = [
                   "Download JSON files from here - https://cam.autodesk.com/hsmtools - and then save the filepath of where you put each file. Ex - /users/you/Desktop/Json_Files/Haas"
            ]
    tools_log = {}  # Dictionary to keep track of all scanned tools

    while True:
        user_input = input("Enter the barcode (or type 'exit' to stop, 'edit' to edit a tool): ").strip()
        if user_input.lower() == 'exit':
            # Save tools_log to a file before exiting
            with open("/Users/henrywright/Desktop/Shop Model/JSON - Tool Libraries/Tools_Log.json", 'w', encoding='utf-8') as f:
                json.dump(tools_log, f)
            print("Exiting program.")
            break
        elif user_input.lower() == 'edit':
            # Allow editing an existing manufacturer
            barcode_to_edit = input("Enter the barcode of the tool to edit: ").strip()
            if barcode_to_edit in tools_log:
                new_manufacturer = input("Enter the new manufacturer: ").strip()
                tools_log[barcode_to_edit] = new_manufacturer
                print(f"Barcode {barcode_to_edit} updated to new manufacturer '{new_manufacturer}'.")
            else:
                print(f"Barcode {barcode_to_edit} not found in the log. Please scan the tool first.")
        else:
            # Normal barcode scanning and logging
            manufacturer = search_in_json_files(user_input, json_files, tools_log)
            print(f"Barcode {user_input} recorded for manufacturer '{manufacturer}'.")



# The scraping functions you will define for each manufacturer
def fetch_garrtool_details(product_link, entered_barcode):
    url = product_link
    print(f"Fetching data from {url}...")
    
    response = requests.get(url)
    print(f"HTTP Response: {response.status_code}")
    
    if response.status_code != 200:
        print("Failed to fetch the webpage.")
        return None

    html_content = response.content
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Extracting tool dimensions
    dimensions = {}
    item_table = soup.find("ul", class_="item-table")
    if not item_table:  # Check if item_table is None
        print(f"Failed to find item table for barcode {entered_barcode} on GARRTOOL website.")
        return None  # Return None or handle as unrecognized entry

    headers = item_table.find("li", class_="head").find_all("span", class_="text")
    values = item_table.find_all("li")[1].find_all("span", class_="text")
    for header, value in zip(headers, values):
        dimensions[header.text] = value.text

    # Extracting tool details
    details_ul = soup.find("ul", class_="info-list")
    details_list = details_ul.find_all("li")
    details = ", ".join([li.text for li in details_list])

    # Extracting pricing information
    pricing_ul = soup.find("ul", class_="list-wrap")
    first_price = pricing_ul.find_all("li")[1].find_all("span", class_="text")[0].text

    # Preparing the DataFrame
    dimensions['Details'] = details
    dimensions['First Price'] = first_price
    dimensions['Source'] = 'GARRTOOL'
    dimensions['Barcode'] = entered_barcode
    
    dimensions['URL'] = url 
    df = pd.DataFrame([dimensions])
    return df


def fetch_helicaltool_details(product_link, entered_barcode):
    url = product_link
    print(f"Fetching data from {url}...")
    response = requests.get(url)
    print(f"HTTP Response: {response.status_code}")

    html_content = response.content
    soup = BeautifulSoup(html_content, 'html.parser')
    comment = soup.find(string=lambda text: isinstance(text, Comment) and "Tool Dimension" in text)

    if comment:
            print("Comment found, processing data...")
            
            def get_dimensions_from_list(dimension_list):
                dimensions = {}
                for item in dimension_list:
                    dimension_text = item.find("span").string.strip(": ")
                    dimension_value = item.find_all("span")[1].text.rstrip('"')
                    if dimension_text != "Catalog Page":
                        dimensions[dimension_text] = dimension_value
                return dimensions

            all_dimensions = {}
            dimension_lists = soup.find_all("ul", class_="dimension-list")
            for dimension_list in dimension_lists:
                all_dimensions.update(get_dimensions_from_list(dimension_list.find_all("li")))

            
            try:
                machining_advisor_pro_link = soup.find("a", href=True, string="Open in Machining Advisor Pro")['href']
            except TypeError:
                print("Could not find Machining Advisor Pro link.")
                machining_advisor_pro_link = "N/A"

            try:
                feeds_speeds_link = soup.find("a", href=True, string="Download Speeds & Feeds PDF")['href']
            except TypeError:
                print("Could not find Speeds & Feeds link.")
                feeds_speeds_link = "N/A"

            try:
                sim_file_link = soup.find("a", href=True, string="Download SIM File")['href']
            except TypeError:
                print("Could not find SIM File link.")
                sim_file_link = "N/A"

            price_span = soup.find("span", string="Price:")
            if price_span:
                price_value = price_span.find_next_sibling("span").text
                all_dimensions['Price'] = price_value

            source = 'HELICAL' if 'helicaltool' in url else 'HARVEY'
            ordered_dimensions = {'Source': source}
            for key in list(all_dimensions.keys())[:-4] + ['Price'] + list(all_dimensions.keys())[-4:]:
                ordered_dimensions[key] = all_dimensions[key]

            ordered_dimensions['Machining Advisor Pro'] = machining_advisor_pro_link
            ordered_dimensions['Feeds & Speeds'] = feeds_speeds_link
            ordered_dimensions['Download SIM File'] = sim_file_link
            ordered_dimensions['Buy Now!!!'] = url
            ordered_dimensions['Barcode'] = entered_barcode
            
            df = pd.DataFrame([ordered_dimensions])
            return df
    else:
            print(f"No comment found for tool dimension on {url}")
    
    return None


def fetch_harveytool_details(product_link, entered_barcode):
    url = product_link
    print(f"Fetching data from {url}...")
    response = requests.get(url)
    print(f"HTTP Response: {response.status_code}")
    if response.status_code != 200:
        print("Failed to fetch the webpage.")
        return None

    html_content = response.content
    soup = BeautifulSoup(html_content, 'html.parser')
    comment = soup.find(string=lambda text: isinstance(text, Comment) and "Tool Dimension" in text)

    if comment:
        print("Comment found, processing data...")
        
        def get_dimensions_from_list(dimension_list):
            dimensions = {}
            for item in dimension_list:
                dimension_text = item.find("span").string.strip(": ")
                dimension_value = item.find_all("span")[1].text.rstrip('"')
                dimensions[dimension_text] = dimension_value
            return dimensions

        all_dimensions = {}
        dimension_lists = soup.find_all("ul", class_="dimension-list")
        for dimension_list in dimension_lists:
            all_dimensions.update(get_dimensions_from_list(dimension_list.find_all("li")))

        # Fetch additional information like links and prices
        try:
            machining_advisor_pro_link = soup.find("a", href=True, string="Open in Machining Advisor Pro")['href']
        except TypeError:
            machining_advisor_pro_link = "N/A"

        try:
            feeds_speeds_link = soup.find("a", href=True, string="Download Speeds & Feeds PDF")['href']
        except TypeError:
            feeds_speeds_link = "N/A"

        try:
            sim_file_link = soup.find("a", href=True, string="Download SIM File")['href']
        except TypeError:
            sim_file_link = "N/A"

        price_span = soup.find("span", string="Price:")
        if price_span:
            price_value = price_span.find_next_sibling("span").text
            all_dimensions['Price'] = price_value

        all_dimensions['Machining Advisor Pro'] = machining_advisor_pro_link
        all_dimensions['Feeds & Speeds'] = feeds_speeds_link
        all_dimensions['Download SIM File'] = sim_file_link
        all_dimensions['Source'] = 'HARVEYTOOL'
        all_dimensions['Barcode'] = entered_barcode
        
        # Reformatting the data into a DataFrame
        df = pd.DataFrame([all_dimensions])
        return df
    else:
        print(f"No comment found for tool dimension on {url}")
        return None


def fetch_haastool_details(product_link, entered_barcode):
    url = product_link
    print(f"Fetching data from {url}...")
    
    response = requests.get(url)
    print(f"HTTP Response: {response.status_code}")
    
    if response.status_code != 200:
        print("Failed to fetch the webpage.")
        return None

    html_content = response.content
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Extracting price
    price_container = soup.find("div", class_="current-price-container")
    if price_container:
        price = price_container.find("span", class_="selected-currency").text.strip()
    else:
        print("Price information not found.")
        return None

    # Extracting tool details
    tool_details = {}
    details_table = soup.find_all("tr")
    for row in details_table:
        cells = row.find_all("td")
        if len(cells) == 2:  # Ensure that there are exactly two cells
            label = cells[0].get_text(strip=True)
            value = cells[1].get_text(strip=True)
            tool_details[label] = value

    # Preparing the DataFrame
    tool_details['Price'] = price
    tool_details['Source'] = 'Haas'
    tool_details['Barcode'] = entered_barcode
    tool_details['URL'] = url 
    df = pd.DataFrame([tool_details])
    return df


def fetch_kodiaktool_details(product_link, entered_barcode):
    print(f"Fetching data from {product_link}...")
    response = requests.get(product_link)
    print(f"HTTP Response: {response.status_code}")
    
    if response.status_code != 200:
        print("Failed to fetch the webpage.")
        return None

    html_content = response.content
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Finding the table row (tr) that contains the barcode
    barcode_td = soup.find("td", attrs={"partnumber": str(entered_barcode)})
    if not barcode_td:
        print(f"Failed to find barcode {entered_barcode} on the Kodiak Tool website.")
        return None

    # Finding the parent row of the barcode cell
    tool_row = barcode_td.parent
    if not tool_row:  # Check if tool_row is None
        print(f"Failed to find the tool row for barcode {entered_barcode}.")
        return None

    # Extracting tool data
    tool_data = {}
    for td in tool_row.find_all("td"):
        data_label = td.get('data-label')
        if data_label and data_label != 'Item':
            tool_data[data_label] = td.get_text(strip=True)

    # Assuming 'Price' is a required field, check if it's present
    if 'Price' not in tool_data:
        print(f"Price information is missing for barcode {entered_barcode}.")
        return None
    
    # Adding additional data
    tool_data['Barcode'] = entered_barcode
    tool_data['URL'] = product_link
    
    # Preparing the DataFrame
    df = pd.DataFrame([tool_data])
    
    return df


def save_to_excel(manufacturer_dfs, path, unrecognized_df=None):
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        for manufacturer, df in manufacturer_dfs.items():
            if not df.empty:
                # Reorder columns to have 'Barcode' first if it exists
                if 'Barcode' in df.columns:
                    cols = ['Barcode'] + [col for col in df if col != 'Barcode']
                    df = df[cols]

                # Write DataFrame to Excel sheet
                df.to_excel(writer, sheet_name=manufacturer, index=False)

        # If there's an unrecognized DataFrame, save it to a separate sheet
        if unrecognized_df is not None and not unrecognized_df.empty:
            unrecognized_df.to_excel(writer, sheet_name='Unrecognized Tools', index=False)

    print(f"Excel file '{path}' successfully created with separate sheets for each manufacturer and unrecognized tools.")



# Function to load the JSON data
def load_json_data(filepath):
    with open(filepath, 'r') as file:
        return json.load(file)

def scrape_and_save_data(json_filepath, excel_filepath):
    data = load_json_data(json_filepath)
    manufacturer_dfs = {}  # Dictionary to keep each manufacturer's DataFrame
    unrecognized_data = []  # List to keep unrecognized barcodes and manufacturers
    
    for barcode, tool_info in data.items():
        df = None
        if tool_info['manufacturer'] == 'Garr Tool':
            df = fetch_garrtool_details(tool_info['product-link'], barcode)
        elif tool_info['manufacturer'] == 'Helical Solutions':
            df = fetch_helicaltool_details(tool_info['product-link'], barcode)
        elif tool_info['manufacturer'] == 'Harvey Tool':
            df = fetch_harveytool_details(tool_info['product-link'], barcode)
        elif tool_info['manufacturer'] == 'Haas Tooling':
            df = fetch_haastool_details(tool_info['product-link'], barcode)
        elif tool_info['manufacturer'] == 'Kodiak':
            df = fetch_kodiaktool_details(tool_info['product-link'], barcode)
        # Add more elif statements for other manufacturers
        
        if df is not None:
            # Check if the manufacturer's sheet already exists, if not create one
            if tool_info['manufacturer'] not in manufacturer_dfs:
                manufacturer_dfs[tool_info['manufacturer']] = pd.DataFrame()
            manufacturer_dfs[tool_info['manufacturer']] = pd.concat([manufacturer_dfs[tool_info['manufacturer']], df], ignore_index=True)
        else:
            # Handle unrecognized barcode
            unrecognized_data.append({'Barcode': barcode, 'Manufacturer': tool_info['manufacturer']})

    # Save each manufacturer's DataFrame to its own sheet
    with pd.ExcelWriter(excel_filepath, engine='openpyxl') as writer:
        for manufacturer, df in manufacturer_dfs.items():
            df.to_excel(writer, sheet_name=manufacturer, index=False)
        
        # Save unrecognized barcodes to the 'Unrecognized Tools' sheet
        pd.DataFrame(unrecognized_data).to_excel(writer, sheet_name='Unrecognized Tools', index=False)

    print(f"Excel file '{excel_filepath}' has been created successfully with separate sheets for each manufacturer.")

# Main entry point of the script
if __name__ == '__main__':
    json_filepath = '/Users/henrywright/Desktop/Shop Model/JSON - Tool Libraries/Tools_Log.json'
    excel_filepath = '/Users/henrywright/Desktop/Shop Model/Tool Management /Tool_List.xlsx'
    scrape_and_save_data(json_filepath, excel_filepath)