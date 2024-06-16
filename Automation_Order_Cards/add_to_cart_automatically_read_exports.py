import os
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException, ElementNotInteractableException
import time

# Define the directory containing the export files and the reference file
export_dir = 'D:/Add to card automatically/Automation_Order_Cards/Exports'
reference_file = 'D:/Add to card automatically/Automation_Order_Cards/reference_file.xlsx'
output_file = 'combined_order_details.xlsx'
aggregated_output_file = 'aggregated_order_details.xlsx'

# Delete the current combined_order_details file if it exists
if os.path.exists(output_file):
    os.remove(output_file)
    print(f"Deleted existing file: {output_file}")

if os.path.exists(aggregated_output_file):
    os.remove(aggregated_output_file)
    print(f"Deleted existing file: {aggregated_output_file}")

# Function to read and combine all order detail files with debugging
def read_order_files(directory):
    combined_df = pd.DataFrame()
    for filename in os.listdir(directory):
        if filename.startswith('export') and filename.endswith('.csv'):
            file_path = os.path.join(directory, filename)
            try:
                df = pd.read_csv(file_path)
                print(f"Successfully read {filename} with columns: {df.columns.tolist()} and shape: {df.shape}")
                combined_df = pd.concat([combined_df, df], ignore_index=True)
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
    return combined_df

# Load the reference data
try:
    reference_df = pd.read_excel(reference_file, engine='openpyxl')
    print(f"Successfully read reference file with columns: {reference_df.columns.tolist()} and shape: {reference_df.shape}")
except FileNotFoundError as e:
    print(f"Reference file not found: {e}")
    exit()
except Exception as e:
    print(f"Error reading reference file: {e}")
    exit()

# Combine all order details and print the combined DataFrame information
order_details_df = read_order_files(export_dir)
print(f"Combined DataFrame columns: {order_details_df.columns.tolist()} and shape: {order_details_df.shape}")

# Check for specific columns in both DataFrames
if 'Final URL' in reference_df.columns and 'Image' in order_details_df.columns:
    order_details_df['Modified_URL'] = order_details_df['Image'].apply(lambda x: str(x).split('?')[0] if pd.notnull(x) else x)
    reference_df['Modified_URL'] = reference_df['Image URL'].apply(lambda x: str(x).split('?')[0] if pd.notnull(x) else x)
else:
    print("Error: 'Final URL' or 'Image' column not found in the respective DataFrame.")
    exit()

# Drop duplicates to keep only the first instance of each card in the reference DataFrame
reference_df = reference_df.drop_duplicates(subset=['Modified_URL'], keep='first')

# Perform the lookup equivalent
combined_df = pd.merge(order_details_df, reference_df, on='Modified_URL', how='left')

# Save the combined data to a new Excel file
combined_df.to_excel(output_file, index=False)

# Aggregate data to get unique URLs and their corresponding quantities
aggregated_df = combined_df.groupby('Final URL').size().reset_index(name='Quantity')

# Save the aggregated data to a new Excel file
aggregated_df.to_excel(aggregated_output_file, index=False)

# Identify cards with blank Final URL
blank_final_url_df = combined_df[combined_df['Final URL'].isnull()]

# Print cards with blank Final URL
if not blank_final_url_df.empty:
    print("\nCards with blank Final URL:")
    for index, row in blank_final_url_df.iterrows():
        print(f"Card: {row['Name']}, Box Name: {row['Box Name']}")

# Initialize the undetected chromedriver
driver = uc.Chrome()

def wait_for_non_empty_text(driver, locator, timeout=10):
    return WebDriverWait(driver, timeout).until(
        lambda d: d.find_element(*locator).text.strip(),
        f"Element with locator {locator} did not have non-empty text after {timeout} seconds"
    )

def gather_listings(card_url):
    driver.get(card_url)
    listings_data = []

    while True:
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.listing-item__listing-data'))
            )
            listings = driver.find_elements(By.CSS_SELECTOR, '.listing-item__listing-data')

            for listing in listings:
                try:
                    available_quantity_element = listing.find_element(By.CSS_SELECTOR, '.add-to-cart__available')
                    available_quantity_text = available_quantity_element.text.strip().split()[-1]
                    available_quantity = int(available_quantity_text)
                    add_to_cart_button = listing.find_element(By.CSS_SELECTOR, '.add-to-cart__submit')
                    listings_data.append((add_to_cart_button, available_quantity))
                except NoSuchElementException:
                    continue  # If any element is not found, move to the next listing

            # Check if there are more pages and navigate to the next page if available
            next_page_buttons = driver.find_elements(By.CSS_SELECTOR, '.pagination__next')
            if next_page_buttons and next_page_buttons[0].is_enabled():
                next_page_buttons[0].click()
                WebDriverWait(driver, 20).until(EC.staleness_of(listings[0]))  # Wait for the listings to refresh
            else:
                break  # No more pages

        except TimeoutException as e:
            print(f"An error occurred while gathering listings: {e}")
            break

    return listings_data

def add_card_to_cart(listings_data, card_name, card_url, desired_quantity):
    total_added = 0
    order_summary = []

    for index, (add_to_cart_button, available_quantity) in enumerate(listings_data):
        if total_added >= desired_quantity:
            break

        while available_quantity > 0 and total_added < desired_quantity:
            try:
                WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable(add_to_cart_button)
                )
                add_to_cart_button.click()
                time.sleep(2)  # Small delay to ensure the action is processed
                
                # Check for popup after attempting to add to cart
                if is_popup_present():
                    print("Popup detected after attempting to add to cart.")
                    time.sleep(2)  # Wait for 2 seconds
                    try:
                        okay_button = driver.find_element(By.CSS_SELECTOR, '.add-item-error__action__primary-btn')
                        okay_button.click()
                        time.sleep(1)  # Small delay to ensure the action is processed
                        print("Clicked 'Okay' button on popup.")
                        break  # Skip to the next listing
                    except (NoSuchElementException, TimeoutException) as popup_e:
                        print(f"Error clicking 'Okay' button: {popup_e}")
                        break  # Skip to the next listing
                else:
                    total_added += 1
                    available_quantity -= 1
                    print(f"Added 1 of {card_name} to cart.")
                
            except (ElementClickInterceptedException, ElementNotInteractableException) as e:
                print(f"Error clicking add to cart button: {e}")
                if is_popup_present():
                    print("Popup detected while attempting to add to cart.")
                    time.sleep(2)  # Wait for 2 seconds
                    try:
                        okay_button = driver.find_element(By.CSS_SELECTOR, '.add-item-error__action__primary-btn')
                        okay_button.click()
                        time.sleep(1)  # Small delay to ensure the action is processed
                        print("Clicked 'Okay' button on popup.")
                        break  # Skip to the next listing
                    except (NoSuchElementException, TimeoutException) as popup_e:
                        print(f"Error clicking 'Okay' button: {popup_e}")
                        break  # Skip to the next listing
                else:
                    continue

    if total_added < desired_quantity:
        print(f"Could not add the desired quantity of {card_name} to the cart. Added {total_added} out of {desired_quantity}.")
    
    order_summary.append((card_name, total_added))
    return order_summary, total_added

def is_popup_present():
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.add-item-error__action__primary-btn'))
        )
        return True
    except (NoSuchElementException, TimeoutException):
        return False

order_summary = []
unable_to_order = []

try:
    # Open the login page
    driver.get('https://store.tcgplayer.com/login')

    # Wait until the login page is loaded
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'Email')))

    # Find the username and password fields and the login button
    username_field = driver.find_element(By.NAME, 'Email')
    password_field = driver.find_element(By.NAME, 'Password')
    login_button = driver.find_element(By.XPATH, '//button[@type="submit"]')

    # Enter your credentials
    username_field.send_keys('maul.happier@gmail.com')
    password_field.send_keys('Omegaaizabubby1!')

    # Click the login button
    login_button.click()

    # Wait for you to log in manually if any additional verification is needed
    print("Please complete any additional verification if prompted. Type 'ok' in the terminal to continue.")
    while input().strip().lower() != 'ok':
        print("Please type 'ok' to continue after logging in.")

    # Read URLs and quantities from the aggregated Excel file
    df = pd.read_excel(aggregated_output_file, engine='openpyxl')

    # Iterate through each card URL and add to cart
    for index, row in df.iterrows():
        card_url = row['Final URL']
        desired_quantity = row['Quantity']
        try:
            listings_data = gather_listings(card_url)
            print(f"Listings for {card_url}:")
            for btn, qty in listings_data:
                print(f"Available quantity: {qty}")

            summary, added_quantity = add_card_to_cart(listings_data, card_url.split('/')[-1], card_url, desired_quantity)
            order_summary.extend(summary)
            if added_quantity < desired_quantity:
                unable_to_order.append((card_url, desired_quantity - added_quantity))
        except Exception as e:
            print(f"An error occurred while processing {card_url}: {e}")
            unable_to_order.append((card_url, desired_quantity))
            continue  # Move to the next card

    print("All cards processed. The browser will remain open for manual actions.")
    print("\nOrder Summary:")
    for card, quantity in order_summary:
        print(f"Ordered {quantity} of {card}")

    if unable_to_order:
        print("\nUnable to Order:")
        for card_url, quantity in unable_to_order:
            print(f"Unable to order {quantity} of {card_url}")

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    # Do not quit the driver to keep the browser open
    pass

# Print cards with blank Final URL
if not blank_final_url_df.empty:
    print("\nCards with blank Final URL:")
    for index, row in blank_final_url_df.iterrows():
        print(f"Card: {row['Name']}, Box Name: {row['Box Name']}")
