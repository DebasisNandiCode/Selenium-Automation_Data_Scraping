from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import time
import pandas as pd
import os
from sqlalchemy import create_engine
import urllib
from urllib.parse import quote_plus
import pyodbc
from io import StringIO
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


# Function to send email
def send_email(subject, body, success=True):
    sender_email = "xxxxxxxx@xxxxxxxx.com"
    password = "xxxxxxxxxxxxx"
    receiver_email = "xxxxxxxx@xxxxxxxx.com"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    # Attached email body
    msg.attach(MIMEText(body, 'plain'))

    # Setup email server
    try:
        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.starttls()
            server.login(sender_email, password)
            text = msg.as_string()
            server.sendmail(sender_email, receiver_email, text)

    except Exception as e:
        print(f"Exception : {e}")


# Create download folder
# Find the previous date
previous_date = datetime.now() - timedelta(1)
year_month = previous_date.strftime("%Y%m")
day = previous_date.strftime("%d")

# Input dates for attendance URL
first_date_of_month = previous_date.replace(day=1).strftime('%m/%d/%Y')
yesterday = previous_date.strftime('%m/%d/%Y')

# Base directory path
base_dir = r"C:\Users\xxxxxx\xxxxxx\xxxxxx\xxxxxx\Raw_Data"

# Check if "yyyymm" folder exists
year_month_dir = os.path.join(base_dir, year_month)
if not os.path.exists(year_month_dir):
    os.makedirs(year_month_dir)
    print(f"Created a folder : {year_month_dir}")
else:
    print(f"Folder already exists : {year_month_dir}")

# Check if "dd" folder exists within the "yyyymm" folder
day_dir = os.path.join(year_month_dir, day)
if not os.path.exists(day_dir):
    os.makedirs(day_dir)
    print(f"Created a folder : {day_dir}")
else:
    print(f"Folder already exists : {day_dir}")


# Configure download directory
download_dir = os.path.join(base_dir, year_month, day)

# delete files in the folder
def delete_files_in_directory(directory_path):
    try:
        files = os.listdir(directory_path)
        for file in files:
            file_path = os.path.join(directory_path, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
        print("All files deleted successfully!")
    except OSError:
        print("Error occur while deleting files!")

delete_files_in_directory(download_dir)


#Configuring options for the Chrome browser 
chrome_options = Options()

prefs = {
    "download.default_directory": download_dir,       # Set the download directory
    "download.prompt_for_download": False,            # Disable download prompts, With this option set to False, files are downloaded automatically without user interaction.
    "directory_upgrade": True                         # Allow directory upgrades, It ensures that the specified download directory is used, even if it's not the default directory.
}

chrome_options.add_experimental_option("prefs", prefs) #This method adds the prefs dictionary to the Chrome options under the "prefs" key.


try:
    # Specify the path to chromedriver
    driver_path = r"C:\chromedriver.exe"
    service = Service(driver_path)

    # Use WebDriverManager to get the ChromeDriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

except Exception as e:
    send_email("Chrome Driver connection Failed for Data", f"Error: {str(e)}")
    raise e


# Navigate to the login page
url = "https://xxxxxxxxxxx/xxxxxxxxx"
driver.get(url)
time.sleep(2)

# Initialize WebDriverWait
wait = WebDriverWait(driver, 180)

# Increase the page load timeout
driver.set_page_load_timeout(300)  # Set to 5 minutes (300 seconds)

# Increase the implicit wait time
driver.implicitly_wait(30)  # Set to 30 seconds
try:
    # Perform login actions
    ID = "xxxxxxxx"
    Password = 'xxxxxxxxxxx'

    # Locate and interact with elements
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="sign-in-username"]'))).send_keys(ID)
    time.sleep(2)
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="password-field"]'))).send_keys(Password)
    time.sleep(2)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="login_area"]'))).click()
    time.sleep(2)

    print("Login Successfully")

except Exception as e:
    send_email("Login Process Failed", f"Error: {str(e)}")
    raise e

# Navigate to the specified URL after login
report_Quality_url = "https://xxxxxxxxxxxxxx/xxxxxxxx/xxxxxxxx/xxxxxxxxx"
driver.get(report_Quality_url)
time.sleep(2)

report_Quality_url = "https://xxxxxxxxxxxxxxxx/xxxxxxxx/xxxxxxxxx/xxxxxxxxxx"
driver.get(report_Quality_url)
time.sleep(2)


locations = ['Delhi', 'Kolkata']  # Add more locations if needed

for location in locations:

    time.sleep(2)
    
    Process_Type = 'My Process'

    # Determine today's date and the previous date
    today = datetime.now()
    previous_date = today - timedelta(days=1)

    # Locate and interact with elements
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="from_date"]'))).clear()
    time.sleep(2)
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="to_date"]'))).clear()
    time.sleep(2)

    if today.weekday() == 0:  # Monday is 0
        # If it's Monday, process Saturday and Sunday
        saturday_date = today - timedelta(days=2)  # Saturday
        sunday_date = today - timedelta(days=1)  # Sunday
        st_dt = saturday_date.strftime('%m/%d/%Y')
        end_dt = sunday_date.strftime('%m/%d/%Y')
    else:
        # Process for the previous day
        st_dt = previous_date.strftime('%m/%d/%Y')
        end_dt = previous_date.strftime('%m/%d/%Y')

    # Locate and interact with elements for date
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="from_date"]'))).send_keys(st_dt)
    time.sleep(2)
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="to_date"]'))).send_keys(end_dt)
    time.sleep(2)

    # Select the location from the dropdown
    location_select = driver.find_element(By.ID, 'foffice_id')
    for option in location_select.find_elements(By.TAG_NAME, 'option'):
        if option.text == location:
            option.click()
            break

    # Select the Process Type from the dropdown
    Process_Type_select = driver.find_element(By.ID, 'form_new_user')
    for option in Process_Type_select.find_elements(By.TAG_NAME, 'option'):
        if option.text == Process_Type:
            option.click()
            break

    # Trigger the download
    show_button = driver.find_element(By.XPATH, '//*[@id="show"]')
    show_button.click()
    time.sleep(2)

    # Click the Export Report button to download CSV file
    try:
        download_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="main_page_content"]'))
        )
        download_button.click()
    except Exception as e:
        print(f"Error clicking download button: {e}")


    # Wait for the file to be downloaded
    def wait_for_new_file(directory, timeout=1200):
        initial_files = set(os.listdir(directory))
        seconds = 0
        while seconds < timeout:
            current_files = set(os.listdir(directory))
            new_files = current_files - initial_files
            for new_file in new_files:
                if new_file.endswith('.xls') or new_file.endswith('.csv') or new_file.endswith('.xlsx'):  # Check for multiple file types
                    return new_file
            time.sleep(5)
            seconds += 5
        return None

    # Wait for the download to complete
    downloaded_file = wait_for_new_file(download_dir)

    if downloaded_file:
        print(f"Detected new file: {downloaded_file}")
        NewFileName = downloaded_file.replace("'", "")

        # Rename the file based on the location
        new_file_path = os.path.join(download_dir, location + '_' + NewFileName)
        os.rename(os.path.join(download_dir, downloaded_file), new_file_path)
        print(f"File downloaded successfully and renamed to: {new_file_path}")
    else:
        print("File download failed or timed out")

    # Processing CSV Data for each location
    csv_file_path = os.path.join(download_dir, location + '_' + NewFileName)

    print(f"Processing file: {csv_file_path}")

    # Read the CSV with the second row as the header (skip the first row)
    df = pd.read_csv(csv_file_path)  # header=1 means second row as header

    # Check if the DataFrame is empty or has no rows of data
    if df.empty or df.shape[0] == 0:
        # Log and send an email notification for empty file
        send_email(f"No Data in {location} File", f"The downloaded file for {location} is empty. No data to insert into the database.")
        print(f"No data in the {location} downloaded file. Email sent.")
        continue  # Skip the rest of the loop for this location and move on to the next one
    else:
        print(f"Processing file: {csv_file_path}")

        # Now you can check if the first row is blank (even after setting the header to the second row)
        if df.iloc[0].isnull().all() or df.iloc[0].str.strip().eq('').all():
            # Drop the first row and reset header to the second row
            df = pd.read_csv(csv_file_path, header=1)
            print("First row was blank, removed it and set the second row as header.")
        else:
            print("First row is not blank, proceeding with the original header.")

        # Check if DataFrame is empty
        if df.empty:
            # Send an email notification
            send_email(f"No Data in {location} File", f"The downloaded file for {location} is empty. No data to insert into the database.")
            print(f"No data in the {location} downloaded file. Email sent.")
            continue  # Skip to the next location (Delhi or others)
        
        else:
            rename_col = {'Please map your .csv file header with database table header. Like my name : My_name'}
        
            # Rename DataFrame columns using the mapping
            df.rename(columns=rename_col, inplace=True)
            
            # Insert data into the database
            try:
                username = 'xxxxxxxxx' # SQL Database user name
                password = 'xxxxxxxxxxxx' # SQL Dataabse password
                password_encoded = urllib.parse.quote(password)
                server = 'xxx.xxx.xxx.xxx' # SQL Database IP
                database = 'xxxxxxxxxxxx' # SQL Database name

                engine = create_engine(f'mssql+pyodbc://{username}:{password_encoded}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server')
                
                # Insert data into the database
                df.to_sql('table_name', con=engine, if_exists='append', index=False, chunksize=1000) # Put your table name in table_name section

                send_email(f"{location} - Data to SQL Processing Complete", f" {location} - {df.shape[0]} data inserted successfully into the SQL database.")
                print(f"{location} data inserted successfully into the SQL database.")

            except Exception as e:
                send_email(f"Error Inserting  {location} Data into SQL Database", f"Error: {str(e)}")
                raise e
