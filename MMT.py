import pandas as pd
import datetime
from datetime import datetime
import time
import random
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from Emailoperations import get_email_attachments
from Emailoperations import send_email_attachment

# Reading Config
try:
    df = pd.read_excel(r"D:\New folder\TestCase\Config\Config.xlsx")
    if 'Key' in df.columns and 'file' in df.columns:
        df.dropna(subset=["Key","file"], inplace=True)
        config = df.set_index('Key')['file'].to_dict()
    else:
        print("Config file missing 'Key' or 'file' columns.")
        print("Found columns:", df.columns)
        
except Exception as e:
    print(f"Error reading Config: {e}")
    
print("Config loaded successfully.")

# Reading Email
get_email_attachments(
    username = config["username"],
    password =  config["password"],
    sender_filter=config["MailID"],
    subject_filter=config["SeachSubject"],
    file_extension=".xlsx",
    unread_only=False,
    save_dir=config["InputFolder"] 
    )
print("Successfully Fecthed The Attachment File")

# Reading InputFile
inputFile = pd.read_excel(config["InputFilePath"])
print(inputFile)

# Declaring Output DF
output_columns = [
            "Traveling_Date",
            "Traveling_Country",
            "Airline_Name",
            "Flight_Available",
            "Lowest_Fare",
            "Customer_Email_ID",
            "Customer Name"
        ]
output_df = pd.DataFrame(columns=output_columns)

# Setup Undetected Chrome
options = uc.ChromeOptions()
options.add_argument("--start-maximized")

# Initialize Driver
driver = uc.Chrome(options=options, use_subprocess=True)
wait = WebDriverWait(driver, 20)
try:
    url = config["url"]
    print(f"Navigating to {url}...")
    driver.get(url)
    
    # Human-like pause
    time.sleep(random.uniform(3, 5)) # Increased wait
    # 1. Handle potential login modal / ads
    try:
        # Click on body to dismiss generic overlays if they exist
        print("Clicking body to close initial modal...")
        driver.find_element(By.TAG_NAME, "body").click()
        time.sleep(1)
        
        # Explicit close button check
        close_btns = driver.find_elements(By.CSS_SELECTOR, ".commonModal__close, span[data-cy='modalClose'], .close, div.loginModal>section>span")
        if close_btns:
            print("Found explicit close button, clicking it...")
            driver.execute_script("arguments[0].click();", close_btns[0])
            print("Closed modal.")
            time.sleep(1)
        else:
            print("No explicit close button found after body click.")
    except Exception as e:
        print(f"Error closing initial modal: {e}")
        pass

    time.sleep(2)
    from_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[@class='coachmark']")))
    from_label.click()
    
    for i in range(len(inputFile)):
        
        try:
            # Enter 'From' City 
            time.sleep(3)
            print("Entering From City...")
            from_label = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='fromCity']")))
            from_label.click()
            time.sleep(1)
            from_input = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='From']")))
            for char in str(inputFile["Traveling_From"][i]):
                from_input.send_keys(char)
                time.sleep(random.uniform(0.05, 0.2))

            time.sleep(2) 
            # Select the suggestion 
            print("Selecting suggestion...")
            suggestion_from = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='react-autowhatever-1-section-0-item-0']")))
            suggestion_from.click()
            time.sleep(1)

            # Enter 'To' City 
            print("Entering To City...")
            to_label = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='toCity']")))
            to_label.click()
            time.sleep(1)
            to_input = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='To']")))
            for char in str(inputFile["Traveling_Country_to"][i]):
                to_input.send_keys(char)
                time.sleep(random.uniform(0.05, 0.2))

            time.sleep(2) 
            # Select the suggestion 
            print("Selecting suggestion...")
            suggestion_to = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='react-autowhatever-1-section-0-item-0']")))
            suggestion_to.click()
            time.sleep(1)

            # Select Date
            input_date = str(inputFile["Traveling_Date"][i]).split(" ")[0]

            date_obj = datetime.strptime(input_date, "%Y-%m-%d")

            date_str = date_obj.strftime("%a %b %d %Y")  # Sat Dec 20 2025
            Chechk_date_str = date_obj.strftime("%B %Y")

            while True:
                month = driver.find_element(By.XPATH, "(//div[@class='DayPicker-Caption']/div)[1]").text
                date_obj2 = datetime.strptime(month, "%B %Y")


                if date_obj > date_obj2 and Chechk_date_str.strip() != month.strip():
                    print("Clicking Next Month")
                    next_btn = driver.find_element(By.XPATH, "//span[@class='DayPicker-NavButton DayPicker-NavButton--next']")
                    next_btn.click()
                    time.sleep(2)
                elif date_obj < date_obj2 and Chechk_date_str.strip() != month.strip():
                    print("Clicking Previous Month")
                    prev_btn = driver.find_element(By.XPATH, "//span[@class='DayPicker-NavButton DayPicker-NavButton--prev']")
                    prev_btn.click()
                    time.sleep(2)
                else:
                    print("Month Found")
                    break
                
            print(f"Selecting Date: {date_str}...")
            try:
                date_element = driver.find_element(By.CSS_SELECTOR, f"div[aria-label='{date_str}']")
            except:
                print("Opening calendar...")
                dept_label = driver.find_element(By.ID, "departure") 
                driver.execute_script("arguments[0].click();", dept_label)
                time.sleep(1)
                date_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"div[aria-label='{date_str}']")))

            date_element.click()
            time.sleep(1)
            
            # Click Search
            print("Clicking Search...")
            search_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.widgetSearchBtn")))
            driver.execute_script("arguments[0].click();", search_btn)

            print("Search initiated. Waiting for results...")
            wait.until(EC.url_contains("search"))
            print("Results page loaded successfully!")

            # Handle Post-Search Ads/Modals
            print("Waiting for post-search ads to stabilize...")
            time.sleep(5) # Wait for potential interstitial/ad
            try:
                print("Checking for post-search overlays...")
                btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='overlayCrossIconV2 newCrossIconV2Wrapper']")))
                btn.click()
                time.sleep(1)

                # Click body again just in case
                driver.find_element(By.TAG_NAME, "body").click()

            except Exception as e:
                print(f"Post-search ad handling warning")


            # Extract Flight Results
            print("Waiting for flight list to load...")
            time.sleep(5) # explicit wait for render

            print("Extracting Data...")

            # Airline Name
            flight_elements = wait.until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//p[@data-test='component-airlineHeading']")
                )
            )
            flight_names = []
            for ele in flight_elements:
                name = ele.text.strip()
                flight_names.append(name)

            # Price
            flight_elements = wait.until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@data-test='component-fare']//following::span[@class=' fontSize18 blackFont']")
                )
            )
            flight_Price = []
            for ele in flight_elements:
                name = ele.text.strip()
                flight_Price.append(name)

            rows = []
            for j in range(len(flight_names)-2):
                rows.append({
                    "Flight Name": flight_names[j],
                    "Price": flight_Price[j]
                })

            print(f"Data fetched for Customer: {i+1}")

            df = pd.DataFrame(rows)
            df["Price_num"] = (
                df["Price"]
                .str.replace("₹", "", regex=False)
                .str.replace(",", "", regex=False)
                .str.strip()
                .astype(int)
            )

            min_row = df.loc[df["Price_num"].idxmin()]

            min_df = pd.DataFrame([{
                "Traveling_Date": input_date,
                "Traveling_Country": inputFile["Traveling_Country_to"][i],
                "Airline_Name": min_row["Flight Name"],
                "Flight_Available": "Yes",
                "Lowest_Fare": min_row["Price"],
                "Customer_Email_ID": "NaN",
                "Customer Name": inputFile["Customer Name"][i]
            }])

            if output_df.empty:
                output_df = min_df.copy()
            else:
                output_df = pd.concat([output_df, min_df], ignore_index=True)

            driver.back()
        except:
            min_df = pd.DataFrame([{
                "Traveling_Date": input_date,
                "Traveling_Country": inputFile["Traveling_Country_to"][i],
                "Airline_Name": "NaN",
                "Flight_Available": "No",
                "Lowest_Fare": "NaN",
                "Customer_Email_ID": "NaN",
                "Customer Name": inputFile["Customer Name"][i]
            }])
            
            if output_df.empty:
                output_df = min_df.copy()
            else:
                output_df = pd.concat([output_df, min_df], ignore_index=True)

            driver.back()
            
    print(output_df)
    output_df.to_excel(config["OutputFilePath"], index=False)
    
    # Sending Email Attachment
    send_email_attachment(
        sender_email = config["username"],
        sender_password = config["password"],
        receiver_email = config["SendMailID"],
        subject = config["SendMailSubject"],
        bodyPara = config["SendMailBody"],
        attachments = config["OutputFilePath"]  
    )

except Exception as e:
    print(f"An error occurred: {e}")
    driver.save_screenshot(config["ErrorScreenshot"])
finally:
    try:
        input("Enter to continue")
        driver.quit()
    except:
        pass
