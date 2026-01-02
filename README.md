# Email-Based Travel Inquiry Automation (Python)

## 📌 Overview

This project is a Python-based automation solution designed to process travel inquiry requests received through Gmail. Each inquiry email contains an Excel attachment with customer travel details. The automation reads these emails, extracts the required data, searches for flight details on MakeMyTrip, and sends the processed results back to the customer in a new Excel file.

The project demonstrates real-world automation concepts such as email handling, Excel processing, web automation, and end-to-end workflow execution using Python.

---

## 🎯 Use Case

Travel agencies or service providers often receive multiple travel inquiries via email. Manually processing these requests is time-consuming and error-prone. This automation eliminates manual effort by:

* Reading travel inquiries automatically
* Extracting customer requirements
* Fetching live flight details
* Sending structured results back to customers

---

## ⚙️ Functional Flow

1. Read unread emails from Gmail with the subject **“Travel Inquiry Request”**
2. Download the attached Excel file from the email
3. Extract the following details from the Excel file:

   * Traveling Date
   * Traveling Country
   * Customer Email ID
4. Search MakeMyTrip using the extracted details to retrieve:

   * Airline Name
   * Flight Availability
   * Lowest Fare
5. Create a new Excel file with the processed flight details
6. Send the output Excel file back to the customer via Gmail

---

## 🛠️ Technologies Used

* **Python**
* **IMAP & SMTP** (Email operations)
* **Pandas** (Excel reading and writing)
* **Selenium** (Web automation)
* **OpenPyXL / Excel Libraries**
* **Undetected Chrome WebDriver**

---

## 📥 Input

* Gmail email with subject: **Travel Inquiry Request**
* Excel attachment containing:

  * Traveling Date
  * Traveling Country
  * Customer Email ID

---

## 📤 Output

* A newly generated Excel file containing:

  * Airline Name
  * Flight Availability
  * Lowest Fare
* Output file is emailed back to the customer automatically

---

## ▶️ How to Run

1. Clone the repository:

   ```bash
   git clone https://github.com/Arunskr01/Email-Based-Travel-Inquiry-Automation-Python-.git
   ```
2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```
3. Configure Gmail credentials (App Password recommended)
4. Ensure ChromeDriver is installed and added to PATH
5. Run the automation:

   ```bash
   python main.py
   ```

---

## 🔐 Notes

* Use **Gmail App Passwords** instead of normal passwords
* Web automation depends on MakeMyTrip UI; changes in UI may require XPath updates
* Intended for educational and automation practice purposes

---

## 🚀 Key Learnings

* Email automation using Python
* Excel data handling with Pandas
* Web scraping and automation using Selenium
* End-to-end process automation design

---

## 📄 License

This project is for learning and demonstration purposes.

---
