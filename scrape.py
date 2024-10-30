from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import policy
from email.generator import BytesGenerator
import time

# Set up Chrome WebDriver using WebDriver Manager
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# URL of the Netskope Trust Portal
URL = "https://trust.netskope.com/#"

# Define keywords to identify Australian locations
AUSTRALIAN_LOCATIONS = ["SYD1", "SYD2", "MEL1", "MEL2", "Brisbane", "Perth"]

# Function to navigate to the Maintenance tab and extract relevant maintenance information
def check_australian_maintenance():
    # Load the Netskope Trust Portal
    driver.get(URL)
    time.sleep(5)  # Allow time for the page to fully load

    # Navigate to the "Maintenance" tab
    try:
        maintenance_tab = driver.find_element(By.XPATH, "//button[contains(text(),'Maintenance')]")
        maintenance_tab.click()
        print("Navigated to the Maintenance tab.")
        time.sleep(3)  # Wait for the maintenance tab content to load
    except Exception as e:
        print(f"Failed to navigate to the Maintenance tab: {e}")
        driver.quit()
        return []

    # Parse the page source to extract maintenance information
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    
    # Locate specific maintenance entries - Adjust based on actual HTML structure
    maintenance_sections = soup.find_all("div", class_="tabcontent")  # Use more specific selectors if possible
    
    maintenance_records = []

    # Check each maintenance entry for mentions of Australian locations
    for maintenance in maintenance_sections:
        # Filter out only relevant text based on presence of Australian location keywords
        maintenance_text = maintenance.get_text()
        
        if any(location in maintenance_text for location in AUSTRALIAN_LOCATIONS):
            # Extract more structured data (title, status, date, details) if available
            title = maintenance.find("h2").get_text() if maintenance.find("h2") else "Maintenance - Australian Data Center"
            status = maintenance.find("strong").get_text() if maintenance.find("strong") else "Scheduled"
            date_time = maintenance.find("time").get_text() if maintenance.find("time") else "N/A"
            
            # Refine description to exclude unrelated data
            description = "\n".join([line for line in maintenance_text.splitlines() if any(loc in line for loc in AUSTRALIAN_LOCATIONS) or "Scheduled" in line or "Completed" in line or "In progress" in line])
            duration = "24 hours"  # Example placeholder duration; adjust as needed
            
            # Add the entry to the list
            maintenance_records.append({
                "Title": title,
                "Status": status,
                "Date and Time": date_time,
                "Duration": duration,
                "Details": description
            })

    return maintenance_records

# Function to create and save an EML file with the maintenance details
def create_eml(maintenance_records):
    # Email setup
    email_message = MIMEMultipart()
    email_message['From'] = "alerts@netskope.com"
    email_message['To'] = "recipient@example.com"
    email_message['Subject'] = "Upcoming Maintenance for Australian Sites"

    # Email body content
    body_content = "<h2>Upcoming Maintenance for Australian Sites</h2>"
    body_content += "<p>Here are the upcoming maintenance windows affecting Australian locations:</p>"
    
    if maintenance_records:
        # Format each maintenance entry as HTML to resemble the structured format
        for record in maintenance_records:
            body_content += f"""
            <div style="margin-bottom: 20px;">
                <h3 style="color: #0073e6;">{record['Title']}</h3>
                <p><strong>Status:</strong> {record['Status']}<br>
                <strong>Date and Time:</strong> {record['Date and Time']}<br>
                <strong>Estimated Duration:</strong> {record['Duration']}<br>
                <strong>Details:</strong><br> {record['Details']}<br>
                </p>
            </div>
            <hr>
            """
    else:
        body_content += "<p>No upcoming maintenance windows found for Australian locations.</p>"

    # Attach HTML body content
    email_message.attach(MIMEText(body_content, "html"))

    # Save the EML file
    with open("Australian_Maintenance_Schedule.eml", "wb") as eml_file:
        generator = BytesGenerator(eml_file, policy=policy.default)
        generator.flatten(email_message)
    print("Maintenance details exported to Australian_Maintenance_Schedule.eml")

# Run the maintenance check and create an EML file
try:
    # Check for maintenance entries affecting Australian locations
    australian_maintenance = check_australian_maintenance()

    # Create an EML file with the maintenance records
    create_eml(australian_maintenance)

finally:
    driver.quit()  # Ensure the browser closes when the script ends
    print("Driver closed and script terminated.")
