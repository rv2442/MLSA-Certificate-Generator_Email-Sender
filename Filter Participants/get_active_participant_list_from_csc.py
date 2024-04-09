# Global imports
import time
import sys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup

# Local Module imports
from get_active_participant_list_from_excel import get_active_participants_from_excel


# Setting variables

#Path to the excelfile having data of the microsoft form that your participants have filled to register for the cloud skills challenge 
excel_path = r"C:\Users\Rahul\OneDriveSky\Desktop\AI900ChallengeParticipantData.xlsx" 

# Name of column with name & email information
name_column = "Full name"
email_column = "Email2"

# Path of the file where final list of participants who completed the challenge will be stored
output_path = r"../Data/Participant List.csv"

# Challenge link
challenge_link = "https://learn.microsoft.com/en-us/training/challenges?id=25cae02c-160e-4307-a1f8-08d14a790bc4&WT.mc_id=cloudskillschallenge_25cae02c-160e-4307-a1f8-08d14a790bc4&wt.mc_id=studentamb_248375"

# Configure Selenium options
options = Options()
options.add_argument('--headless')  # Run Chrome in headless mode, no GUI needed

driver = webdriver.Chrome()

# Open challenge link in chrome
driver.get(challenge_link)

# Wait for the leaderboard to load (adjust the wait time as needed)
time.sleep(7)

# Get the updated page source after dynamic loading
page_source = driver.page_source

# Parse the HTML content using BeautifulSoup
soup = BeautifulSoup(page_source, "html.parser")

# Find the div with id "leaderboard-list"
leaderboard_div = soup.find("div", id="leaderboard-list")

# Initilaize participant list
participant_list = []

def extract_active_participants(leaderboard_items):

    for item in leaderboard_items:
        # Extract user's name
        user_name = item.find("span", class_="leaderboard-name").text.strip()
        
        # Extract user's score
        user_score = item.find("span", class_="visually-hidden").text.strip().split(",")[2]

        # Extract user's no of modules completed
        no_of_modules_completed = int(item.find("span", class_="visually-hidden").text.strip().split(",")[2].split('/')[0])

        # # Print the user's name and modules completed
        # print(f"User: {user_name}, Modules Completed: {no_of_modules_completed}/{total_no_of_modules}")
        
        # If user has completed all modules add him to participant list
        if no_of_modules_completed == total_no_of_modules:
            participant_list.append(user_name)

        else:
            # close chrome
            driver.quit()

            # compare and map names from excel sheet and save only matching names to output file
            get_active_participants_from_excel(excel_path, name_column, email_column, participant_list, output_path)

            print("Active Participant list extracted successfully")
            print(f"Please check it at: {output_path}")

            sys.exit()
        
            

if leaderboard_div:

    # Find all list items within the leaderboard div
    leaderboard_items = leaderboard_div.find_all("li", class_="is-unstyled")

    # Total No of Modules
    total_no_of_modules = int(leaderboard_items[0].find("span", class_="visually-hidden").text.strip().split(",")[2].split('/')[1].split(' ')[0])

    # Find the button for the first page
    first_page_button = soup.find("button", class_="pagination-link is-current")

    # Extract total number of pages from the first page button
    total_pages = int(first_page_button["aria-label"].split()[-1])

    # Extract active participants from the first page   
    extract_active_participants(leaderboard_items)

    for page in range(2, total_pages + 1):
        
        # Find the next page button
        next_page_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.pagination-next[aria-label='Next']"))
        )

        # Click the next page button
        next_page_button.click()

        # Wait for the next page to load, load time may differ based on internet speed
        time.sleep(7)

        # Get the updated page source after dynamic loading
        page_source = driver.page_source

        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(page_source, "html.parser")

        # Find the div with id "leaderboard-list"
        leaderboard_div = soup.find("div", id="leaderboard-list")

        # Find all list items within the leaderboard div
        leaderboard_items = leaderboard_div.find_all("li", class_="is-unstyled")

        # Extract active participants from the current page
        extract_active_participants(leaderboard_items)

else:
    print("Leaderboard div not found or browser session expired.")
    driver.quit()