import pandas as pd
from fuzzywuzzy import fuzz

def get_active_participants_from_excel(excel_path, name_column, email_column, participant_list, output_path):

    # Read the Excel file
    excel_data = pd.read_excel(excel_path)

    # Change all data to lowercase for better similarity ratios
    excel_data = excel_data.applymap(lambda x: x.lower() if isinstance(x, str) else x)
    excel_data = excel_data.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Drop duplicate entries based on name
    unique_data = excel_data.drop_duplicates(subset=[name_column])

    # Reset index to ensure sequential indices
    unique_data.reset_index(drop=True, inplace=True)

    # Drop duplicate entries based on email
    unique_excel_data = unique_data.drop_duplicates(subset=[email_column])

    # Reset index to ensure sequential indices
    unique_excel_data.reset_index(drop=True, inplace=True)

    # Extract names and emails from the Excel data
    excel_names = unique_excel_data[name_column]
    excel_emails = unique_excel_data[email_column]

    # List to store matched names and emails
    matched_participants = []

    # Iterate over each row in the Excel data
    for i, name in enumerate(excel_names):
        # Iterate over each participant extracted from the challenge page
        for participant in participant_list:
            # Compare the names using fuzzy string matching (Token Set Ratio)
            similarity_ratio = fuzz.token_set_ratio(name, participant)
            # If similarity ratio is above the threshold, consider it a match
            if similarity_ratio >= 75:
                # Store the matched name and email
                matched_participants.append({"Name": name, "Email": excel_emails[i]})
                break  # Move to the next name in the Excel data

    # Remove duplicate entries from the list of matched participants
    matched_participants_unique = [dict(t) for t in {tuple(d.items()) for d in matched_participants}]

    # Open a file for writing
    with open(output_path, "w") as file:
        # Write the header
        file.write("Name,Email\n")
        
        # Write each participant's name and email in the specified format
        for participant in matched_participants:
            file.write(f"{participant['Name'].title()},{participant['Email']}\n")
        
        file.close()
    
