# Microsoft Learn Student Ambassador Certificate Automation and Email Sender  
This repo simply uses a template certificate docx file and generates certificates in both docx and pdf file formats. It can help you send personalised emails with certificates / LinkedIn Premium Voucher Links attached to it using the MSLA Email template or a Custom email template.   
It can compare names of participants who have completed a Cloud Skills Challenge with your participant data taken using microsoft forms in an excel sheet to only give out the list of active participants.  
It can also help you select winners (randomly) for a lucky draw.  

###  Working on Windows & Chrome only.

### Run these commands in your terminal

#### Clone the GitHub repository and change directory
```
git clone https://github.com/rv2442/MLSA-Certificate-Generator_Email-Sender.git
cd Certificate-Generator-MLSA
```
  
#### Install all needed libraries
```
pip install -r requirements.txt
```
  
### Steps for Cloud Skills Challenges
#### Change directory to .\Filter Participants\ 
```
cd .\Filter Participants\
```

#### Open get_active_participant_list_from_csc.py and edit the following variables
```
excel_path = r"C:\Users\Rahul\OneDriveSky\Desktop\AI900ChallengeParticipantData.xlsx" 

# Name of column with name & email information
name_column = "Full name"
email_column = "Email2"

# Challenge link
challenge_link = "https://learn.microsoft.com/en-us/training/challenges?id=25cae02c-160e-4307-a1f8-08d14a790bc4&WT.mc_id=cloudskillschallenge_25cae02c-160e-4307-a1f8-08d14a790bc4&wt.mc_id=studentamb_248375"

```
  
#### Get active participant list 
This code will open your challenge page using chrome and get the names of all of the participants who have completed the challenge. It will then compare those names against the excel sheet of participants data of the microsoft form where your participants have registered for the challenge. Finally it will find relevant (active participant's) Name & Email data and put it into <e><i>`.\Data\Participant List.csv`</i></e>.  
```
python get_active_participant_list_from_csc.py
```
  
#### Clean and check data mannually for typing errors, irrelevant data in <e><i>`.\Data\Participant List.csv`</i></e>. 

#### Change directory to the previous directory
```
cd ..
```
  
#### Generate certificates for all participants
The code will ask for a prompt asking to run the code in Test Mode (Y/N).  
Press -> N  
```
python main_certificate.py
```
All certificates will be available in `.\Output\Doc\ & .\Output\PDF\`.  
  
### Steps for other Online/Offline Events
Copy your Participant List to the Data Folder and save in the file named as `.\Data\Participant List.csv`. <br>
<e>The list must have the following fields only: <i>```Name,Email```</i></e>. <br>
<e><b>NOTE:</b> The first line of the Participant List.csv file must be ```Name,Email```. Please add your partipant data below this line, do not edit the first line </e>.
```
python main_certificate.py
```  
To use data from ```Data\Participant List.csv``` put N (No) as a prompt when the code asks if you want to run code in test mode.  
Test mode is for developers who wish to tweak code, it has sample data to generate just 1 certificate.

## Customization
- You can change the certificate template file in the `Data` folder.
- You can change the email template in the `Jupyter Notebooks`.

## How to send emails?
- You can use the `Send_certificates_via_email.ipynb` file to send emails with certificates attached to your participants.  
- You can use the `Send_linkedinPremiumVouchers_via_email.ipynb` file to send LinkedIn Premium Vouchers to the Winners.  
- You don't need to change anything in the code logic itself, however you do have to change some data fields such as Event Name, Ambassador Name, Voucher link etc.  
- You need to do is feed Participant Data either by using the scripts in `Filter Participants/` or manually populate the `Participant List.csv` file.  
- Make sure that you have Outlook Installed with the MLSA email id as the primary account, before you run any code that sends emails to your participants.  
- Click on `Run` in the respective Jupyter Notebook blocks.  
- Click on Outbox/Sent Items in Outlook to see the mails being sent one by one.  

## Video Tutorial
Link & Image (yet to be prepared)

### There are 2 email templates:
- MLSA Email Template  
![image](https://github.com/rv2442/MLSA-Certificate-Generator_Email-Sender/assets/69571769/3478c021-c31d-4f67-a031-f2e5d40ddb00)

- Custom Email Template (by Sabyasachi)  
![image](https://github.com/rv2442/MLSA-Certificate-Generator_Email-Sender/assets/69571769/fb7f3429-0250-469e-accb-7387aba04d7d)



Souce Repo: <a href="https://github.com/Sabyasachi-Seal/Certificate-Generator-MLSA">Click Here</a>
