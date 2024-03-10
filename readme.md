# Microsoft Learn Student Ambassador Certificate Automation and Email Sender

This repo simply use a template certificate docx file and generates certificates both docx and pdf. It also sends personalised emails with Certificates / LinkedIn Premium Voucher Links attached to it.

###  Works on Windows only.

## Run these commands in your terminal

```
git clone https://github.com/rv2442/MLSA-Certificate-Generator_Email-Sender.git
cd Certificate-Generator-MLSA
```
Now Copy your Participant List to the Data Folder and rename it as `Participant List.csv`. <br>
<e>The list must have the following fields only: <i>```Email,Name```</i></e>. <br>
<e><b>NOTE:</b> The first line of the Participant List.csv file must be ```Email,Name```. Please add your partipant data below this line, do not edit the first line </e>.
```
pip install -r requirements.txt
python main_certificate.py
```
Use test mode to check if code is functioning properly. Test mode will use ```Data\Temp.csv```.  
To use data from ```Data\Participant List.csv``` put N (No) as a prompt when the code asks if you want to run code in test mode.

## Customization
- You can change the certificate template file in the `Data` folder.
- You can change the email template in the `Jupyter Notebooks`.

## How to send emails?
- You can use the `Send_certificates_via_email.ipynb` file to send emails to the participants.
- You can use the `Send_linkedinPremiumVouchers_via_email.ipynb` file to send LinkedIn Premium Vouchers to the Winners.
- You don't need to change anything in the logic of the code.
- You do need to change data contents within the code for the jupyter notebooks to send certs/vouchers via email. (Fields such as ambassador name, event name etc)
- Then all you need to do is feed Participant Data under ```Participant List.csv``` and then click on ```Run``` in the respective Jupyter Notebook.
- Now open outlook and login.
- Click on outbox and see the mails being sent one by one.

### There are 2 email templates:
- MLSA Email Template  
![image](https://github.com/rv2442/MLSA-Certificate-Generator_Email-Sender/assets/69571769/3478c021-c31d-4f67-a031-f2e5d40ddb00)

- Custom Email Template (by Sabyasachi)  
![image](https://github.com/rv2442/MLSA-Certificate-Generator_Email-Sender/assets/69571769/fb7f3429-0250-469e-accb-7387aba04d7d)


Souce Repo(I made some changes) : <a href="https://github.com/Sabyasachi-Seal/Certificate-Generator-MLSA">Click Here</a>
