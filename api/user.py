import gspread
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import time
from transformers import pipeline, AutoTokenizer, AutoModelForQuestionAnswering
from nltk.stem import WordNetLemmatizer
import os
from os.path import dirname, abspath, join
import nltk
# Download NLTK data
nltk.download('wordnet')


# Function to process user input
def process_user_input(Content, Shipping_Value):

    # Assuming the script is in the project root
    script_dir = dirname(abspath(__file__))
    file_path = join(script_dir, '..', 'data', 'HS_Code_Refined_Data.csv')
    df = pd.read_csv(file_path)
    df = df.dropna(subset=['HS Code', 'Item Description', 'Basic Duty (SCH)', 'IGST', '10% SWS', 'Total duty with SWS of 10% on BCD'])
    HS_code = df['HS Code'].tolist()
    item_descriptions = df['Item Description'].tolist()
    basic_duty = df['Basic Duty (SCH)'].tolist()
    igst = df['IGST'].tolist()
    sws_10_percent = df['10% SWS'].tolist()
    total_duty_with_sws = df['Total duty with SWS of 10% on BCD'].tolist()

    # Load pre-trained question-answering model (assuming this part remains unchanged)
    model_name = "bert-large-uncased-whole-word-masking-finetuned-squad"
    tokenizer = AutoTokenizer.from_pretrained(model_name)
    model = AutoModelForQuestionAnswering.from_pretrained(model_name)
    qa_pipeline = pipeline('question-answering', model=model, tokenizer=tokenizer)

    # Lemmatize function to get the base form of a word
    lemmatizer = WordNetLemmatizer()
    # Lemmatization of user input using WordNet lemmatizer
    Content_lemmatized = lemmatizer.lemmatize(Content.lower())

    # Process user input
    for i in range(len(item_descriptions)):
        item_description_lemmatized = lemmatizer.lemmatize(item_descriptions[i].lower())

            # Check if lemmatized user input is a substring of lemmatized item description
        if Content_lemmatized in item_description_lemmatized:
            question = f"What is the HS code and Total duty for {item_descriptions[i]}?"
            context = f"HS Code: {HS_code[i]}, Item Description: {item_descriptions[i]}, Basic Duty: {basic_duty[i]}, IGST: {igst[i]}, 10% SWS: {sws_10_percent[i]}, Total duty: {total_duty_with_sws[i]}"

            # Using the pipeline to get the answer
            answer = qa_pipeline(question=question, context=context)

            # Extracting relevant information from the answer
            hs_code_info = HS_code[i].replace(" ", '')
            total_duty_info = total_duty_with_sws[i]
            total_value_info = total_duty_info / Shipping_Value

            print(f"[+] HS Code Info: {hs_code_info}")
            print(f"[+] Total Duty Info: {total_duty_info}")
            print(f"[+] Total Value Info: {total_value_info}\n")
            return f"\n [+] HS Code Info: {hs_code_info}\n [+] Total Duty Info: {total_duty_info}\n [+] Total Value Info: {total_value_info}"
    return  f'[*] Sorry AI Fails Here Due to Less Human intervention on tarning Data! \n [!] Please provide more information of {Content}'



def initialize_google_sheet():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    script_dir0 = dirname(abspath(__file__))
    file_path0 = join(script_dir0, '..', 'data', 'innate-booking-399709-4ca128908c3c.json')
    gc = gspread.service_account(filename=file_path0)
    sh_glide = gc.open_by_key('1TgA5YVucHUnqm0t-X2763sJQUauUAE90CDg-0V-IwfU')
    worksheet = sh_glide.worksheet("Sheet1")
    value_list = worksheet.get_all_records()
    return worksheet, value_list


def get_rates(weight_inputs):
    # Load Excel data
    script_dir1 = dirname(abspath(__file__))
    file_path1 = join(script_dir1, '..', 'data', 'MyShippingGenie_ShippingRate_Sample.xlsx')
    shipping_df = pd.read_excel(file_path1)
    matching_rows = shipping_df[shipping_df['Weight  (lbs)'] == weight_inputs]

    #rate = None

    if not matching_rows.empty:
        print(f"Shipping rates for {weight_inputs} lbs:")
        for index, row in matching_rows.iterrows():
            from_country = row['From Country']
            to_country = row['To Country']
            rate = row['Rate (USD)']
            #print(f"From {from_country} to {to_country}: ${rate}")
        return f'From {from_country} to {to_country} rate is ${rate}'  # Return rate after the loop





def send_email(email, message):
    email_smtp = 'smtp.gmail.com'
    email_login = 'amanaman68499@gmail.com'
    email_password = 'cyzk ddyt nzei eulk'

    msg = MIMEMultipart()
    msg['From'] = email_login
    msg['To'] = email
    msg['Subject'] = 'Your Instant rates'
    msg.attach(MIMEText(message, 'plain'))

    # Connect to the SMTP server
    server = smtplib.SMTP(email_smtp, 587)
    server.starttls()  # Use TLS encryption

    # Login to the email account
    server.login(email_login, email_password)

    # Send the email
    server.sendmail(email_login, email, msg.as_string())

    # Close the connection
    server.quit()

def process_google_sheet():
    worksheet, value_list = initialize_google_sheet()

    # Initialize the last processed index
    last_processed_index = len(value_list)

    while True:
        print("Getting the latest data from the Google Sheet...")
        updated_value_list = worksheet.get_all_records()
        print("Latest data retrieved.")

        # Check for new entries starting from the last processed index
        new_entries = updated_value_list[last_processed_index:]

        # Process new entries
        for entry in new_entries:
            email = entry['your_email']
            weight_inputs = entry['Weight_inputs']
            Content = entry['Content']
            Shipping_Value = entry['Shipping_Value']
            # Ensure the required fields are not empty
            if email and Content and Shipping_Value and weight_inputs:
                print(f"Processing entry for {email} and {weight_inputs}...")
                user_prediction_1 = get_rates(weight_inputs)
                print(f"Got rates: {user_prediction_1}")
                print(f"Processing entry for {email} , {Content} and {Shipping_Value}...")
                user_prediction_2 = process_user_input(Content, Shipping_Value)
                print(f"Got Custom Duty info: {user_prediction_2}")
                row_number_1 = updated_value_list.index(entry) + 2
                worksheet.update_cell(row_number_1, 3, user_prediction_1)
                row_number_2 = updated_value_list.index(entry) + 2
                worksheet.update_cell(row_number_2, 6, user_prediction_2)
                message = f"Here is your Rate and Custom Duty Info ({Content} and Shipping Value {Shipping_Value}):- \n [+]  ${user_prediction_1:} for a weight {weight_inputs}kg \n {user_prediction_2}"
                send_email(email, message)
                print("Email sent.")
                last_processed_index = row_number_1 - 1  # Adjusting for 0-based indexin
                last_processed_index = row_number_2 - 1  # Adjusting for 0-based indexin




            elif email and weight_inputs:
                print(f"Processing entry for {email} and {weight_inputs}...")
                user_prediction = get_rates(weight_inputs)
                print(f"Got rates: {user_prediction}")
                row_number = updated_value_list.index(entry) + 2
                worksheet.update_cell(row_number, 3, user_prediction)
                message = f'Here is your Rate: ${user_prediction:} for a weight of {weight_inputs}kg'
                send_email(email, message)
                print("Email sent.")
                # Update the last processed index
                last_processed_index = row_number - 1  # Adjusting for 0-based indexing

            elif email and Content and Shipping_Value :
                print(f"Processing entry for {email} , {Content} and {Shipping_Value}...")
                user_prediction = process_user_input(Content, Shipping_Value)
                print(f"Got Custom Duty info: {user_prediction}")
                row_number = updated_value_list.index(entry) + 2
                worksheet.update_cell(row_number, 6, user_prediction)
                message = f'Here is your Custom Duty Info for {Content} and Shipping Value {Shipping_Value} :\n{user_prediction}'
                send_email(email, message)
                print("Email sent.")
                # Update the last processed index
                last_processed_index = row_number - 1  # Adjusting for 0-based indexing



        # Adjust the sleep time as needed
        time.sleep(3)

# Call the function to start the process
process_google_sheet()
