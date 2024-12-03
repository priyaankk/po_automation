import streamlit as st
import imaplib
import email
from email.header import decode_header
import os
from groq import Groq
import io
from PIL import Image
import pytesseract
import openpyxl
import requests


import os
os.environ['GROQ_API_KEY'] = 'INSERT API KEY HERE'
import fitz  # PyMuPDF

def generate_csv_export_link(sheet_url):
                # Extract the sheet_id from the Google Sheets URL
                sheet_id = sheet_url.split('/d/')[1].split('/')[0]
                
                # Construct the CSV export URL
                csv_export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
                
                return csv_export_url


def get_google_sheet_data(spreadsheet_url: str, output_format='text'):
    # Define the scope of the Google Sheets API
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets.readonly"]
    
    # Authenticate using the service account credentials
    creds = ServiceAccountCredentials.from_json_keyfile_name('path/to/your/credentials.json', scope)
    
    # Authorize the client
    client = gspread.authorize(creds)
    
    # Open the spreadsheet by URL
    spreadsheet = client.open_by_url(spreadsheet_url)
    
    # Get the first sheet (or specify another sheet name if needed)
    sheet = spreadsheet.sheet1
    
    # Get all values from the sheet (returns a list of rows)
    data = sheet.get_all_values()

    # Optionally print the raw data for inspection
    # print(data)

    if output_format == 'text':
        # Format the data as text
        formatted_data = ""
        for row in data:
            formatted_data += " | ".join(row) + "\n"  # Using "|" as delimiter between columns
        return formatted_data


def extract_pdf_text_from_memory(pdf_doc):
    text = ""
    for page in pdf_doc:
        text += page.get_text()
    return text
# Define the function to analyze email using the LLM
def llama3_agent(q):
    client = Groq(
        api_key=os.environ.get("GROQ_API_KEY"),
    )


    chat_completion = client.chat.completions.create(
        messages=[
            {
                "role": "user",
                "content": q,
            }
        ],
        model="llama3-70b-8192",
    )

    return chat_completion.choices[0].message.content

# Function to fetch the latest email from Gmail
def fetch_latest_email(username, password):
    # Connect to the mail server
    mail = imaplib.IMAP4_SSL("imap.gmail.com")  # Use the appropriate server for your email provider
    mail.login(username, password)

    # Select the mailbox you want to read (INBOX by default)
    mail.select("inbox")

    # Search for all emails in the inbox
    status, messages = mail.search(None, "ALL")

    # Get the list of email IDs
    email_ids = messages[0].split()

    # Get the most recent email (the last email ID in the list)
    latest_email_id = email_ids[-1]

    # Fetch the email by ID
    status, msg_data = mail.fetch(latest_email_id, "(RFC822)")

    # Parse the email content
    for response_part in msg_data:
        attachment_text = None
        attachment_present = False
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])

            # Decode the email subject
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding if encoding else "utf-8")

            # Decode the sender's email
            from_ = msg.get("From")
            message_id = msg["Message-ID"]

            # Extract email body
            body = ""
            
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    pdf_attachments = []
                    image_attachments = []

                    if "attachment" not in content_disposition:
                        if content_type == "text/plain":
                            body = part.get_payload(decode=True).decode()
                    if "attachment" in content_disposition:
                        attachment_present = True
                        filename = part.get_filename()
                        if filename.endswith(".pdf"):
                            pdf_attachments.append(filename)
                            # Save the attachment to the local system
                            payload = part.get_payload(decode=True)

                            # Use BytesIO to load the binary content into memory
                            pdf_data = io.BytesIO(payload)

                            # Open the PDF from the in-memory binary data
                            doc = fitz.open(stream=pdf_data)
                            attachment_text = extract_pdf_text_from_memory(doc)
                            # print("this is the pdf text", attachment_text)
                        if filename.endswith(".png"):
                            image_attachments.append(filename)
                            # Save the attachment to the local system
                            payload = part.get_payload(decode=True)

                            # Use BytesIO to load the binary content into memory
                            image_data = io.BytesIO(payload)

                            # Open the image from the in-memory binary data
                            img = Image.open(image_data)

                            # Use pytesseract to extract text from the image
                            attachment_text = pytesseract.image_to_string(img)

                            # print("This is the image text:", attachment_text)
                        if filename.endswith(".xlsx"):
                            file_payload = part.get_payload(decode=True)
            
                            # Use BytesIO to load the binary content into memory
                            file_data = io.BytesIO(file_payload)

                            # Open the .xlsx file from the in-memory binary data
                            wb = openpyxl.load_workbook(file_data)

                            # Get the first sheet
                            sheet = wb.active

                            # Extract text from the sheet and structure it in a way that the LLM can understand
                            sheet_text = ""
                            for row in sheet.iter_rows(values_only=True):
                                row_text = " | ".join([str(cell) if cell is not None else "" for cell in row])
                                sheet_text += row_text + "\n"

                            # Print the formatted text for inspection
                            attachment_text = sheet_text
                            # print("Formatted text from the .xlsx file:\n", sheet_text)

                        else:
                            pass
            else:
                body = msg.get_payload(decode=True).decode()

            return from_, subject, body, message_id, attachment_text,attachment_present

    mail.close()
    mail.logout()

    return None, None, None, None, attachment_text,attachment_present  # In case no email is found

# Streamlit UI setup
st.title("Purchase Order Extraction Tool")

# Pre-configured email credentials (hardcoded)
username = "xyz@gmail.com"  # Replace with your email
password = ""  # Replace with your password or app password if you have not turned on less secure apps

from_, subject, body, message_id, attachment_text, attachment_present = fetch_latest_email(username, password)

if from_ and subject and body:
    # Display email in a structured format
    st.subheader("📧 Email Details")
    # st.markdown(f"**From**: {from_}")
    # st.markdown(f"**Subject**: {subject}")
    # st.markdown(f"**Body**:\n\n{body}...")  # Display first 1000 characters of the email body
    
    # # Email header section with a more 'email-like' format
    # st.markdown(
    #     f"""
    #     ---
    #     **Subject**: {subject}  
    #     **From**: {from_}  
    #     **Message**: {body}
    #     ---
    #     """
    # )
    
    st.markdown(f"""
    <div style="border: 1px solid #ccc; padding: 10px; border-radius: 5px;">
        <strong>Subject:</strong> {subject}<br>
        <strong>From:</strong> {from_}<br><br>
        <strong>Message:</strong><br>{body} .
    </div>
""", unsafe_allow_html=True)




    # Prepare the prompt for LLM
    system_prompt = """
    You are an AI assistant trained to analyze emails and respond with one of the following options based on the content of the email:

    - "Maybe" if the email has any attachments or links.
    - "Yes" if the email contains product names and quantities in the purchase order details. Do not respond if only 'purchase order' is mentioned or if the details are in a link or does not have quantities.
    - "No" if the email does not contain a purchase order or any attachment.

    Respond Yes, No, or Maybe without apostrophes.
    The mail content is as follows:
    """

    prompt = system_prompt + subject + body

    # Get the LLM response
    response = llama3_agent(prompt)
    if response == "Maybe":
        if not attachment_present:
            response = "No"
            text_to_output = "No purchase order found"
        else:
            text_to_output = "Might be in the attachment"
    if response == "Yes":
        text_to_output = "Purchase order found"
    if response == "No":
        text_to_output = "No purchase order found"
    # st.write(f"**LLM Response**: {text_to_output}")
    st.markdown(f"<span style='color: yellow; '>LLM Response:</span> {text_to_output}", unsafe_allow_html=True)#font-weight: bold;


    # If response is "Yes", ask for elaboration on the purchase order
    if response == "Yes":
        elaborate_prompt = """You need to extract the purchase order and display the captured fields and their values in a user-friendly tabular way. Do not give anything else in the answer."""
        po = llama3_agent(elaborate_prompt + body)
        # st.write(f"**Purchase Order Details**: {po}")
        st.markdown(f"<span style='color: blue; font-weight: bold;'>Purchase Order Details:</span> {po}", unsafe_allow_html=True)

    elif response == "No":
        st.write("No purchase order found.")
    elif response == "Maybe":
        # elaborate_prompt = """You need to extract the purchase order and display the captured fields and their values in a user-friendly tabular way. Do not give anything else in the answer, no preamble no conclusion, just the table(s) with name(s)."""
        elaborate_prompt = """Extract the purchase order details and display the captured fields and their values in a tabular format, with no introductory text or explanations."""
        if attachment_text is None:
            get_link_prompt = f"""
    You are an AI that extracts only the attachment link from the email body.
    Given the email content below, return the link to the attachment, or return 'None' without apostrophes if no attachment link is present:
    """
            attachment_link = llama3_agent(get_link_prompt + body)
            # st.write(f"**Purchase Order**: {attachment_link}")
            check_link = generate_csv_export_link(attachment_link)
            response = requests.get(check_link)
            if response.status_code == 200:
                attachment_text = response.text
                po = llama3_agent(elaborate_prompt + attachment_text)
                # st.write(f"**Purchase Order Details**: {po}")
                st.markdown(f"<span style='color: yellow; '>Purchase Order Details:</span> {po}", unsafe_allow_html=True) #font-weight: bold;

        else:
            po = llama3_agent(elaborate_prompt + attachment_text)
            # st.write(f"**Purchase Order Details**: {po}")
            st.markdown(f"<span style='color: yellow; font-weight: bold;'>Purchase Order Details:</span> {po}", unsafe_allow_html=True)

else:
    st.write("No emails found.")
