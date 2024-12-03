# Email Purchase Order Classifier

## Overview

This project reads your most recent email and classifies it based on the presence of purchase order information. The classification categories are:

- **Yes**: Purchase order info in the email body.
- **Maybe**: Email contains an attachment.
- **None**: No purchase order info found.

Streamlit is used for the frontend.

## Setup

1. **GROQ API Key**: Obtain a free GROQ API key.
2. **Google Account**: 
   - Enable less secure apps in Google Admin Console or
   - Use an App Password from Google for better security.
3. **Install Dependencies**: Run the following to install required libraries:

   ```bash
   pip install -r requirements.txt
   ```

## Running the Project

To run the app:

```bash
streamlit run app.py
```
