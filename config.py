import gspread
from google.oauth2.service_account import Credentials
sheet_id = "1il6yvUxbZt1Oq2aLeFGvrNnGnyFUvyQefngoj1JZSH8"
scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
creds = Credentials.from_service_account_file("../credentials.json", scopes=scopes)
client = gspread.authorize(creds)