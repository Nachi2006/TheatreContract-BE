from fastapi import FastAPI, Depends, HTTPException, status, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from jose import JWTError, jwt
from passlib.context import CryptContext
from datetime import datetime, timedelta
import pandas as pd
import json
import os
import io
import zipfile
from typing import List, Optional
from pydantic import BaseModel
from dotenv import load_dotenv

load_dotenv()

SECRET_KEY = os.getenv("SECRET_KEY")
ALGORITHM = os.getenv("ALGORITHM")
ACCESS_TOKEN_EXPIRE_MINUTES = int(os.getenv("ACCESS_TOKEN_EXPIRE_MINUTES"))
USERS_FILE = os.getenv("USERS_FILE")
origins = [os.getenv("VERCEL_LINK")]

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token")

class User(BaseModel):
    username: str
    is_admin: bool

class UserCreate(BaseModel):
    username: str
    password: str
    is_admin: bool = False

def load_users():
    if not os.path.exists(USERS_FILE):
        default_user = os.getenv("ADMIN_DEFAULT_USER")
        default_pass = os.getenv("ADMIN_DEFAULT_PASSWORD")
        return {default_user: {"password": pwd_context.hash(default_pass), "is_admin": True}}
    with open(USERS_FILE, 'r') as f:
        return json.load(f)

def save_users(users):
    with open(USERS_FILE, 'w') as f:
        json.dump(users, f, indent=4)

def create_access_token(data: dict):
    to_encode = data.copy()
    expire = datetime.utcnow() + timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)

async def get_current_user(token: str = Depends(oauth2_scheme)):
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        username: str = payload.get("sub")
        if username is None:
            raise HTTPException(status_code=401, detail="Invalid credentials")
        users = load_users()
        if username not in users:
            raise HTTPException(status_code=401, detail="User not found")
        return {"username": username, "is_admin": users[username]["is_admin"]}
    except JWTError:
        raise HTTPException(status_code=401, detail="Invalid token")

@app.post("/token")
async def login(form_data: OAuth2PasswordRequestForm = Depends()):
    users = load_users()
    user = users.get(form_data.username)
    if not user or not pwd_context.verify(form_data.password, user["password"]):
        raise HTTPException(status_code=400, detail="Incorrect username or password")
    
    access_token = create_access_token(data={"sub": form_data.username})
    return {"access_token": access_token, "token_type": "bearer", "is_admin": user["is_admin"]}

def create_excel_with_totals(df: pd.DataFrame, output_io: io.BytesIO, screen_column_name: str):
    target_col_indices = [14, 16]
    
    target_cols = [df.columns[i] for i in target_col_indices if i < len(df.columns)]
    
    with pd.ExcelWriter(output_io, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Summary')
        
        workbook = writer.book
        worksheet = writer.sheets['Summary']
        
        start_row = len(df) + 2
        
        if screen_column_name in df.columns and target_cols:
            screen_totals = df.groupby(screen_column_name)[target_cols].sum().reset_index()
            
            worksheet.write_string(start_row, 0, f"{screen_column_name} Totals")
            
            current_row = start_row + 1
            
            for index, row in screen_totals.iterrows():
                worksheet.write_string(current_row, 0, str(row[screen_column_name]))
                
                for col_name in target_cols:
                    col_pos = df.columns.get_loc(col_name)
                    worksheet.write_number(current_row, col_pos, row[col_name])
                    
                current_row += 1
            
            grand_total_row = current_row + 1
            worksheet.write_string(grand_total_row, 0, "Grand Total")
            
            for col_name in target_cols:
                col_pos = df.columns.get_loc(col_name)
                grand_total_val = screen_totals[col_name].sum()
                worksheet.write_number(grand_total_row, col_pos, grand_total_val)


@app.post("/extract-theatres")
async def extract_theatres(file: UploadFile = File(...), current_user: dict = Depends(get_current_user)):
    contents = await file.read()
    df = pd.read_excel(io.BytesIO(contents), header=3)
    theatres = df['Theatre Name'].dropna().unique().tolist()
    return {"theatres": theatres}

@app.post("/generate-custom-excel")
async def generate_custom_excel(
    file: UploadFile = File(...),
    selected_theatres: str = Form(...),
    screen_column_name: str = Form("Screen Name"),
    current_user: dict = Depends(get_current_user)
):
    theatres_list = json.loads(selected_theatres)
    contents = await file.read()
    df = pd.read_excel(io.BytesIO(contents), header=3)
    
    filtered_df = df[df['Theatre Name'].isin(theatres_list)]
    
    output = io.BytesIO()
    create_excel_with_totals(filtered_df, output, screen_column_name)
    output.seek(0)
    
    return StreamingResponse(
        output, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=selected_theatres_summary.xlsx"}
    )

@app.post("/generate-all-zip")
async def generate_all_zip(
    file: UploadFile = File(...),
    screen_column_name: str = Form("Screen Name"),
    current_user: dict = Depends(get_current_user)
):
    contents = await file.read()
    df = pd.read_excel(io.BytesIO(contents), header=3)
    
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for theatre_name in df['Theatre Name'].dropna().unique():
            theatre_df = df[df['Theatre Name'] == theatre_name]
            
            if theatre_df.empty:
                continue
                
            output = io.BytesIO()
            create_excel_with_totals(theatre_df, output, screen_column_name)
            
            zipf.writestr(f"{theatre_name}_summary.xlsx", output.getvalue())

    zip_buffer.seek(0)
    return StreamingResponse(
        zip_buffer, 
        media_type="application/x-zip-compressed",
        headers={"Content-Disposition": "attachment; filename=all_theatres_summary.zip"}
    )

@app.get("/health")
def health_check():
    return {"status": "healthy"}

@app.get("/ping")
async def ping():
    return {"status": "alive"}

@app.get("/users")
async def get_all_users(current_user: dict = Depends(get_current_user)):
    if not current_user["is_admin"]:
        raise HTTPException(status_code=403, detail="Not authorized")
    users = load_users()
    return [{"username": u, "is_admin": info["is_admin"]} for u, info in users.items()]

@app.post("/users")
async def add_new_user(user: UserCreate, current_user: dict = Depends(get_current_user)):
    if not current_user["is_admin"]:
        raise HTTPException(status_code=403, detail="Not authorized")
    users = load_users()
    if user.username in users:
        raise HTTPException(status_code=400, detail="User already exists")
    users[user.username] = {
        "password": pwd_context.hash(user.password),
        "is_admin": user.is_admin
    }
    save_users(users)
    return {"message": "User created"}

@app.get("/debug")
def debug():
    return {
        "SECRET_KEY": SECRET_KEY,
        "ALGORITHM": ALGORITHM,
        "ACCESS_TOKEN_EXPIRE_MINUTES": ACCESS_TOKEN_EXPIRE_MINUTES,
        "USERS_FILE": USERS_FILE,
        "VERCEL_LINK": os.getenv("VERCEL_LINK"),
        "origins": origins
    }