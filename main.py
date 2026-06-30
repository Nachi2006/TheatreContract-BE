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

# --- HELPER FUNCTION FOR TOTALS ---
def create_excel_with_totals(df: pd.DataFrame, output_io: io.BytesIO):
    # 1. Create a deep copy to prevent Memory Slice corruption
    df = df.copy()
    
    # 2. Scrub Pandas NaNs (which corrupt MS Excel XML files)
    df = df.fillna("")
    
    COL_A_IDX = 0
    COL_C_IDX = 2
    COL_O_IDX = 14
    COL_Q_IDX = 16
    
    if len(df.columns) <= COL_Q_IDX:
        writer = pd.ExcelWriter(output_io, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Summary')
        writer.close()
        return

    col_c_name = df.columns[COL_C_IDX]
    col_o_name = df.columns[COL_O_IDX]
    col_q_name = df.columns[COL_Q_IDX]
    
    # Force numeric conversion, setting blanks to 0
    df[col_o_name] = pd.to_numeric(df[col_o_name], errors='coerce').fillna(0)
    df[col_q_name] = pd.to_numeric(df[col_q_name], errors='coerce').fillna(0)
    
    # 3. Explicit writer management
    writer = pd.ExcelWriter(output_io, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Summary')
    
    workbook = writer.book
    worksheet = writer.sheets['Summary']
    
    start_row = len(df) + 2
    screen_totals = df.groupby(col_c_name)[[col_o_name, col_q_name]].sum().reset_index()
    
    current_row = start_row
    for index, row in screen_totals.iterrows():
        # Ensure pure string and float types are passed to xlsxwriter
        screen_val = str(row[col_c_name]) if str(row[col_c_name]).strip() != "" else "Unknown Screen"
        
        worksheet.write_string(current_row, COL_A_IDX, screen_val)
        worksheet.write_number(current_row, COL_O_IDX, float(row[col_o_name]))
        worksheet.write_number(current_row, COL_Q_IDX, float(row[col_q_name]))
        current_row += 1
        
    if len(screen_totals) > 1:
        worksheet.write_string(current_row, COL_A_IDX, "Grand Total")
        worksheet.write_number(current_row, COL_O_IDX, float(screen_totals[col_o_name].sum()))
        worksheet.write_number(current_row, COL_Q_IDX, float(screen_totals[col_q_name].sum()))
        
    # 4. Explicitly close the file to seal the binary stream
    writer.close()


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
    current_user: dict = Depends(get_current_user)
):
    theatres_list = json.loads(selected_theatres)
    contents = await file.read()
    df = pd.read_excel(io.BytesIO(contents), header=3)
    
    filtered_df = df[df['Theatre Name'].isin(theatres_list)]
    
    output = io.BytesIO()
    create_excel_with_totals(filtered_df, output)
    output.seek(0)
    
    return StreamingResponse(
        output, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=selected_theatres_summary.xlsx"}
    )

@app.post("/generate-all-zip")
async def generate_all_zip(
    file: UploadFile = File(...),
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
            create_excel_with_totals(theatre_df, output)
            
            # Sanitize the filename to prevent ZIP corruption from illegal characters
            safe_name = str(theatre_name).replace("/", "-").replace("\\", "-")
            zipf.writestr(f"{safe_name}_summary.xlsx", output.getvalue())

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