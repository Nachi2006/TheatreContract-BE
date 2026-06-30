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
def create_excel_with_totals(df: pd.DataFrame, output_io: io.BytesIO, with_third_party: bool = False):
    df = df.copy()
    df = df.fillna("")
    
    # Ensure minimum required columns exist
    if len(df.columns) < 19:
        writer = pd.ExcelWriter(output_io, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Summary')
        writer.close()
        return

    # 1. Define Columns to Keep/Drop
    # Indices to Drop: B(1), E(4), F(5), G(6), I(8), J(9)
    indices_to_drop = [1, 4, 5, 6, 8, 9]
    # Keep A to S (0 to 18) minus dropped
    indices_to_keep = [i for i in range(19) if i not in indices_to_drop]
    
    # Append X(23) and Y(24) if Third Party is checked
    if with_third_party and len(df.columns) > 24:
        indices_to_keep.extend([23, 24])
        
    cols_to_keep = [df.columns[i] for i in indices_to_keep]
    
    # 2. Define exactly which columns to sum: P(15), R(17), S(18) | And X(23), Y(24)
    sum_indices = [15, 17, 18]
    if with_third_party and len(df.columns) > 24:
        sum_indices.extend([23, 24])
        
    sum_cols = [df.columns[i] for i in sum_indices if i < len(df.columns)]
    
    # Force numeric conversion on the exact sum columns
    for col in sum_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # Sort DataFrame by the injected 'Combined Name'
    df = df.sort_values(by='Combined Name')
    
    # Create the output dataframe with dropped columns
    df_out = df[cols_to_keep]
    
    # Scrub "Unnamed:" headers caused by pd.read_excel on merged rows
    cleaned_headers = [c if not str(c).startswith("Unnamed:") else "" for c in df_out.columns]
    df_out.columns = cleaned_headers
    
    # Re-map our sum columns to their cleaned header names so we can find them
    cleaned_sum_cols = [c if not str(c).startswith("Unnamed:") else "" for c in sum_cols]
    
    # Find the NEW indices of the sum columns in the pruned dataframe
    new_sum_locs = {col: df_out.columns.get_loc(col) for col in cleaned_sum_cols}
    
    # Setup Writer
    writer = pd.ExcelWriter(output_io, engine='xlsxwriter')
    workbook = writer.book
    
    subtotal_fmt = workbook.add_format({'bold': True, 'bg_color': '#E2EFDA'})
    summary_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2'})
    
    # Print Headers
    df_out[:0].to_excel(writer, index=False, sheet_name='Summary')
    worksheet = writer.sheets['Summary']
    
    current_row = 1
    
    # 3. Iterate and print Chunks by 'Combined Name'
    for screen_name, group in df.groupby('Combined Name', sort=False):
        group_out = group[cols_to_keep]
        group_out.to_excel(writer, index=False, header=False, startrow=current_row, sheet_name='Summary')
        current_row += len(group_out)
        
        # Subtotals row
        screen_val = str(screen_name) if str(screen_name).strip() != "" else "Unknown Screen"
        worksheet.write_string(current_row, 0, f"{screen_val} Subtotal", subtotal_fmt)
        
        # Inject the math for each specific sum column
        for original_col, cleaned_col in zip(sum_cols, cleaned_sum_cols):
            loc = new_sum_locs[cleaned_col]
            worksheet.write_number(current_row, loc, float(group[original_col].sum()), subtotal_fmt)
            
        current_row += 1 

    # 4. Overall Summary at the Bottom
    current_row += 2 
    screen_totals = df.groupby('Combined Name')[sum_cols].sum().reset_index()
    
    worksheet.write_string(current_row, 0, "OVERALL SUMMARY", summary_fmt)
    current_row += 1
    
    for index, row in screen_totals.iterrows():
        screen_val = str(row['Combined Name'])
        worksheet.write_string(current_row, 0, screen_val)
        
        for original_col, cleaned_col in zip(sum_cols, cleaned_sum_cols):
            loc = new_sum_locs[cleaned_col]
            worksheet.write_number(current_row, loc, float(row[original_col]))
            
        current_row += 1
        
    if len(screen_totals) > 1:
        worksheet.write_string(current_row, 0, "Grand Total", summary_fmt)
        for original_col, cleaned_col in zip(sum_cols, cleaned_sum_cols):
            loc = new_sum_locs[cleaned_col]
            worksheet.write_number(current_row, loc, float(screen_totals[original_col].sum()), summary_fmt)
        
    writer.close()


@app.post("/extract-theatres")
async def extract_theatres(file: UploadFile = File(...), current_user: dict = Depends(get_current_user)):
    file_ext = os.path.splitext(file.filename)[1].lower()
    engine_choice = 'pyxlsb' if file_ext == '.xlsb' else None
    
    contents = await file.read()
    df = pd.read_excel(io.BytesIO(contents), header=3, engine=engine_choice)
    
    # Create D + C combination (Screen Name + Screen Code)
    if len(df.columns) > 3:
        df['Combined Name'] = df.iloc[:, 3].astype(str) + " - " + df.iloc[:, 2].astype(str)
        theatres = df['Combined Name'].dropna().unique().tolist()
        theatres.sort()
    else:
        theatres = []
        
    return {"theatres": theatres}


@app.post("/generate-custom-excel")
async def generate_custom_excel(
    file: UploadFile = File(...),
    selected_theatres: str = Form(...),
    with_third_party: str = Form("false"),
    current_user: dict = Depends(get_current_user)
):
    file_ext = os.path.splitext(file.filename)[1].lower()
    engine_choice = 'pyxlsb' if file_ext == '.xlsb' else None
    
    is_third_party = with_third_party.lower() == "true"
    theatres_list = json.loads(selected_theatres)
    
    contents = await file.read()
    df = pd.read_excel(io.BytesIO(contents), header=3, engine=engine_choice)
    
    if len(df.columns) > 3:
        df['Combined Name'] = df.iloc[:, 3].astype(str) + " - " + df.iloc[:, 2].astype(str)
        filtered_df = df[df['Combined Name'].isin(theatres_list)]
    else:
        filtered_df = df
    
    output = io.BytesIO()
    create_excel_with_totals(filtered_df, output, is_third_party)
    output.seek(0)
    
    return StreamingResponse(
        output, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=selected_theatres_summary.xlsx"}
    )


@app.post("/generate-all-zip")
async def generate_all_zip(
    file: UploadFile = File(...),
    with_third_party: str = Form("false"),
    current_user: dict = Depends(get_current_user)
):
    file_ext = os.path.splitext(file.filename)[1].lower()
    engine_choice = 'pyxlsb' if file_ext == '.xlsb' else None
    
    is_third_party = with_third_party.lower() == "true"
    
    contents = await file.read()
    df = pd.read_excel(io.BytesIO(contents), header=3, engine=engine_choice)
    
    if len(df.columns) > 3:
        df['Combined Name'] = df.iloc[:, 3].astype(str) + " - " + df.iloc[:, 2].astype(str)
        unique_theatres = df['Combined Name'].dropna().unique()
    else:
        unique_theatres = []
    
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for theatre_name in unique_theatres:
            theatre_df = df[df['Combined Name'] == theatre_name]
            
            if theatre_df.empty:
                continue
                
            output = io.BytesIO()
            create_excel_with_totals(theatre_df, output, is_third_party)
            
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