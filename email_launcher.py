import argparse
import os
import pandas as pd
import random
import subprocess
import sys
import time
from typing import List, Tuple

def load_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        return pd.read_excel(path, engine="openpyxl")
    elif ext == ".csv":
        for enc in ["utf-8", "utf-8-sig", "cp1252", "latin-1"]:
            try:
                return pd.read_csv(path, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise UnicodeDecodeError("Could not decode CSV file")
    else:
        raise ValueError(f"Unsupported file type for {path}")

def safe_get(value, default=""):
    if pd.isna(value) or value is None:
        return default
    return str(value).strip()

def split_recipients_evenly(recipients: pd.DataFrame, num_chunks: int) -> List[pd.DataFrame]:
    """Split recipients into roughly equal chunks for each sender"""
    if num_chunks <= 0 or len(recipients) == 0:
        return [recipients]
    
    chunks = []
    total = len(recipients)
    base_size = total // num_chunks
    remainder = total % num_chunks
    
    start = 0
    for i in range(num_chunks):
        # Distribute remainder among first few chunks
        size = base_size + (1 if i < remainder else 0)
        end = start + size
        if start < total:
            chunk = recipients.iloc[start:end].copy()
            chunks.append(chunk)
            start = end
        else:
            chunks.append(pd.DataFrame())  # Empty chunk
    
    return chunks

def save_recipients_chunk(recipients: pd.DataFrame, chunk_index: int) -> str:
    """Save a chunk of recipients to a temporary file and return the filename"""
    filename = f"recipients_chunk_{chunk_index}.csv"
    recipients.to_csv(filename, index=False)
    return filename

    from indus_license_gate import require_license
    require_license()

def main():
    parser = argparse.ArgumentParser(description="Launch multiple terminal email senders")
    parser.add_argument("--recipients", default="recipients.xlsx", help="Path to recipients file")
    parser.add_argument("--senders", default="senders.csv", help="Path to senders file")
    parser.add_argument("--log", default="sent_log.csv", help="Central CSV log path")
    parser.add_argument("--min-delay", type=float, default=5.0)
    parser.add_argument("--max-delay", type=float, default=15.0)
    parser.add_argument("--max-per-sender", type=int, default=450)
    parser.add_argument("--retries", type=int, default=2)
    parser.add_argument("--limit", type=int, default=None, help="Optional global cap")
    parser.add_argument("--start-delay", type=int, default=30, help="Delay between starting terminals (seconds)")
    
    args = parser.parse_args()

    # Load senders
    try:
        senders = load_table(args.senders)
        senders["Email"] = senders["Email"].astype(str).str.strip()
        senders["AppPassword"] = senders["AppPassword"].astype(str).str.strip()
        
        if "Name" not in senders.columns:
            senders["Name"] = senders["Email"]
    except Exception as e:
        print(f"Error loading senders: {e}")
        return

    # Load recipients
    try:
        recipients = load_table(args.recipients)
        recipients = recipients.dropna(subset=["Email"])
        recipients["Email"] = recipients["Email"].astype(str).str.strip().str.lower()
        
        # Apply global limit if specified
        if args.limit and args.limit > 0:
            recipients = recipients.head(args.limit)
            
        print(f"📧 Total recipients to send: {len(recipients)}")
    except Exception as e:
        print(f"Error loading recipients: {e}")
        return

    # Split recipients into chunks for each sender
    recipient_chunks = split_recipients_evenly(recipients, len(senders))
    
    processes = []
    temp_files = []  # To keep track of temporary files for cleanup
    
    print("🚀 Launching email senders in separate terminals...")
    print(f"📊 Distributing {len(recipients)} recipients among {len(senders)} senders")
    
    for i, (_, sender) in enumerate(senders.iterrows()):
        email = safe_get(sender["Email"])
        password = safe_get(sender["AppPassword"])
        name = safe_get(sender["Name"])
        
        if not email or not password:
            print(f"Skipping sender {i+1}: missing email or password")
            continue
            
        # Get this sender's chunk of recipients
        if i < len(recipient_chunks):
            chunk = recipient_chunks[i]
            if len(chunk) == 0:
                print(f"Skipping sender {i+1} ({email}): no recipients to send")
                continue
                
            # Save chunk to temporary file
            chunk_filename = save_recipients_chunk(chunk, i)
            temp_files.append(chunk_filename)
            
            print(f"📨 Sender {i+1} ({email}) will send to {len(chunk)} recipients")
        else:
            print(f"Skipping sender {i+1} ({email}): no recipients chunk available")
            continue
            
        # Build command
        cmd = [
            sys.executable, "email_sender.py",
            "--email", email,
            "--password", password,
            "--name", name,
            "--recipients", chunk_filename,  # Use the chunk file instead of main recipients
            "--log", args.log,
            "--min-delay", str(args.min_delay),
            "--max-delay", str(args.max_delay),
            "--max-per-sender", str(min(args.max_per_sender, len(chunk))),  # Don't exceed chunk size
            "--retries", str(args.retries)
        ]
        
        # Launch in new terminal based on OS
        try:
            if os.name == 'nt':  # Windows
                process = subprocess.Popen(
                    ["start", "cmd", "/k"] + cmd,
                    shell=True
                )
            elif os.name == 'posix':  # Linux/Mac
                terminal_cmd = f"x-terminal-emulator -e '{' '.join(cmd)}'"
                process = subprocess.Popen(terminal_cmd, shell=True)
            else:
                print(f"Unsupported OS: {os.name}")
                return
                
            processes.append(process)
            print(f"✅ Launched sender {i+1}: {email}")
            
            # Delay between launches
            if i < len(senders) - 1:
                time.sleep(args.start_delay)
                
        except Exception as e:
            print(f"❌ Failed to launch sender {email}: {e}")
    
    print(f"\n🎯 Launched {len(processes)} sender terminals!")
    print("💡 Each terminal will send to a different set of recipients")
    print("⏳ Check each terminal for progress and any issues")
    
    # Cleanup function (will run when program exits)
    import atexit
    def cleanup_temp_files():
        for filename in temp_files:
            try:
                if os.path.exists(filename):
                    os.remove(filename)
                    print(f"🧹 Cleaned up temporary file: {filename}")
            except:
                pass
                
    atexit.register(cleanup_temp_files)

if __name__ == "__main__":
    main()