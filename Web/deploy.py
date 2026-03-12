
import paramiko
import os
import sys

# Configuration
SERVER_IP = "182.92.141.58"
USERNAME = "root"
PASSWORD = "PKxyz1@3"
LOCAL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
REMOTE_DIR = "/var/www/html"

def deploy():
    print(f"Connecting to {SERVER_IP}...")
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    
    try:
        ssh.connect(SERVER_IP, username=USERNAME, password=PASSWORD)
        print("Connected successfully.")
    except Exception as e:
        print(f"Failed to connect: {e}")
        # Try Windows username just in case
        try:
            print("Retrying with 'Administrator'...")
            ssh.connect(SERVER_IP, username="Administrator", password=PASSWORD)
            global REMOTE_DIR
            REMOTE_DIR = "C:/inetpub/wwwroot" # Default IIS path
            print("Connected as Administrator.")
        except Exception as e2:
            print(f"Failed to connect as Administrator: {e2}")
            return

    sftp = ssh.open_sftp()

    # Ensure remote directory exists (Linux)
    if REMOTE_DIR.startswith("/"):
        try:
            sftp.stat(REMOTE_DIR)
        except FileNotFoundError:
            print(f"Remote directory {REMOTE_DIR} does not exist. Creating...")
            ssh.exec_command(f"mkdir -p {REMOTE_DIR}")

    print(f"Uploading files from {LOCAL_DIR} to {REMOTE_DIR}...")
    
    files = os.listdir(LOCAL_DIR)
    for file in files:
        local_path = os.path.join(LOCAL_DIR, file)
        remote_path = f"{REMOTE_DIR}/{file}"
        if os.path.isfile(local_path):
            print(f"Uploading {file}...")
            sftp.put(local_path, remote_path)

    print("Upload complete.")
    
    # Check if web server is running (Linux only for now)
    if REMOTE_DIR.startswith("/"):
        stdin, stdout, stderr = ssh.exec_command("systemctl is-active nginx")
        status = stdout.read().decode().strip()
        if status != "active":
            print("Nginx is not active. Trying to install/start...")
            # Try to install nginx if missing
            ssh.exec_command("yum install -y nginx || apt-get install -y nginx")
            ssh.exec_command("systemctl start nginx")
            ssh.exec_command("systemctl enable nginx")
        else:
            print("Nginx is running.")

    sftp.close()
    ssh.close()
    print("Deployment finished successfully!")

if __name__ == "__main__":
    deploy()
