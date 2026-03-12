
import paramiko

SERVER_IP = "182.92.141.58"
USERNAME = "root"
PASSWORD = "PKxyz1@3"
REMOTE_CERT = "/etc/nginx/cert/costspread.site.crt"

def check_cert():
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(SERVER_IP, username=USERNAME, password=PASSWORD)
    
    print("Checking certificate file content...")
    stdin, stdout, stderr = ssh.exec_command(f"cat {REMOTE_CERT}")
    content = stdout.read().decode()
    
    count = content.count("BEGIN CERTIFICATE")
    print(f"Found {count} certificates in the CRT file.")
    
    if count < 2:
        print("WARNING: Only 1 certificate found. Intermediate chain might be missing!")
    else:
        print("OK: Certificate chain appears to be present.")

    ssh.close()

if __name__ == "__main__":
    check_cert()
