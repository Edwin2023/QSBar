
import paramiko
import re

SERVER_IP = "182.92.141.58"
USERNAME = "root"
PASSWORD = "PKxyz1@3"
REMOTE_CERT = "/etc/nginx/cert/costspread.site.crt"

def clean_cert():
    print(f"Cleaning certificate on {SERVER_IP}...")
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(SERVER_IP, username=USERNAME, password=PASSWORD)

    # Read the messy cert file
    stdin, stdout, stderr = ssh.exec_command(f"cat {REMOTE_CERT}")
    content = stdout.read().decode()

    # Extract only the CERTIFICATE blocks
    matches = re.findall(r"-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----", content, re.DOTALL)
    
    if not matches:
        print("ERROR: No certificates found!")
        return

    print(f"Found {len(matches)} certificate blocks.")
    
    # Join them back together
    clean_content = "\n".join(matches) + "\n"
    
    # Write back to a temp file then move it
    temp_file = "/tmp/clean.crt"
    sftp = ssh.open_sftp()
    with sftp.file(temp_file, "w") as f:
        f.write(clean_content)
    sftp.close()
    
    # Overwrite the original file
    ssh.exec_command(f"mv {temp_file} {REMOTE_CERT}")
    
    # Reload Nginx
    print("Reloading Nginx...")
    ssh.exec_command("systemctl reload nginx")
    
    print("Certificate cleaned and Nginx reloaded.")
    ssh.close()

if __name__ == "__main__":
    clean_cert()
