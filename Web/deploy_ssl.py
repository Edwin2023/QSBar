
import paramiko
import os
import time

# Configuration
SERVER_IP = "182.92.141.58"
USERNAME = "root"
PASSWORD = "PKxyz1@3"
DOMAIN = "costspread.site"

# Local paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CERT_DIR = os.path.join(SCRIPT_DIR, "cert", "costspread.site_tomcat")
PFX_FILE = os.path.join(CERT_DIR, "costspread.site.pfx")
PASS_FILE = os.path.join(CERT_DIR, "keystorePass.txt")

# Remote paths
DEFAULT_NGINX_CONF_DIR = "/etc/nginx/conf.d"
REMOTE_CERT_DIR = "/etc/nginx/cert"

def get_keystore_pass():
    with open(PASS_FILE, 'r') as f:
        return f.read().strip()

def deploy_ssl():
    # Use a local variable to track the actual config directory
    remote_conf_dir = DEFAULT_NGINX_CONF_DIR

    print(f"Connecting to {SERVER_IP}...")
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(SERVER_IP, username=USERNAME, password=PASSWORD)
    sftp = ssh.open_sftp()

    print("Checking Nginx installation...")
    stdin, stdout, stderr = ssh.exec_command("which nginx")
    if not stdout.read():
        print("Installing nginx...")
        ssh.exec_command("yum install -y nginx || apt-get install -y nginx")
        time.sleep(10) # Wait for install

    print("Checking dependencies...")
    stdin, stdout, stderr = ssh.exec_command("which openssl")
    if not stdout.read():
        print("Installing openssl...")
        ssh.exec_command("yum install -y openssl || apt-get install -y openssl")

    # Create directories
    print(f"Creating directories...")
    ssh.exec_command(f"mkdir -p {REMOTE_CERT_DIR}")
    ssh.exec_command(f"mkdir -p {remote_conf_dir}")

    # Upload PFX
    remote_pfx = f"{REMOTE_CERT_DIR}/{DOMAIN}.pfx"
    print(f"Uploading PFX file to {remote_pfx}...")
    sftp.put(PFX_FILE, remote_pfx)

    # Convert PFX to PEM (CRT + KEY)
    password = get_keystore_pass()
    print("Extracting certificates from PFX...")
    
    # Extract Private Key
    cmd_key = f"openssl pkcs12 -in {remote_pfx} -nocerts -out {REMOTE_CERT_DIR}/{DOMAIN}.key -nodes -passin pass:{password}"
    stdin, stdout, stderr = ssh.exec_command(cmd_key)
    err = stderr.read().decode()
    if err: print(f"Key extraction warning/error: {err}")

    # Extract Certificate (Server + CA Chain)
    # Remove -clcerts to include CA chain, verify output contains multiple certificates
    cmd_crt = f"openssl pkcs12 -in {remote_pfx} -nokeys -out {REMOTE_CERT_DIR}/{DOMAIN}.crt -passin pass:{password}"
    stdin, stdout, stderr = ssh.exec_command(cmd_crt)
    err = stderr.read().decode()
    if err: print(f"Cert extraction warning/error: {err}")

    # Generate Nginx Config
    nginx_conf = f"""
server {{
    listen 80;
    server_name {DOMAIN} www.{DOMAIN};
    return 301 https://$host$request_uri;
}}

server {{
    listen 443 ssl;
    server_name {DOMAIN} www.{DOMAIN};

    ssl_certificate {REMOTE_CERT_DIR}/{DOMAIN}.crt;
    ssl_certificate_key {REMOTE_CERT_DIR}/{DOMAIN}.key;

    ssl_session_timeout 5m;
    ssl_ciphers ECDHE-RSA-AES128-GCM-SHA256:ECDHE:ECDH:AES:HIGH:!NULL:!aNULL:!MD5:!ADH:!RC4;
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_prefer_server_ciphers on;

    root /var/www/html;
    index index.html;

    location / {{
        try_files $uri $uri/ =404;
    }}
}}
"""
    local_conf_path = os.path.join(SCRIPT_DIR, "qsbar.conf")
    with open(local_conf_path, "w", encoding="utf-8") as f:
        f.write(nginx_conf)

    # Upload Nginx Config
    remote_conf_path = f"{remote_conf_dir}/qsbar.conf"
    print(f"Uploading Nginx configuration to {remote_conf_path}...")
    try:
        sftp.put(local_conf_path, remote_conf_path)
    except FileNotFoundError:
        # If /etc/nginx/conf.d doesn't exist, try /etc/nginx/sites-enabled (Ubuntu/Debian style)
        remote_conf_dir = "/etc/nginx/sites-enabled"
        ssh.exec_command(f"mkdir -p {remote_conf_dir}")
        remote_conf_path = f"{remote_conf_dir}/qsbar.conf"
        print(f"Retrying upload to {remote_conf_path}...")
        sftp.put(local_conf_path, remote_conf_path)

    # Restart Nginx
    print("Restarting Nginx...")
    ssh.exec_command("nginx -t")
    ssh.exec_command("systemctl restart nginx || service nginx restart")

    print("SSL Deployment Complete!")
    print(f"You can now visit https://{DOMAIN}")

    sftp.close()
    ssh.close()
    
    if os.path.exists(local_conf_path):
        os.remove(local_conf_path)

if __name__ == "__main__":
    deploy_ssl()
