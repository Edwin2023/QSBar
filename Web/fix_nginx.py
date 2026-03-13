
import paramiko
import time

SERVER_IP = "182.92.141.58"
USERNAME = "root"
PASSWORD = "PKxyz1@3"

def fix_nginx():
    print(f"Fixing Nginx on {SERVER_IP}...")
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(SERVER_IP, username=USERNAME, password=PASSWORD)

    print("\n[Step 1] Killing Nginx and checking port usage...")
    ssh.exec_command("killall nginx")
    ssh.exec_command("systemctl stop nginx")
    time.sleep(2)
    
    # Check if anything is still listening on 80/443
    stdin, stdout, stderr = ssh.exec_command("netstat -tulnp | grep ':80\\|:443'")
    usage = stdout.read().decode().strip()
    
    if usage:
        print(f"WARNING: Ports still in use:\n{usage}")
        # If it's httpd (Apache), stop it
        if "httpd" in usage or "apache" in usage:
            print("Detected Apache/httpd. Stopping it...")
            ssh.exec_command("systemctl stop httpd")
            ssh.exec_command("systemctl disable httpd")
        else:
            # Force kill whatever is using port 80/443
            print("Force killing processes on port 80/443...")
            ssh.exec_command("fuser -k 80/tcp")
            ssh.exec_command("fuser -k 443/tcp")
    else:
        print("Ports 80/443 are free.")

    print("\n[Step 2] Testing Nginx Configuration...")
    stdin, stdout, stderr = ssh.exec_command("nginx -t")
    config_test = stderr.read().decode().strip()
    print(f"Config Test:\n{config_test}")
    
    if "successful" in config_test:
        print("\n[Step 3] Starting Nginx...")
        ssh.exec_command("systemctl start nginx")
        time.sleep(2)
        
        # Verify
        stdin, stdout, stderr = ssh.exec_command("systemctl status nginx")
        status = stdout.read().decode().strip()
        if "active (running)" in status:
            print("SUCCESS: Nginx restarted successfully.")
        else:
            print(f"FAIL: Nginx failed to start. Status:\n{status}")
            stdin, stdout, stderr = ssh.exec_command("journalctl -xeu nginx --no-pager | tail -n 20")
            print(f"Journal Log:\n{stdout.read().decode().strip()}")

    ssh.close()

if __name__ == "__main__":
    fix_nginx()
