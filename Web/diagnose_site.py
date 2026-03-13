
import paramiko
import os

# Configuration
SERVER_IP = "182.92.141.58"
USERNAME = "root"
PASSWORD = "PKxyz1@3"

def diagnose():
    print(f"Diagnosing {SERVER_IP}...")
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        ssh.connect(SERVER_IP, username=USERNAME, password=PASSWORD)
    except Exception as e:
        print(f"Failed to connect: {e}")
        return

    print("\n[Check 1] Is Nginx running and listening on port 443?")
    # Check if nginx process is running
    stdin, stdout, stderr = ssh.exec_command("ps aux | grep nginx | grep -v grep")
    nginx_procs = stdout.read().decode().strip()
    if nginx_procs:
        print("PASS: Nginx process is running.")
    else:
        print("FAIL: Nginx is NOT running!")
    
    # Check if port 443 is open
    stdin, stdout, stderr = ssh.exec_command("netstat -tuln | grep :443")
    port_status = stdout.read().decode().strip()
    if port_status:
        print(f"PASS: Port 443 is listening: {port_status}")
    else:
        print("FAIL: Port 443 is NOT listening!")

    print("\n[Check 2] Local connectivity test (curl localhost)")
    stdin, stdout, stderr = ssh.exec_command("curl -I -k https://localhost")
    curl_out = stdout.read().decode().strip()
    if "HTTP/1.1 200 OK" in curl_out or "HTTP/1.1 301 Moved Permanently" in curl_out:
        print("PASS: Local curl request succeeded.")
    else:
        print(f"FAIL: Local curl failed. Output:\n{curl_out}")
        print(f"Error:\n{stderr.read().decode().strip()}")

    print("\n[Check 3] Firewall status (firewalld/iptables)")
    stdin, stdout, stderr = ssh.exec_command("systemctl status firewalld")
    firewall_status = stdout.read().decode().strip()
    if "active (running)" in firewall_status:
        print("WARN: firewalld is running. Checking if https service is allowed...")
        stdin, stdout, stderr = ssh.exec_command("firewall-cmd --list-all")
        rules = stdout.read().decode().strip()
        if "https" in rules:
            print("PASS: 'https' service is allowed in firewalld.")
        else:
            print("FAIL: 'https' service is NOT allowed in firewalld! Trying to add it...")
            ssh.exec_command("firewall-cmd --permanent --add-service=https")
            ssh.exec_command("firewall-cmd --permanent --add-service=http")
            ssh.exec_command("firewall-cmd --reload")
            print("Action: Added http/https to firewalld rules.")
    else:
        print("INFO: firewalld is not active (which is fine if relying on security groups).")

    print("\n[Check 4] Nginx Configuration Test")
    stdin, stdout, stderr = ssh.exec_command("nginx -t")
    config_test = stderr.read().decode().strip() # nginx -t output goes to stderr
    print(config_test)

    ssh.close()

if __name__ == "__main__":
    diagnose()
