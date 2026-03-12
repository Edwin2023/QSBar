
import paramiko

SERVER_IP = "182.92.141.58"
USERNAME = "root"
PASSWORD = "PKxyz1@3"

def diagnose_v2():
    print(f"Diagnosing V2 for {SERVER_IP}...")
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(SERVER_IP, username=USERNAME, password=PASSWORD)

    print("\n[Check 1] Web Content Check")
    stdin, stdout, stderr = ssh.exec_command("ls -la /var/www/html/index.html")
    file_info = stdout.read().decode().strip()
    if file_info:
        print(f"PASS: index.html found: {file_info}")
    else:
        print("FAIL: index.html NOT found in /var/www/html!")

    print("\n[Check 2] SELinux Status")
    stdin, stdout, stderr = ssh.exec_command("getenforce")
    selinux = stdout.read().decode().strip()
    print(f"SELinux Status: {selinux}")
    
    if selinux == "Enforcing":
        print("WARNING: SELinux is Enforcing. This might block Nginx.")
        # Check audit log for denials (briefly)
        stdin, stdout, stderr = ssh.exec_command("grep nginx /var/log/audit/audit.log | tail -n 5")
        audit_log = stdout.read().decode().strip()
        if audit_log:
            print(f"Audit Log Hints:\n{audit_log}")
        else:
            print("No recent audit denials found for nginx.")
        
        # Temporary fix suggestion (or apply it)
        print("Trying to set SELinux to Permissive for testing...")
        ssh.exec_command("setenforce 0")
        print("Action: Set SELinux to Permissive temporarily.")

    print("\n[Check 3] Nginx Error Log (Last 10 lines)")
    stdin, stdout, stderr = ssh.exec_command("tail -n 10 /var/log/nginx/error.log")
    error_log = stdout.read().decode().strip()
    if error_log:
        print(f"Error Log:\n{error_log}")
    else:
        print("Error log is empty or unreadable.")

    ssh.close()

if __name__ == "__main__":
    diagnose_v2()
