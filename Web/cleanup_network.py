
import paramiko
import time

SERVER_IP = "182.92.141.58"
USERNAME = "root"
PASSWORD = "PKxyz1@3"

def cleanup_network():
    print(f"Cleaning up network on {SERVER_IP}...")
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(SERVER_IP, username=USERNAME, password=PASSWORD)

    print("Stopping Docker to release network rules...")
    ssh.exec_command("systemctl stop docker")
    ssh.exec_command("systemctl stop docker.socket")
    time.sleep(2)

    print("Flushing iptables rules (NAT & Filter)...")
    # Set default policies to ACCEPT to avoid lockout
    ssh.exec_command("iptables -P INPUT ACCEPT")
    ssh.exec_command("iptables -P FORWARD ACCEPT")
    ssh.exec_command("iptables -P OUTPUT ACCEPT")
    
    # Flush rules
    ssh.exec_command("iptables -t nat -F")
    ssh.exec_command("iptables -t mangle -F")
    ssh.exec_command("iptables -F")
    ssh.exec_command("iptables -X")
    
    print("Restarting Nginx...")
    ssh.exec_command("systemctl restart nginx")
    
    print("Network cleanup complete.")
    ssh.close()

if __name__ == "__main__":
    cleanup_network()
