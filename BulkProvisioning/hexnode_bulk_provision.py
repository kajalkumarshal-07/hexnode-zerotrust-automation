"""
Hexnode Zero Trust - Bulk Device Provisioning
Reads device inventory from CSV and provisions devices with appropriate policies.
"""

import csv
import json
import requests
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class HexnodeBulkProvisioner:
    def __init__(self, config_path: str = None):
        if config_path is None:
            config_path = Path(__file__).parent.parent / "Config" / "hexnode_config.json"
        
        with open(config_path) as f:
            self.config = json.load(f)
        
        self.api_url = self.config["hexnode"]["api_url"]
        self.api_key = self.config["hexnode"]["api_key"]
        self.policy_mapping = self.config["provisioning"]["policy_mapping"]
        self.headers = {
            "Authorization": self.api_key,
            "Content-Type": "application/json"
        }
        
        self.results = {
            "timestamp": datetime.now().isoformat(),
            "total_devices": 0,
            "provisioned": 0,
            "failed": 0,
            "skipped": 0,
            "errors": []
        }
    
    def get_device_by_serial(self, serial_number: str) -> Optional[Dict]:
        """Find device in Hexnode by serial number."""
        try:
            response = requests.get(
                f"{self.api_url}/devices/?serial_number={serial_number}",
                headers=self.headers
            )
            response.raise_for_status()
            devices = response.json().get("devices", [])
            return devices[0] if devices else None
        except requests.RequestException as e:
            logger.error(f"Failed to find device {serial_number}: {e}")
            return None
    
    def add_device_to_group(self, device_id: int, group_id: int) -> bool:
        """Add device to a policy group."""
        try:
            payload = {
                "devices": [device_id],
                "devicegroups": [group_id]
            }
            response = requests.post(
                f"{self.api_url}/actions/add_to_device_group/",
                headers=self.headers,
                json=payload
            )
            response.raise_for_status()
            return True
        except requests.RequestException as e:
            logger.error(f"Failed to add device {device_id} to group: {e}")
            return False
    
    def apply_policy(self, device_id: int, policy_id: int) -> bool:
        """Apply a policy to a device."""
        try:
            payload = {
                "devices": [device_id],
                "policies": [policy_id]
            }
            response = requests.post(
                f"{self.api_url}/actions/apply_policy/",
                headers=self.headers,
                json=payload
            )
            response.raise_for_status()
            return True
        except requests.RequestException as e:
            logger.error(f"Failed to apply policy to device {device_id}: {e}")
            return False
    
    def install_app(self, device_id: int, app_id: int) -> bool:
        """Install an app on a device."""
        try:
            payload = {
                "devices": [device_id],
                "apps": [app_id]
            }
            response = requests.post(
                f"{self.api_url}/actions/install_applications/",
                headers=self.headers,
                json=payload
            )
            response.raise_for_status()
            return True
        except requests.RequestException as e:
            logger.error(f"Failed to install app on device {device_id}: {e}")
            return False
    
    def send_message(self, device_id: int, message: str) -> bool:
        """Send a message to a device."""
        try:
            payload = {
                "devices": [device_id],
                "message": message
            }
            response = requests.post(
                f"{self.api_url}/actions/send_message/",
                headers=self.headers,
                json=payload
            )
            response.raise_for_status()
            return True
        except requests.RequestException as e:
            logger.error(f"Failed to send message to device {device_id}: {e}")
            return False
    
    def provision_device(self, device_data: Dict) -> bool:
        """Provision a single device with all required configurations."""
        serial = device_data.get("serial_number")
        device_name = device_data.get("device_name")
        location = device_data.get("location")
        department = device_data.get("department")
        
        logger.info(f"Provisioning: {device_name} ({serial})")
        
        # Find device in Hexnode
        device = self.get_device_by_serial(serial)
        if not device:
            logger.warning(f"Device not found: {serial}")
            self.results["skipped"] += 1
            return False
        
        device_id = device["id"]
        success = True
        
        # 1. Add to location-based group
        group_id = self.policy_mapping.get(location.lower())
        if group_id:
            if self.add_device_to_group(device_id, group_id):
                logger.info(f"  Added to group: {location}")
            else:
                success = False
                self.results["errors"].append(f"Failed to add {device_name} to group")
        
        # 2. Apply department-specific policy
        policy_id = self.policy_mapping.get(department.lower())
        if policy_id:
            if self.apply_policy(device_id, policy_id):
                logger.info(f"  Applied policy: {department}")
            else:
                success = False
                self.results["errors"].append(f"Failed to apply policy to {device_name}")
        
        # 3. Install required apps (example: security agent)
        security_app_id = 100  # Replace with your security app ID
        if self.install_app(device_id, security_app_id):
            logger.info(f"  Installed security agent")
        
        # 4. Send welcome message
        welcome_msg = f"Welcome to Zero Trust network. Your device has been provisioned for {location}."
        self.send_message(device_id, welcome_msg)
        
        if success:
            self.results["provisioned"] += 1
            logger.info(f"  Provisioning complete: {device_name}")
        else:
            self.results["failed"] += 1
        
        return success
    
    def provision_from_csv(self, csv_path: str) -> Dict:
        """Provision all devices from a CSV file."""
        logger.info(f"Starting bulk provisioning from: {csv_path}")
        
        with open(csv_path, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            devices = list(reader)
        
        self.results["total_devices"] = len(devices)
        logger.info(f"Found {len(devices)} devices to provision")
        
        for device_data in devices:
            self.provision_device(device_data)
        
        logger.info(f"Provisioning complete: {self.results['provisioned']} success, "
                    f"{self.results['failed']} failed, {self.results['skipped']} skipped")
        
        return self.results
    
    def provision_smart_group(self, criteria: Dict) -> Dict:
        """Provision devices matching specific criteria."""
        logger.info(f"Provisioning devices matching criteria: {criteria}")
        
        # Get all devices
        try:
            response = requests.get(
                f"{self.api_url}/devices/",
                headers=self.headers
            )
            response.raise_for_status()
            devices = response.json().get("devices", [])
        except requests.RequestException as e:
            logger.error(f"Failed to fetch devices: {e}")
            return self.results
        
        # Filter by criteria
        matching_devices = []
        for device in devices:
            match = True
            for key, value in criteria.items():
                if device.get(key) != value:
                    match = False
                    break
            if match:
                matching_devices.append(device)
        
        logger.info(f"Found {len(matching_devices)} matching devices")
        
        # Provision matching devices
        for device in matching_devices:
            device_data = {
                "serial_number": device.get("serial_number"),
                "device_name": device.get("device_name"),
                "location": device.get("location", "unknown"),
                "department": device.get("department", "unknown")
            }
            self.provision_device(device_data)
        
        return self.results


def main():
    """Example usage."""
    provisioner = HexnodeBulkProvisioner()
    
    # Provision from CSV
    csv_path = Path(__file__).parent / "devices.csv"
    if csv_path.exists():
        results = provisioner.provision_from_csv(str(csv_path))
        print(json.dumps(results, indent=2))
    else:
        logger.warning(f"CSV file not found: {csv_path}")
        logger.info("Creating sample CSV...")
        
        # Create sample CSV
        sample_data = [
            {"device_name": "NYC-PC-001", "serial_number": "SN12345678", "location": "nyc", "department": "corporate"},
            {"device_name": "LA-PC-001", "serial_number": "SN23456789", "location": "la", "department": "retail"},
            {"device_name": "CHI-PC-001", "serial_number": "SN34567890", "location": "chi", "department": "corporate"}
        ]
        
        with open(csv_path, 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=sample_data[0].keys())
            writer.writeheader()
            writer.writerows(sample_data)
        
        logger.info(f"Sample CSV created at: {csv_path}")
        logger.info("Please update with your device inventory")


if __name__ == "__main__":
    main()
