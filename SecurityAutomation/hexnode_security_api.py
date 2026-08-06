"""
Hexnode Zero Trust - Security Automation API Client
Handles device quarantine, compliance checks, and automated remediation via Hexnode REST API.
"""

import json
import requests
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class HexnodeZeroTrust:
    def __init__(self, config_path: str = None):
        if config_path is None:
            config_path = Path(__file__).parent.parent / "Config" / "hexnode_config.json"
        
        with open(config_path) as f:
            self.config = json.load(f)
        
        self.api_url = self.config["hexnode"]["api_url"]
        self.api_key = self.config["hexnode"]["api_key"]
        self.quarantine_group_id = self.config["quarantine"]["quarantine_group_id"]
        self.headers = {
            "Authorization": self.api_key,
            "Content-Type": "application/json"
        }
    
    def get_all_devices(self) -> List[Dict]:
        """Fetch all managed devices from Hexnode."""
        try:
            response = requests.get(
                f"{self.api_url}/devices/",
                headers=self.headers
            )
            response.raise_for_status()
            return response.json().get("devices", [])
        except requests.RequestException as e:
            logger.error(f"Failed to fetch devices: {e}")
            return []
    
    def get_device_compliance(self, device_id: int) -> Dict:
        """Get compliance status for a specific device."""
        try:
            response = requests.get(
                f"{self.api_url}/devices/{device_id}/",
                headers=self.headers
            )
            response.raise_for_status()
            device = response.json()
            return {
                "device_id": device_id,
                "device_name": device.get("device_name"),
                "platform": device.get("platform"),
                "compliance_status": device.get("compliance_status"),
                "os_version": device.get("os_version"),
                "last_seen": device.get("last_seen")
            }
        except requests.RequestException as e:
            logger.error(f"Failed to get device {device_id}: {e}")
            return {}
    
    def quarantine_device(self, device_id: int, reason: str = "Non-compliant") -> bool:
        """Move device to quarantine group (restricted policy)."""
        try:
            payload = {
                "devices": [device_id],
                "devicegroups": [self.quarantine_group_id]
            }
            response = requests.post(
                f"{self.api_url}/actions/add_to_device_group/",
                headers=self.headers,
                json=payload
            )
            response.raise_for_status()
            logger.info(f"Device {device_id} quarantined: {reason}")
            return True
        except requests.RequestException as e:
            logger.error(f"Failed to quarantine device {device_id}: {e}")
            return False
    
    def remove_from_quarantine(self, device_id: int, target_group_id: int) -> bool:
        """Remove device from quarantine and assign to target group."""
        try:
            # Remove from quarantine group
            payload = {
                "devices": [device_id],
                "devicegroups": [self.quarantine_group_id]
            }
            requests.post(
                f"{self.api_url}/actions/remove_from_device_group/",
                headers=self.headers,
                json=payload
            )
            
            # Add to target group
            payload = {
                "devices": [device_id],
                "devicegroups": [target_group_id]
            }
            response = requests.post(
                f"{self.api_url}/actions/add_to_device_group/",
                headers=self.headers,
                json=payload
            )
            response.raise_for_status()
            logger.info(f"Device {device_id} removed from quarantine")
            return True
        except requests.RequestException as e:
            logger.error(f"Failed to unquarantine device {device_id}: {e}")
            return False
    
    def enable_lost_mode(self, device_id: int, message: str = "Security Alert - Device Quarantined") -> bool:
        """Enable lost mode on a compromised device."""
        try:
            payload = {
                "devices": [device_id],
                "message": message
            }
            response = requests.post(
                f"{self.api_url}/actions/enable_lost_mode/",
                headers=self.headers,
                json=payload
            )
            response.raise_for_status()
            logger.info(f"Lost mode enabled on device {device_id}")
            return True
        except requests.RequestException as e:
            logger.error(f"Failed to enable lost mode on {device_id}: {e}")
            return False
    
    def wipe_device(self, device_id: int, confirm: bool = False) -> bool:
        """Wipe a compromised device (use with caution)."""
        if not confirm:
            logger.warning("Wipe not confirmed - skipping")
            return False
        
        try:
            payload = {"devices": [device_id]}
            response = requests.post(
                f"{self.api_url}/actions/wipe_device/",
                headers=self.headers,
                json=payload
            )
            response.raise_for_status()
            logger.warning(f"Device {device_id} wipe initiated")
            return True
        except requests.RequestException as e:
            logger.error(f"Failed to wipe device {device_id}: {e}")
            return False
    
    def execute_custom_script(self, device_id: int, script_id: int) -> bool:
        """Execute a custom script on a device."""
        try:
            payload = {
                "devices": [device_id],
                "scripts": [script_id]
            }
            response = requests.post(
                f"{self.api_url}/actions/execute_script/",
                headers=self.headers,
                json=payload
            )
            response.raise_for_status()
            logger.info(f"Script {script_id} executed on device {device_id}")
            return True
        except requests.RequestException as e:
            logger.error(f"Failed to execute script on {device_id}: {e}")
            return False
    
    def scan_and_quarantine(self) -> Dict:
        """Scan all devices and quarantine non-compliant ones."""
        results = {
            "timestamp": datetime.now().isoformat(),
            "total_devices": 0,
            "compliant": 0,
            "quarantined": 0,
            "errors": 0,
            "details": []
        }
        
        devices = self.get_all_devices()
        results["total_devices"] = len(devices)
        
        for device in devices:
            device_id = device.get("id")
            compliance = self.get_device_compliance(device_id)
            
            if not compliance:
                results["errors"] += 1
                continue
            
            device_detail = {
                "device_id": device_id,
                "device_name": compliance.get("device_name"),
                "compliance_status": compliance.get("compliance_status"),
                "action_taken": "none"
            }
            
            if compliance.get("compliance_status") == "non-compliant":
                success = self.quarantine_device(device_id, "Auto-quarantined by Zero Trust policy")
                if success:
                    results["quarantined"] += 1
                    device_detail["action_taken"] = "quarantined"
                else:
                    results["errors"] += 1
                    device_detail["action_taken"] = "quarantine_failed"
            else:
                results["compliant"] += 1
                device_detail["action_taken"] = "compliant"
            
            results["details"].append(device_detail)
        
        logger.info(f"Scan complete: {results['compliant']} compliant, {results['quarantined']} quarantined")
        return results
    
    def process_siem_alert(self, alert: Dict) -> Dict:
        """Process an alert from SIEM (Splunk, etc.) and take action."""
        device_id = alert.get("device_id")
        severity = alert.get("severity", "low")
        alert_type = alert.get("type", "unknown")
        
        action_result = {
            "device_id": device_id,
            "alert_type": alert_type,
            "severity": severity,
            "action_taken": "none"
        }
        
        if severity == "critical":
            # Critical: Enable lost mode and quarantine
            self.enable_lost_mode(device_id, f"CRITICAL ALERT: {alert_type}")
            self.quarantine_device(device_id, f"SIEM Alert: {alert_type}")
            action_result["action_taken"] = "lost_mode_and_quarantine"
        
        elif severity == "high":
            # High: Quarantine device
            self.quarantine_device(device_id, f"SIEM Alert: {alert_type}")
            action_result["action_taken"] = "quarantined"
        
        elif severity == "medium":
            # Medium: Run remediation script
            self.execute_custom_script(device_id, script_id=1)  # Remediation script ID
            action_result["action_taken"] = "remediation_script_executed"
        
        else:
            # Low: Log only
            action_result["action_taken"] = "logged_only"
        
        return action_result


def main():
    """Example usage."""
    zt = HexnodeZeroTrust()
    
    # Run compliance scan
    results = zt.scan_and_quarantine()
    print(json.dumps(results, indent=2))
    
    # Example: Process SIEM alert
    # siem_alert = {
    #     "device_id": 12345,
    #     "type": "Malware Detected",
    #     "severity": "critical"
    # }
    # zt.process_siem_alert(siem_alert)


if __name__ == "__main__":
    main()
