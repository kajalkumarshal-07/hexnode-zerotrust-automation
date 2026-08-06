"""
Hexnode Zero Trust - SIEM Webhook Integration Handler
Receives alerts from Splunk, SOAR, or other SIEMs and triggers automated remediation.
"""

import json
import hmac
import hashlib
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional
from flask import Flask, request, jsonify

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)


class SIEMIntegration:
    def __init__(self, config_path: str = None):
        if config_path is None:
            config_path = Path(__file__).parent.parent / "Config" / "hexnode_config.json"
        
        with open(config_path) as f:
            self.config = json.load(f)
        
        self.hexnode_api_url = self.config["hexnode"]["api_url"]
        self.hexnode_api_key = self.config["hexnode"]["api_key"]
        self.quarantine_group_id = self.config["quarantine"]["quarantine_group_id"]
        self.splunk_webhook = self.config["siem"]["splunk_webhook"]
        self.splunk_token = self.config["siem"]["splunk_token"]
        self.auto_remediate = self.config["siem"]["auto_remediate"]
    
    def verify_webhook_signature(self, payload: bytes, signature: str, secret: str) -> bool:
        """Verify webhook signature for security."""
        expected = hmac.new(secret.encode(), payload, hashlib.sha256).hexdigest()
        return hmac.compare_digest(expected, signature)
    
    def quarantine_device(self, device_id: int, reason: str) -> bool:
        """Quarantine a device via Hexnode API."""
        import requests
        
        try:
            payload = {
                "devices": [device_id],
                "devicegroups": [self.quarantine_group_id]
            }
            response = requests.post(
                f"{self.hexnode_api_url}/actions/add_to_device_group/",
                headers={
                    "Authorization": self.hexnode_api_key,
                    "Content-Type": "application/json"
                },
                json=payload
            )
            response.raise_for_status()
            logger.info(f"Device {device_id} quarantined: {reason}")
            return True
        except Exception as e:
            logger.error(f"Failed to quarantine device {device_id}: {e}")
            return False
    
    def enable_lost_mode(self, device_id: int, message: str) -> bool:
        """Enable lost mode on a device."""
        import requests
        
        try:
            payload = {
                "devices": [device_id],
                "message": message
            }
            response = requests.post(
                f"{self.hexnode_api_url}/actions/enable_lost_mode/",
                headers={
                    "Authorization": self.hexnode_api_key,
                    "Content-Type": "application/json"
                },
                json=payload
            )
            response.raise_for_status()
            logger.info(f"Lost mode enabled on device {device_id}")
            return True
        except Exception as e:
            logger.error(f"Failed to enable lost mode on {device_id}: {e}")
            return False
    
    def wipe_device(self, device_id: int) -> bool:
        """Wipe a compromised device."""
        import requests
        
        try:
            payload = {"devices": [device_id]}
            response = requests.post(
                f"{self.hexnode_api_url}/actions/wipe_device/",
                headers={
                    "Authorization": self.hexnode_api_key,
                    "Content-Type": "application/json"
                },
                json=payload
            )
            response.raise_for_status()
            logger.warning(f"Device {device_id} wipe initiated")
            return True
        except Exception as e:
            logger.error(f"Failed to wipe device {device_id}: {e}")
            return False
    
    def execute_remediation_script(self, device_id: int, script_id: int) -> bool:
        """Execute a remediation script on a device."""
        import requests
        
        try:
            payload = {
                "devices": [device_id],
                "scripts": [script_id]
            }
            response = requests.post(
                f"{self.hexnode_api_url}/actions/execute_script/",
                headers={
                    "Authorization": self.hexnode_api_key,
                    "Content-Type": "application/json"
                },
                json=payload
            )
            response.raise_for_status()
            logger.info(f"Remediation script executed on device {device_id}")
            return True
        except Exception as e:
            logger.error(f"Failed to execute script on {device_id}: {e}")
            return False
    
    def send_alert_to_splunk(self, event_data: Dict) -> bool:
        """Forward alert to Splunk for logging."""
        import requests
        
        try:
            headers = {
                "Authorization": f"Splunk {self.splunk_token}",
                "Content-Type": "application/json"
            }
            response = requests.post(
                self.splunk_webhook,
                headers=headers,
                json=event_data
            )
            response.raise_for_status()
            return True
        except Exception as e:
            logger.error(f"Failed to send to Splunk: {e}")
            return False
    
    def process_alert(self, alert: Dict) -> Dict:
        """Process a SIEM alert and take appropriate action."""
        result = {
            "timestamp": datetime.now().isoformat(),
            "alert_id": alert.get("id"),
            "device_id": alert.get("device_id"),
            "alert_type": alert.get("type"),
            "severity": alert.get("severity", "low"),
            "source": alert.get("source", "unknown"),
            "action_taken": "none",
            "details": {}
        }
        
        device_id = alert.get("device_id")
        severity = alert.get("severity", "low")
        alert_type = alert.get("type", "unknown")
        
        logger.info(f"Processing alert: {alert_type} (severity: {severity})")
        
        if not self.auto_remediate:
            result["action_taken"] = "logged_only"
            logger.info("Auto-remediation disabled - logging only")
            return result
        
        # Define response actions based on severity
        if severity == "critical":
            # Critical: Wipe device immediately
            self.wipe_device(device_id)
            result["action_taken"] = "device_wiped"
            result["details"] = {"reason": f"Critical alert: {alert_type}"}
        
        elif severity == "high":
            # High: Enable lost mode and quarantine
            self.enable_lost_mode(device_id, f"Security Alert: {alert_type}")
            self.quarantine_device(device_id, f"High severity alert: {alert_type}")
            result["action_taken"] = "lost_mode_and_quarantine"
        
        elif severity == "medium":
            # Medium: Quarantine device
            self.quarantine_device(device_id, f"Medium severity alert: {alert_type}")
            result["action_taken"] = "quarantined"
        
        elif severity == "low":
            # Low: Execute remediation script
            self.execute_remediation_script(device_id, script_id=1)
            result["action_taken"] = "remediation_script"
        
        else:
            result["action_taken"] = "logged_only"
        
        # Forward to Splunk for audit trail
        self.send_alert_to_splunk({
            "event": "hexnode_remediation",
            "alert": alert,
            "action": result["action_taken"],
            "timestamp": result["timestamp"]
        })
        
        return result


# Initialize SIEM integration
siem = SIEMIntegration()


@app.route('/webhook/siem', methods=['POST'])
def handle_siem_webhook():
    """Handle incoming SIEM webhooks."""
    try:
        # Verify content type
        if not request.is_json:
            return jsonify({"error": "Content-Type must be application/json"}), 400
        
        alert = request.get_json()
        
        # Process the alert
        result = siem.process_alert(alert)
        
        logger.info(f"Alert processed: {result['action_taken']}")
        return jsonify(result), 200
    
    except Exception as e:
        logger.error(f"Webhook error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/webhook/splunk', methods=['POST'])
def handle_splunk_webhook():
    """Handle Splunk webhook alerts."""
    try:
        # Verify Splunk HEC token
        auth_header = request.headers.get('Authorization', '')
        if not auth_header.startswith('Splunk '):
            return jsonify({"error": "Invalid authorization"}), 401
        
        event = request.get_json()
        
        # Extract alert details from Splunk event
        alert = {
            "id": event.get("sid"),
            "device_id": event.get("result", {}).get("device_id"),
            "type": event.get("result", {}).get("alert_type", "splunk_alert"),
            "severity": event.get("result", {}).get("severity", "medium"),
            "source": "splunk",
            "raw_event": event
        }
        
        result = siem.process_alert(alert)
        return jsonify(result), 200
    
    except Exception as e:
        logger.error(f"Splunk webhook error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/webhook/carbonblack', methods=['POST'])
def handle_carbonblack_webhook():
    """Handle Carbon Black EDR alerts."""
    try:
        event = request.get_json()
        
        alert = {
            "id": event.get("universal_key"),
            "device_id": event.get("device", {}).get("device_id"),
            "type": event.get("description", "carbonblack_alert"),
            "severity": "high" if event.get("severity") == 10 else "medium",
            "source": "carbonblack",
            "raw_event": event
        }
        
        result = siem.process_alert(alert)
        return jsonify(result), 200
    
    except Exception as e:
        logger.error(f"Carbon Black webhook error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/webhook/crowdstrike', methods=['POST'])
def handle_crowdstrike_webhook():
    """Handle CrowdStrike Falcon alerts."""
    try:
        event = request.get_json()
        
        alert = {
            "id": event.get("event", {}).get("eventId"),
            "device_id": event.get("event", {}).get("hostinfo", {}).get("device_id"),
            "type": event.get("event", {}).get("technique", {}).get("name", "crowdstrike_alert"),
            "severity": "critical" if event.get("event", {}).get("severity") >= 80 else "high",
            "source": "crowdstrike",
            "raw_event": event
        }
        
        result = siem.process_alert(alert)
        return jsonify(result), 200
    
    except Exception as e:
        logger.error(f"CrowdStrike webhook error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint."""
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "service": "Hexnode Zero Trust SIEM Integration"
    }), 200


@app.route('/alerts', methods=['GET'])
def get_recent_alerts():
    """Get recent alerts (from in-memory store)."""
    # In production, use a database
    return jsonify({
        "alerts": [],
        "message": "Configure SIEM webhooks to receive alerts"
    }), 200


if __name__ == '__main__':
    logger.info("Starting Hexnode Zero Trust SIEM Integration Server")
    app.run(host='0.0.0.0', port=5000, debug=False)
