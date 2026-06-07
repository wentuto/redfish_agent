"""
Multicast Discovery Module for CDU Firmware Updater
Provides master (discovery) and slave (responder) functionality
"""

import socket
import struct
import json
import time
import random
import threading
from typing import List, Dict, Callable, Optional


class MulticastMaster:
    """
    Multicast Master for device discovery
    Sends discovery requests and collects responses from slaves
    """
    
    def __init__(
        self,
        multicast_group: str = '224.0.0.250', #'239.255.10.10','239.255.255.250'
        port: int = 19200,
        total_rounds: int = 3,
        interval: int = 1,
        response_timeout: int = 5,
        callback: Optional[Callable[[Dict], None]] = None,
        protocol_id: str = "DEV_DISCOVERY_V1",
        secret: Optional[str] = None
    ):
        """
        Initialize MulticastMaster
        
        Args:
            multicast_group: Multicast IP address (239.x.x.x for private use)
            port: UDP port number
            total_rounds: Number of discovery rounds
            interval: Seconds between rounds
            response_timeout: Timeout for receiving responses
            callback: Optional callback function called when a device is discovered
        """
        self.multicast_group = multicast_group
        self.port = port
        self.total_rounds = total_rounds
        self.interval = interval
        self.response_timeout = response_timeout
        self.callback = callback
        self.protocol_id = protocol_id
        self.secret = secret
        
        self.all_devices = []
        self.response_times = []
        self.is_running = False
        self.sock = None
        
    def setup_socket(self):
        """Setup UDP socket for multicast"""
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM, socket.IPPROTO_UDP)
        self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.sock.setsockopt(socket.IPPROTO_IP, socket.IP_MULTICAST_TTL, 2)
        # 增加接收緩衝區以處理大量回應
        self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_RCVBUF, 8 * 1024 * 1024)  # 8MB
        # 綁定到任意地址和端口（讓系統分配）以接收 unicast 回應
        self.sock.bind(('', 0))
        self.sock.settimeout(self.response_timeout)
        
    def discover(self) -> List[Dict]:
        """
        Execute discovery process
        
        Returns:
            List of discovered devices with their information
        """
        self.is_running = True
        self.all_devices = []
        self.response_times = []
        
        try:
            self.setup_socket()
            
            for round_num in range(1, self.total_rounds + 1):
                if not self.is_running:
                    break
                    
                devices = self._discover_round(round_num)
                
                # 將新設備加入總列表 (使用 IP 判斷唯一性)
                existing_ips = {d.get('ip') for d in self.all_devices}
                for device in devices:
                    if device.get('ip') not in existing_ips:
                        self.all_devices.append(device)
                        existing_ips.add(device.get('ip'))
                        
                        # 呼叫回調函數
                        if self.callback:
                            self.callback(device)
                
                # 等待下一輪 (除了最後一輪)
                if round_num < self.total_rounds and self.is_running:
                    time.sleep(self.interval)
                    
        finally:
            if self.sock:
                self.sock.close()
            self.is_running = False
            
        return self.all_devices
    
    def _discover_round(self, round_num: int) -> List[Dict]:
        """
        Execute a single discovery round
        
        Args:
            round_num: Current round number
            
        Returns:
            List of devices discovered in this round
        """
        req = {
            "type": "DISCOVER_REQ",
            "from": "controller",
            "seq": round_num,
            "protocol_id": self.protocol_id,
            "secret": self.secret
        }
        
        start_time = time.time()
        self.sock.sendto(json.dumps(req).encode(), (self.multicast_group, self.port))
        
        devices = []
        ack_count = 0
        
        try:
            while True:
                data, addr = self.sock.recvfrom(4096)
                resp = json.loads(data.decode())
                devices.append(resp)
                
                # 發送 ACK 確認封包給 slave
                ack = {
                    "type": "DISCOVER_ACK",
                    "seq": resp.get("seq"),
                    "from": "controller"
                }
                try:
                    print(f"Received response from {resp.get('ip')}:{resp.get('api_port')}, sending ACK...")
                    print(f"ACK content: {ack}, addr: {addr}")
                    self.sock.sendto(json.dumps(ack).encode(), addr)
                    ack_count += 1
                except Exception:
                    pass  # 忽略 ACK 發送失敗
                    
        except socket.timeout:
            pass
        
        elapsed = time.time() - start_time
        self.response_times.append(elapsed)
        
        return devices
    
    def stop(self):
        """Stop the discovery process"""
        self.is_running = False
        
    def get_statistics(self) -> Dict:
        """
        Get discovery statistics
        
        Returns:
            Dictionary containing statistics
        """
        avg_time = sum(self.response_times) / len(self.response_times) if self.response_times else 0
        return {
            'total_devices': len(self.all_devices),
            'total_rounds': len(self.response_times),
            'average_response_time': avg_time,
            'response_times': self.response_times
        }


class MulticastSlave:
    """
    Multicast Slave for responding to discovery requests
    Used for testing and simulation
    """
    
    def __init__(
        self,
        multicast_group: str = '239.255.255.250',
        port: int = 19200,
        device_id: Optional[str] = None,
        api_port: Optional[int] = None,
        pause_duration: int = 20,
        protocol_id: str = "DEV_DISCOVERY_V1",
        secret: Optional[str] = None
    ):
        """
        Initialize MulticastSlave
        
        Args:
            multicast_group: Multicast IP address
            port: UDP port number
            device_id: Unique device identifier (auto-generated if None)
            api_port: API port number (random if None)
            pause_duration: Seconds to pause after receiving ACK
        """
        self.multicast_group = multicast_group
        self.port = port
        self.device_id = device_id or f"CDU-{random.randint(1, 999):03d}"
        self.api_port = api_port or random.randint(8443, 8543)
        self.pause_duration = pause_duration
        self.protocol_id = protocol_id
        self.secret = secret
        
        self.is_running = False
        self.paused_until = 0
        self.sock = None
        
    def setup_socket(self):
        """Setup UDP socket for multicast"""
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM, socket.IPPROTO_UDP)
        self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEPORT, 1)
        self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_RCVBUF, 256 * 1024)  # 256KB
        self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_SNDBUF, 256 * 1024)  # 256KB
        self.sock.bind(('', self.port))
        
        # 加入 multicast 群組
        mreq = struct.pack(
            "4s4s",
            socket.inet_aton(self.multicast_group),
            socket.inet_aton("0.0.0.0")
        )
        self.sock.setsockopt(socket.IPPROTO_IP, socket.IP_ADD_MEMBERSHIP, mreq)
        
    def start(self):
        """Start listening for discovery requests"""
        self.is_running = True
        self.setup_socket()
        
        try:
            while self.is_running:
                # 檢查是否需要暫停
                current_time = time.time()
                if current_time < self.paused_until:
                    full_remaining = self.paused_until - current_time
                    # 隨機睡眠時間：1/2 到 1 倍之間
                    remaining = random.uniform(full_remaining * 0.5, full_remaining)
                    time.sleep(remaining)
                    self.paused_until = 0
                    
                    # 清空接收緩衝區
                    self._clear_buffer()
                
                # 接收封包
                try:
                    data, addr = self.sock.recvfrom(2048)
                    msg = json.loads(data.decode())
                    
                    if msg.get("type") == "DISCOVER_REQ":
                        # Validate protocol_id and secret
                        if msg.get("protocol_id") != self.protocol_id:
                            continue  # Ignore requests with wrong protocol_id
                        if self.secret is not None and msg.get("secret") != self.secret:
                            continue  # Ignore requests with wrong secret
                        self._handle_discovery_request(msg, addr)
                    elif msg.get("type") == "DISCOVER_ACK":
                        self._handle_ack(msg)
                        
                except socket.timeout:
                    continue
                    
        finally:
            if self.sock:
                self.sock.close()
            self.is_running = False
            
    def _handle_discovery_request(self, msg: Dict, addr: tuple):
        """Handle discovery request from master"""
        # 加入隨機延遲避免封包碰撞
        delay = random.uniform(0.001, 0.5)
        time.sleep(delay)
        
        resp = {
            "type": "DISCOVER_RESP",
            "ip": socket.gethostbyname(socket.gethostname()),
            "api_port": self.api_port,
            "seq": msg.get("seq", 0)
        }
        
        self.sock.sendto(json.dumps(resp).encode(), addr)
        
        # 等待 ACK
        self._wait_for_ack(msg, addr)
        
    def _wait_for_ack(self, original_msg: Dict, master_addr: tuple):
        """Wait for ACK from master"""
        self.sock.settimeout(0.5)
        ack_start = time.time()
        
        try:
            while time.time() - ack_start < 0.5:
                ack_data, ack_addr = self.sock.recvfrom(2048)
                ack_msg = json.loads(ack_data.decode())
                
                if (ack_msg.get("type") == "DISCOVER_ACK" and
                    ack_msg.get("seq") == original_msg.get("seq") and
                    ack_addr == master_addr):
                    self.paused_until = time.time() + self.pause_duration
                    break
                    
        except socket.timeout:
            pass
        finally:
            self.sock.settimeout(None)
            
    def _handle_ack(self, msg: Dict):
        """Handle ACK received outside of wait period"""
        self.paused_until = time.time() + self.pause_duration
        
    def _clear_buffer(self):
        """Clear socket receive buffer"""
        self.sock.setblocking(False)
        try:
            while True:
                self.sock.recvfrom(2048)
        except BlockingIOError:
            pass
        finally:
            self.sock.setblocking(True)
            
    def stop(self):
        """Stop the slave"""
        self.is_running = False


def discover_devices(
    multicast_group: str = '239.255.255.250',
    port: int = 19200,
    total_rounds: int = 5,
    interval: int = 2,
    response_timeout: int = 5,
    callback: Optional[Callable[[Dict], None]] = None
) -> List[Dict]:
    """
    Convenience function for device discovery
    
    Args:
        multicast_group: Multicast IP address
        port: UDP port number
        total_rounds: Number of discovery rounds
        interval: Seconds between rounds
        response_timeout: Timeout for receiving responses
        callback: Optional callback function for each discovered device
        
    Returns:
        List of discovered devices
    """
    master = MulticastMaster(
        multicast_group=multicast_group,
        port=port,
        total_rounds=total_rounds,
        interval=interval,
        response_timeout=response_timeout,
        callback=callback
    )
    return master.discover()
