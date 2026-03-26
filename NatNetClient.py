

# file: NatNetClient.py
# Minimal NatNet client for OptiTrack Motive streaming
# Compatible with NatNet 3.x / Python 3.x

import socket
import struct
import sys
import threading

class NatNetClient:
    def __init__(self, server_ip="127.0.0.1", multicast_ip="239.255.42.99", command_port=1510, data_port=1511):
        self.server_ip = server_ip
        self.multicast_ip = multicast_ip
        self.command_port = command_port
        self.data_port = data_port

        # Listeners (callbacks you can assign)
        self.rigidBodyListener = None
        self.markerListener = None

        # Internal sockets
        self.data_socket = None
        self.command_socket = None

    def run(self):
        # Create data socket
        self.data_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self.data_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)

        try:
            self.data_socket.bind(('', self.data_port))
        except socket.error as msg:
            print("Bind failed:", msg)
            sys.exit()

        mreq = struct.pack("4sl", socket.inet_aton(self.multicast_ip), socket.INADDR_ANY)
        self.data_socket.setsockopt(socket.IPPROTO_IP, socket.IP_ADD_MEMBERSHIP, mreq)

        # Launch listener thread
        thread = threading.Thread(target=self._listen)
        thread.daemon = True
        thread.start()

    def _listen(self):
        while True:
            try:
                data, addr = self.data_socket.recvfrom(32768)  # 32k buffer
                self._parse_packet(data)
            except Exception as e:
                print("Error receiving data:", e)

    def _parse_packet(self, data):
        # Simple parser for rigid bodies and markers
        # (NatNet has many packet types, this is minimal)

        try:
            message_id, packet_size = struct.unpack("hh", data[0:4])
        except:
            return

        # 7 = Frame of Mocap Data
        if message_id == 7:
            offset = 4

            # Frame number
            frame_number = struct.unpack("i", data[offset:offset+4])[0]
            offset += 4

            # Marker sets count
            marker_set_count = struct.unpack("i", data[offset:offset+4])[0]
            offset += 4

            for i in range(marker_set_count):
                # Skip marker set name (null-terminated string)
                name = b''
                while data[offset:offset+1] != b'\0':
                    name += data[offset:offset+1]
                    offset += 1
                offset += 1  # null
                marker_count = struct.unpack("i", data[offset:offset+4])[0]
                offset += 4
                offset += marker_count * 12  # skip markers

            # Unlabeled markers
            unlabeled_count = struct.unpack("i", data[offset:offset+4])[0]
            offset += 4
            for i in range(unlabeled_count):
                pos = struct.unpack("fff", data[offset:offset+12])
                offset += 12
                if self.markerListener is not None:
                    self.markerListener(i, pos)

            # Rigid bodies
            rigid_body_count = struct.unpack("i", data[offset:offset+4])[0]
            offset += 4
            for j in range(rigid_body_count):
                rb_id = struct.unpack("i", data[offset:offset+4])[0]
                offset += 4
                pos = struct.unpack("fff", data[offset:offset+12])
                offset += 12
                rot = struct.unpack("ffff", data[offset:offset+16])  # quaternion
                offset += 16
                if self.rigidBodyListener is not None:
                    self.rigidBodyListener(rb_id, pos, rot)