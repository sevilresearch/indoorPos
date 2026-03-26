"""
Crazyflie 2.1+ Lighthouse + OptiTrack Dual Logger GUI
Works with the custom NatNetClient.py in the same folder.

Motive setup:
- Broadcast Frame Data = ON
- Rigid Bodies = ON
- Local Interface = Loopback / 127.0.0.1
- If multiple rigid bodies are streamed, enter the Rigid Body ID for Robot_3 in the GUI

Install:
    pip install cflib pandas openpyxl
"""

import csv
import math
import os
import threading
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

import cflib.crtp
from cflib.crazyflie import Crazyflie
from cflib.crazyflie.log import LogConfig
from cflib.crazyflie.syncCrazyflie import SyncCrazyflie
from cflib.positioning.position_hl_commander import PositionHlCommander

try:
    import pandas as pd
except Exception:
    pd = None

try:
    from NatNetClient import NatNetClient
    NATNET_AVAILABLE = True
except Exception:
    NatNetClient = None
    NATNET_AVAILABLE = False


DEFAULT_URI = "radio://0/80/2M/E7E7E7E7C3"

LIVE_LOG_PERIOD_MS = 100
SHAPE_LOG_PERIOD_MS = 10
EXPECTED_LOG_PERIOD_MS = 10
LOSS_GAP_THRESHOLD_MS = 30
FREEZE_EPSILON_M = 0.0005
OT_FRESHNESS_LIMIT_MS = 30

# Change only if OptiTrack axes do not match Lighthouse
OT_AXIS_ORDER = ("x", "y", "z")
OT_AXIS_SIGN = (1.0, 1.0, 1.0)
OT_OFFSET_M = (0.0, 0.0, 0.0)


class CrazyflieDualLoggerGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Crazyflie Lighthouse + OptiTrack Dual Logger")
        self.root.geometry("1240x900")

        self.scf = None
        self.cf = None
        self.pc = None

        self.connected = False
        self.in_air = False
        self.busy = False

        self.shape_logconf = None
        self.live_logconf = None

        self.logging_active = False
        self.log_rows = []
        self.current_shape_name = None
        self.current_shape_size = None
        self.shape_start_time = None

        # Lighthouse loss tracking
        self.prev_lh_timestamp = None
        self.prev_lh_xyz = None
        self.total_lh_loss_time_ms = 0.0
        self.lh_loss_event_count = 0

        # Hover/origin
        self.hover_x = 0.0
        self.hover_y = 0.0
        self.hover_z = 0.5
        self.origin_x = 0.0
        self.origin_y = 0.0
        self.origin_z = 0.6

        # Live Lighthouse
        self.live_lh_x = 0.0
        self.live_lh_y = 0.0
        self.live_lh_z = 0.0

        # OptiTrack
        self.ot_client = None
        self.ot_connected = False
        self.ot_streaming_ok = False
        self.ot_rigid_body_seen = False
        self.ot_frame_count = 0

        self.ot_latest_pc_time_s = None
        self.ot_latest_frame_id = None
        self.ot_latest_rb_id = None
        self.ot_latest_x = None
        self.ot_latest_y = None
        self.ot_latest_z = None
        self.ot_latest_tracked = False

        self.live_ot_x = None
        self.live_ot_y = None
        self.live_ot_z = None
        self.live_ot_tracked = False

        self.ot_prev_loss_flag = 0
        self.total_ot_loss_time_ms = 0.0
        self.ot_loss_event_count = 0

        self.data_lock = threading.Lock()

        self._build_gui()
        cflib.crtp.init_drivers(enable_debug_driver=False)

    # -------------------------------------------------
    # GUI
    # -------------------------------------------------

    def _build_gui(self):
        pad = {"padx": 8, "pady": 6}

        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True, padx=12, pady=12)

        settings_frame = ttk.LabelFrame(main, text="Settings")
        settings_frame.pack(fill="x", **pad)

        ttk.Label(settings_frame, text="Crazyflie URI").grid(row=0, column=0, sticky="w", **pad)
        self.uri_var = tk.StringVar(value=DEFAULT_URI)
        ttk.Entry(settings_frame, textvariable=self.uri_var, width=34).grid(row=0, column=1, sticky="w", **pad)

        ttk.Label(settings_frame, text="Takeoff Height (m)").grid(row=1, column=0, sticky="w", **pad)
        self.height_var = tk.StringVar(value="0.6")
        ttk.Entry(settings_frame, textvariable=self.height_var, width=12).grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(settings_frame, text="Shape Size (m)").grid(row=2, column=0, sticky="w", **pad)
        self.size_var = tk.StringVar(value="0.6")
        ttk.Entry(settings_frame, textvariable=self.size_var, width=12).grid(row=2, column=1, sticky="w", **pad)

        ttk.Label(settings_frame, text="Flight Speed (m/s)").grid(row=3, column=0, sticky="w", **pad)
        self.speed_var = tk.StringVar(value="0.3")
        ttk.Entry(settings_frame, textvariable=self.speed_var, width=12).grid(row=3, column=1, sticky="w", **pad)

        ttk.Label(settings_frame, text="Rigid Body Name").grid(row=0, column=2, sticky="w", **pad)
        self.ot_body_name_var = tk.StringVar(value="Robot_3")
        ttk.Entry(settings_frame, textvariable=self.ot_body_name_var, width=18).grid(row=0, column=3, sticky="w", **pad)

        ttk.Label(settings_frame, text="Rigid Body ID (optional)").grid(row=1, column=2, sticky="w", **pad)
        self.ot_body_id_var = tk.StringVar(value="")
        ttk.Entry(settings_frame, textvariable=self.ot_body_id_var, width=18).grid(row=1, column=3, sticky="w", **pad)

        natnet_msg = "NatNetClient.py found" if NATNET_AVAILABLE else "NatNetClient.py NOT found"
        self.natnet_status_var = tk.StringVar(value=natnet_msg)
        ttk.Label(settings_frame, textvariable=self.natnet_status_var).grid(row=2, column=2, columnspan=2, sticky="w", **pad)

        self.ot_status_var = tk.StringVar(value="OptiTrack: not connected")
        ttk.Label(settings_frame, textvariable=self.ot_status_var).grid(row=3, column=2, columnspan=2, sticky="w", **pad)

        self.record_ready_var = tk.StringVar(value="Record Ready: No")
        ttk.Label(settings_frame, textvariable=self.record_ready_var).grid(row=4, column=2, columnspan=2, sticky="w", **pad)

        waypoint_frame = ttk.LabelFrame(main, text="Waypoint")
        waypoint_frame.pack(fill="x", **pad)

        ttk.Label(waypoint_frame, text="X (m)").grid(row=0, column=0, sticky="w", **pad)
        self.wp_x_var = tk.StringVar(value="0.0")
        ttk.Entry(waypoint_frame, textvariable=self.wp_x_var, width=10).grid(row=0, column=1, sticky="w", **pad)

        ttk.Label(waypoint_frame, text="Y (m)").grid(row=0, column=2, sticky="w", **pad)
        self.wp_y_var = tk.StringVar(value="0.0")
        ttk.Entry(waypoint_frame, textvariable=self.wp_y_var, width=10).grid(row=0, column=3, sticky="w", **pad)

        ttk.Label(waypoint_frame, text="Z (m)").grid(row=0, column=4, sticky="w", **pad)
        self.wp_z_var = tk.StringVar(value="0.6")
        ttk.Entry(waypoint_frame, textvariable=self.wp_z_var, width=10).grid(row=0, column=5, sticky="w", **pad)

        ttk.Button(waypoint_frame, text="Go To Waypoint", command=self.goto_waypoint_clicked).grid(row=0, column=6, **pad)
        ttk.Button(waypoint_frame, text="Return To Origin", command=self.return_to_origin_clicked).grid(row=0, column=7, **pad)

        ot_control_frame = ttk.LabelFrame(main, text="OptiTrack Controls")
        ot_control_frame.pack(fill="x", **pad)

        ttk.Button(ot_control_frame, text="Test OptiTrack", command=self.test_optitrack_clicked).grid(row=0, column=0, **pad)
        ttk.Button(ot_control_frame, text="Reconnect OptiTrack", command=self.reconnect_optitrack_clicked).grid(row=0, column=1, **pad)
        ttk.Button(ot_control_frame, text="Stop OptiTrack", command=self.stop_optitrack_clicked).grid(row=0, column=2, **pad)

        live_frame = ttk.LabelFrame(main, text="Live Position")
        live_frame.pack(fill="x", **pad)

        self.live_lh_x_var = tk.StringVar(value="0.000")
        self.live_lh_y_var = tk.StringVar(value="0.000")
        self.live_lh_z_var = tk.StringVar(value="0.000")

        self.live_ot_x_var = tk.StringVar(value="nan")
        self.live_ot_y_var = tk.StringVar(value="nan")
        self.live_ot_z_var = tk.StringVar(value="nan")
        self.live_ot_track_var = tk.StringVar(value="No")
        self.live_ot_id_var = tk.StringVar(value="")

        ttk.Label(live_frame, text="Lighthouse X").grid(row=0, column=0, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_lh_x_var, width=10).grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(live_frame, text="Lighthouse Y").grid(row=0, column=2, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_lh_y_var, width=10).grid(row=0, column=3, sticky="w", **pad)
        ttk.Label(live_frame, text="Lighthouse Z").grid(row=0, column=4, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_lh_z_var, width=10).grid(row=0, column=5, sticky="w", **pad)

        ttk.Label(live_frame, text="OptiTrack X").grid(row=1, column=0, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_ot_x_var, width=10).grid(row=1, column=1, sticky="w", **pad)
        ttk.Label(live_frame, text="OptiTrack Y").grid(row=1, column=2, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_ot_y_var, width=10).grid(row=1, column=3, sticky="w", **pad)
        ttk.Label(live_frame, text="OptiTrack Z").grid(row=1, column=4, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_ot_z_var, width=10).grid(row=1, column=5, sticky="w", **pad)
        ttk.Label(live_frame, text="Tracked").grid(row=1, column=6, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_ot_track_var, width=8).grid(row=1, column=7, sticky="w", **pad)
        ttk.Label(live_frame, text="RB ID").grid(row=1, column=8, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_ot_id_var, width=8).grid(row=1, column=9, sticky="w", **pad)

        control_frame = ttk.LabelFrame(main, text="Flight Controls")
        control_frame.pack(fill="x", **pad)

        ttk.Button(control_frame, text="Connect Crazyflie", command=self.connect_clicked).grid(row=0, column=0, **pad)
        ttk.Button(control_frame, text="Disconnect Crazyflie", command=self.disconnect_clicked).grid(row=0, column=1, **pad)
        ttk.Button(control_frame, text="Take Off", command=self.takeoff_clicked).grid(row=0, column=2, **pad)
        ttk.Button(control_frame, text="Land", command=self.land_clicked).grid(row=0, column=3, **pad)

        ttk.Button(control_frame, text="Fly Square", command=lambda: self.shape_clicked("square")).grid(row=1, column=0, **pad)
        ttk.Button(control_frame, text="Fly Circle", command=lambda: self.shape_clicked("circle")).grid(row=1, column=1, **pad)
        ttk.Button(control_frame, text="Fly Triangle", command=lambda: self.shape_clicked("triangle")).grid(row=1, column=2, **pad)
        ttk.Button(control_frame, text="EMERGENCY STOP", command=self.emergency_stop_clicked).grid(row=1, column=3, **pad)

        status_frame = ttk.LabelFrame(main, text="Status")
        status_frame.pack(fill="both", expand=True, **pad)

        self.status_text = tk.Text(status_frame, height=24, wrap="word")
        self.status_text.pack(fill="both", expand=True, padx=8, pady=8)

        self._status("Ready.")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _status(self, msg: str):
        stamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{stamp}] {msg}\n"

        def write():
            self.status_text.insert("end", line)
            self.status_text.see("end")

        self.root.after(0, write)

    def _safe_float(self, var: tk.StringVar, name: str, minimum: float = None) -> float:
        try:
            value = float(var.get())
        except ValueError:
            raise ValueError(f"{name} must be a number.")
        if minimum is not None and value < minimum:
            raise ValueError(f"{name} must be at least {minimum}.")
        return value

    def _safe_optional_int(self, var: tk.StringVar):
        text = var.get().strip()
        if text == "":
            return None
        return int(text)

    def _run_threaded(self, target):
        if self.busy:
            self._status("Busy. Wait for the current action to finish.")
            return
        threading.Thread(target=target, daemon=True).start()

    def _format_filename_number(self, value: float) -> str:
        return f"{value:.2f}".replace(".", "p")

    def _sanitize_filename_part(self, text: str) -> str:
        return "".join(c if c.isalnum() or c in ("_", "-") else "_" for c in text)

    def _update_ready_status(self):
        ready = self.connected and self.ot_connected and self.ot_streaming_ok and self.ot_rigid_body_seen
        self.record_ready_var.set("Record Ready: Yes" if ready else "Record Ready: No")

        body_name = self.ot_body_name_var.get().strip() or "Robot_3"
        body_id_text = self.ot_body_id_var.get().strip()

        if not NATNET_AVAILABLE:
            self.ot_status_var.set("OptiTrack: NatNetClient.py missing")
        elif not self.ot_connected:
            self.ot_status_var.set("OptiTrack: not connected")
        elif self.ot_connected and not self.ot_streaming_ok:
            self.ot_status_var.set("OptiTrack: connected, waiting for frames")
        elif self.ot_connected and self.ot_streaming_ok and not self.ot_rigid_body_seen:
            if body_id_text:
                self.ot_status_var.set(f"OptiTrack: streaming, RB ID {body_id_text} not seen")
            else:
                self.ot_status_var.set(f"OptiTrack: streaming, waiting for rigid body ({body_name})")
        else:
            tracked_str = "tracked" if self.live_ot_tracked else "not tracked"
            self.ot_status_var.set(f"OptiTrack: ready, {tracked_str}")

    # -------------------------------------------------
    # Crazyflie connection
    # -------------------------------------------------

    def connect_clicked(self):
        self._run_threaded(self.connect_cf)

    def disconnect_clicked(self):
        self._run_threaded(self.disconnect_cf)

    def connect_cf(self):
        self.busy = True
        try:
            if self.connected:
                self._status("Crazyflie already connected.")
                return

            uri = self.uri_var.get().strip()
            self._status(f"Connecting to {uri} ...")

            self.scf = SyncCrazyflie(uri, cf=Crazyflie(rw_cache="./cache"))
            self.scf.open_link()
            self.cf = self.scf.cf

            time.sleep(1.0)

            try:
                self.cf.platform.send_arming_request(True)
                time.sleep(1.0)
                self._status("Arming request sent.")
            except Exception as e:
                self._status(f"Arming request skipped: {e}")

            try:
                self.cf.param.set_value("stabilizer.estimator", "2")
                time.sleep(0.2)
            except Exception as e:
                self._status(f"Could not set stabilizer.estimator=2: {e}")

            self.reset_estimator()
            self.start_live_logging()

            self.connected = True
            self._status("Crazyflie connected successfully.")
            self._update_ready_status()

        except Exception as e:
            self._status(f"Crazyflie connection failed: {e}")
            self.connected = False
            self.scf = None
            self.cf = None
            self.pc = None
            self._update_ready_status()
        finally:
            self.busy = False

    def disconnect_cf(self):
        self.busy = True
        try:
            if self.logging_active:
                self.stop_shape_logging()

            self.stop_live_logging()

            if self.in_air:
                self._status("Drone is in the air. Land before disconnecting.")
                return

            if self.scf is not None:
                self.scf.close_link()
                self._status("Crazyflie disconnected.")
            else:
                self._status("Crazyflie was not connected.")

            self.connected = False
            self.scf = None
            self.cf = None
            self.pc = None
            self._update_ready_status()

        except Exception as e:
            self._status(f"Disconnect error: {e}")
        finally:
            self.busy = False

    def reset_estimator(self):
        if not self.cf:
            return
        self._status("Resetting Kalman estimator...")
        self.cf.param.set_value("kalman.resetEstimation", "1")
        time.sleep(0.1)
        self.cf.param.set_value("kalman.resetEstimation", "0")
        time.sleep(2.0)
        self._status("Estimator reset complete.")

    # -------------------------------------------------
    # Lighthouse live logging
    # -------------------------------------------------

    def start_live_logging(self):
        if not self.cf:
            return

        try:
            self.live_logconf = LogConfig(name="LivePosition", period_in_ms=LIVE_LOG_PERIOD_MS)
            self.live_logconf.add_variable("stateEstimate.x", "float")
            self.live_logconf.add_variable("stateEstimate.y", "float")
            self.live_logconf.add_variable("stateEstimate.z", "float")
            self.live_logconf.data_received_cb.add_callback(self._live_log_callback)
            self.cf.log.add_config(self.live_logconf)
            self.live_logconf.start()
            self._status("Live Lighthouse display started.")
        except Exception as e:
            self.live_logconf = None
            self._status(f"Could not start live Lighthouse display: {e}")

    def stop_live_logging(self):
        try:
            if self.live_logconf is not None:
                try:
                    self.live_logconf.stop()
                except Exception:
                    pass
                try:
                    self.cf.log.delete_config(self.live_logconf)
                except Exception:
                    pass
        finally:
            self.live_logconf = None

    def _live_log_callback(self, timestamp, data, logconf):
        with self.data_lock:
            self.live_lh_x = float(data["stateEstimate.x"])
            self.live_lh_y = float(data["stateEstimate.y"])
            self.live_lh_z = float(data["stateEstimate.z"])
        self.root.after(0, self._update_live_labels)

    # -------------------------------------------------
    # OptiTrack / NatNet
    # -------------------------------------------------

    def test_optitrack_clicked(self):
        self._run_threaded(self.test_optitrack)

    def reconnect_optitrack_clicked(self):
        self._run_threaded(self.reconnect_optitrack)

    def stop_optitrack_clicked(self):
        self._run_threaded(self.stop_optitrack_manual)

    def test_optitrack(self):
        self.busy = True
        try:
            self._status("Testing OptiTrack ...")
            self._start_optitrack_client()

            timeout_s = 3.0
            t0 = time.time()
            while time.time() - t0 < timeout_s:
                with self.data_lock:
                    connected = self.ot_connected
                    streaming = self.ot_streaming_ok
                    seen = self.ot_rigid_body_seen
                if connected and streaming and seen:
                    break
                time.sleep(0.1)

            with self.data_lock:
                connected = self.ot_connected
                streaming = self.ot_streaming_ok
                seen = self.ot_rigid_body_seen

            if connected and streaming and seen:
                self._status("OptiTrack test passed.")
            elif connected and streaming:
                self._status("OptiTrack is streaming, but target rigid body has not been seen yet.")
            elif connected:
                self._status("OptiTrack connected, but no frame data received yet.")
            else:
                self._status("OptiTrack test failed.")

            self._update_ready_status()
        finally:
            self.busy = False

    def reconnect_optitrack(self):
        self.busy = True
        try:
            self._status("Reconnecting OptiTrack ...")
            self._stop_optitrack_client()
            time.sleep(0.5)
            self._start_optitrack_client()
            self._status("Reconnect finished.")
            self._update_ready_status()
        finally:
            self.busy = False

    def stop_optitrack_manual(self):
        self.busy = True
        try:
            self._stop_optitrack_client()
            self._status("OptiTrack stopped.")
            self._update_ready_status()
        finally:
            self.busy = False

    def _start_optitrack_client(self):
        if not NATNET_AVAILABLE:
            self._status("NatNetClient.py not found.")
            self.ot_connected = False
            self.ot_streaming_ok = False
            self.ot_rigid_body_seen = False
            self._update_ready_status()
            return

        self._stop_optitrack_client()

        try:
            self.ot_client = NatNetClient(server_ip="127.0.0.1")
            self.ot_client.newFrameListener = self._ot_new_frame_callback
            self.ot_client.rigidBodyListener = self._ot_rigid_body_callback

            ok = self.ot_client.run()
            if ok is False:
                self._status("OptiTrack client failed to start.")
                self.ot_connected = False
                self._update_ready_status()
                return

            with self.data_lock:
                self.ot_connected = True
                self.ot_streaming_ok = False
                self.ot_rigid_body_seen = False
                self.ot_frame_count = 0
                self.ot_latest_pc_time_s = None
                self.ot_latest_frame_id = None
                self.ot_latest_rb_id = None
                self.ot_latest_x = None
                self.ot_latest_y = None
                self.ot_latest_z = None
                self.ot_latest_tracked = False
                self.live_ot_x = None
                self.live_ot_y = None
                self.live_ot_z = None
                self.live_ot_tracked = False

            self._status("OptiTrack client started.")
            self.root.after(0, self._update_live_labels)
            self.root.after(0, self._update_ready_status)

        except Exception as e:
            self.ot_client = None
            self.ot_connected = False
            self.ot_streaming_ok = False
            self.ot_rigid_body_seen = False
            self._status(f"OptiTrack start failed: {e}")
            self._update_ready_status()

    def _stop_optitrack_client(self):
        try:
            if self.ot_client is not None:
                try:
                    self.ot_client.shutdown()
                except Exception:
                    pass
                self.ot_client = None
        finally:
            with self.data_lock:
                self.ot_connected = False
                self.ot_streaming_ok = False
                self.ot_rigid_body_seen = False
                self.ot_frame_count = 0
                self.ot_latest_pc_time_s = None
                self.ot_latest_frame_id = None
                self.ot_latest_rb_id = None
                self.ot_latest_x = None
                self.ot_latest_y = None
                self.ot_latest_z = None
                self.ot_latest_tracked = False
                self.live_ot_x = None
                self.live_ot_y = None
                self.live_ot_z = None
                self.live_ot_tracked = False
            self.root.after(0, self._update_live_labels)

    def _ot_new_frame_callback(self, frame_number):
        with self.data_lock:
            self.ot_streaming_ok = True
            self.ot_frame_count += 1
            self.ot_latest_frame_id = frame_number
        self.root.after(0, self._update_ready_status)

    def _ot_rigid_body_callback(self, rigid_body_id, position, rotation):
        try:
            target_id = self._safe_optional_int(self.ot_body_id_var)

            # If a target ID is entered, only accept that rigid body
            if target_id is not None and rigid_body_id != target_id:
                return

            x = float(position[0])
            y = float(position[1])
            z = float(position[2])

            x_cf, y_cf, z_cf = self._map_optitrack_xyz(x, y, z)
            now_pc_time_s = time.perf_counter()

            with self.data_lock:
                self.ot_connected = True
                self.ot_streaming_ok = True
                self.ot_rigid_body_seen = True
                self.ot_latest_pc_time_s = now_pc_time_s
                self.ot_latest_rb_id = rigid_body_id
                self.ot_latest_x = x_cf
                self.ot_latest_y = y_cf
                self.ot_latest_z = z_cf
                self.ot_latest_tracked = True

                self.live_ot_x = x_cf
                self.live_ot_y = y_cf
                self.live_ot_z = z_cf
                self.live_ot_tracked = True

            self.root.after(0, self._update_live_labels)
            self.root.after(0, self._update_ready_status)

        except Exception:
            pass

    def _map_optitrack_xyz(self, x, y, z):
        raw = {"x": x, "y": y, "z": z}
        ox = OT_AXIS_SIGN[0] * raw[OT_AXIS_ORDER[0]] + OT_OFFSET_M[0]
        oy = OT_AXIS_SIGN[1] * raw[OT_AXIS_ORDER[1]] + OT_OFFSET_M[1]
        oz = OT_AXIS_SIGN[2] * raw[OT_AXIS_ORDER[2]] + OT_OFFSET_M[2]
        return ox, oy, oz

    def _update_live_labels(self):
        with self.data_lock:
            self.live_lh_x_var.set(f"{self.live_lh_x:.3f}")
            self.live_lh_y_var.set(f"{self.live_lh_y:.3f}")
            self.live_lh_z_var.set(f"{self.live_lh_z:.3f}")

            self.live_ot_x_var.set("nan" if self.live_ot_x is None else f"{self.live_ot_x:.3f}")
            self.live_ot_y_var.set("nan" if self.live_ot_y is None else f"{self.live_ot_y:.3f}")
            self.live_ot_z_var.set("nan" if self.live_ot_z is None else f"{self.live_ot_z:.3f}")
            self.live_ot_track_var.set("Yes" if self.live_ot_tracked else "No")
            self.live_ot_id_var.set("" if self.ot_latest_rb_id is None else str(self.ot_latest_rb_id))

    # -------------------------------------------------
    # Shape logging
    # -------------------------------------------------

    def start_shape_logging(self, shape_name: str, shape_size: float):
        if not self.cf:
            return

        self.log_rows = []
        self.current_shape_name = shape_name
        self.current_shape_size = shape_size
        self.shape_start_time = time.time()

        self.prev_lh_timestamp = None
        self.prev_lh_xyz = None
        self.total_lh_loss_time_ms = 0.0
        self.lh_loss_event_count = 0

        self.ot_prev_loss_flag = 0
        self.total_ot_loss_time_ms = 0.0
        self.ot_loss_event_count = 0

        self.shape_logconf = LogConfig(name="ShapePositionLog", period_in_ms=SHAPE_LOG_PERIOD_MS)
        self.shape_logconf.add_variable("stateEstimate.x", "float")
        self.shape_logconf.add_variable("stateEstimate.y", "float")
        self.shape_logconf.add_variable("stateEstimate.z", "float")
        self.shape_logconf.data_received_cb.add_callback(self._shape_log_callback)
        self.cf.log.add_config(self.shape_logconf)
        self.shape_logconf.start()
        self.logging_active = True
        self._status(f"Started dual logging for {shape_name} at 100 Hz.")

    def _shape_log_callback(self, timestamp, data, logconf):
        if not self.logging_active:
            return

        lh_pc_time_s = time.perf_counter()
        lh_x = float(data["stateEstimate.x"])
        lh_y = float(data["stateEstimate.y"])
        lh_z = float(data["stateEstimate.z"])

        lh_gap_ms = 0
        lh_loss_flag = 0
        lh_estimated_lost_time_ms = 0
        lh_loss_reason = ""

        if self.prev_lh_timestamp is not None:
            lh_gap_ms = int(timestamp - self.prev_lh_timestamp)

            if lh_gap_ms > LOSS_GAP_THRESHOLD_MS:
                lh_loss_flag = 1
                lh_estimated_lost_time_ms = max(0, lh_gap_ms - EXPECTED_LOG_PERIOD_MS)
                self.total_lh_loss_time_ms += lh_estimated_lost_time_ms
                self.lh_loss_event_count += 1
                lh_loss_reason = "timestamp_gap"

            if self.prev_lh_xyz is not None:
                px, py, pz = self.prev_lh_xyz
                dist = math.sqrt((lh_x - px) ** 2 + (lh_y - py) ** 2 + (lh_z - pz) ** 2)
                if lh_gap_ms > LOSS_GAP_THRESHOLD_MS and dist < FREEZE_EPSILON_M:
                    lh_loss_flag = 1
                    lh_loss_reason = "timestamp_gap+frozen_estimate" if lh_loss_reason else "frozen_estimate"

        with self.data_lock:
            ot_pc_time_s = self.ot_latest_pc_time_s
            ot_frame_id = self.ot_latest_frame_id
            ot_rb_id = self.ot_latest_rb_id
            ot_x = self.ot_latest_x
            ot_y = self.ot_latest_y
            ot_z = self.ot_latest_z
            ot_tracked = self.ot_latest_tracked

        ot_age_ms = None
        ot_gap_ms = None
        ot_loss_flag = 1
        ot_loss_reason = "no_sample"
        path_error_3d_m = None

        if ot_pc_time_s is not None:
            ot_age_ms = int(max(0.0, (lh_pc_time_s - ot_pc_time_s) * 1000.0))
            ot_gap_ms = ot_age_ms

            if ot_tracked and ot_age_ms <= OT_FRESHNESS_LIMIT_MS:
                ot_loss_flag = 0
                ot_loss_reason = ""
            elif not ot_tracked:
                ot_loss_flag = 1
                ot_loss_reason = "untracked"
            else:
                ot_loss_flag = 1
                ot_loss_reason = "stale_sample"

        if ot_loss_flag:
            self.total_ot_loss_time_ms += EXPECTED_LOG_PERIOD_MS
            if self.ot_prev_loss_flag == 0:
                self.ot_loss_event_count += 1
        self.ot_prev_loss_flag = ot_loss_flag

        if (
            ot_loss_flag == 0
            and ot_x is not None and ot_y is not None and ot_z is not None
            and lh_loss_flag == 0
        ):
            path_error_3d_m = math.sqrt(
                (lh_x - ot_x) ** 2 + (lh_y - ot_y) ** 2 + (lh_z - ot_z) ** 2
            )

        row = [
            round(lh_pc_time_s, 6),
            int(timestamp),
            round(lh_x, 5),
            round(lh_y, 5),
            round(lh_z, 5),
            int(lh_gap_ms),
            int(lh_loss_flag),
            int(lh_estimated_lost_time_ms),
            lh_loss_reason,
            "" if ot_frame_id is None else ot_frame_id,
            "" if ot_rb_id is None else ot_rb_id,
            "" if ot_pc_time_s is None else round(ot_pc_time_s, 6),
            "" if ot_x is None else round(ot_x, 5),
            "" if ot_y is None else round(ot_y, 5),
            "" if ot_z is None else round(ot_z, 5),
            int(1 if ot_tracked else 0),
            "" if ot_gap_ms is None else int(ot_gap_ms),
            int(ot_loss_flag),
            ot_loss_reason,
            "" if path_error_3d_m is None else round(path_error_3d_m, 6),
        ]

        with self.data_lock:
            self.log_rows.append(row)

        self.prev_lh_timestamp = timestamp
        self.prev_lh_xyz = (lh_x, lh_y, lh_z)

    def stop_shape_logging(self):
        flight_time_s = 0.0
        if self.shape_start_time is not None:
            flight_time_s = time.time() - self.shape_start_time

        try:
            if self.shape_logconf is not None:
                try:
                    self.shape_logconf.stop()
                except Exception:
                    pass
                try:
                    self.cf.log.delete_config(self.shape_logconf)
                except Exception:
                    pass
        finally:
            self.shape_logconf = None
            self.logging_active = False

        if not self.log_rows:
            self._status("No shape data was logged.")
            self._reset_run_state()
            return

        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        os.makedirs(desktop, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        shape_name = self._sanitize_filename_part(self.current_shape_name or "shape")
        shape_size = self.current_shape_size if self.current_shape_size is not None else 0.0

        size_str = self._format_filename_number(shape_size)
        time_str = self._format_filename_number(flight_time_s)
        base_name = f"crazyflie_{shape_name}_size{size_str}m_time{time_str}s_{timestamp}"

        csv_path = os.path.join(desktop, base_name + ".csv")
        xlsx_path = os.path.join(desktop, base_name + ".xlsx")

        total_run_time_ms = flight_time_s * 1000.0

        lh_loss_percent = 0.0
        if total_run_time_ms > 0:
            lh_loss_percent = 100.0 * self.total_lh_loss_time_ms / total_run_time_ms

        ot_loss_percent = 0.0
        if total_run_time_ms > 0:
            ot_loss_percent = 100.0 * self.total_ot_loss_time_ms / total_run_time_ms

        valid_errors = []
        for row in self.log_rows:
            val = row[-1]
            if val != "":
                valid_errors.append(float(val))

        mean_err = sum(valid_errors) / len(valid_errors) if valid_errors else None
        rmse_err = math.sqrt(sum(v * v for v in valid_errors) / len(valid_errors)) if valid_errors else None
        max_err = max(valid_errors) if valid_errors else None

        columns = [
            "pc_time_s",
            "lh_timestamp_ms",
            "lh_x_m",
            "lh_y_m",
            "lh_z_m",
            "lh_gap_ms",
            "lh_loss_flag",
            "lh_estimated_lost_time_ms",
            "lh_loss_reason",
            "ot_frame_id",
            "ot_rb_id",
            "ot_pc_time_s",
            "ot_x_m",
            "ot_y_m",
            "ot_z_m",
            "ot_tracked_flag",
            "ot_gap_ms",
            "ot_loss_flag",
            "ot_loss_reason",
            "path_error_3d_m",
        ]

        with open(csv_path, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(columns)
            writer.writerows(self.log_rows)

        self._status(f"Saved CSV: {csv_path}")
        self._status(
            f"LH loss={lh_loss_percent:.2f}% ({self.lh_loss_event_count} events), "
            f"OT loss={ot_loss_percent:.2f}% ({self.ot_loss_event_count} events)"
        )

        if pd is not None:
            try:
                df = pd.DataFrame(self.log_rows, columns=columns)
                summary_df = pd.DataFrame(
                    [
                        ["shape", self.current_shape_name],
                        ["shape_size_m", shape_size],
                        ["flight_time_s", round(flight_time_s, 3)],
                        ["lh_loss_event_count", self.lh_loss_event_count],
                        ["lh_estimated_total_loss_time_ms", round(self.total_lh_loss_time_ms, 1)],
                        ["lh_estimated_loss_percent", round(lh_loss_percent, 3)],
                        ["ot_loss_event_count", self.ot_loss_event_count],
                        ["ot_estimated_total_loss_time_ms", round(self.total_ot_loss_time_ms, 1)],
                        ["ot_estimated_loss_percent", round(ot_loss_percent, 3)],
                        ["mean_3d_error_m", "" if mean_err is None else round(mean_err, 6)],
                        ["rmse_3d_error_m", "" if rmse_err is None else round(rmse_err, 6)],
                        ["max_3d_error_m", "" if max_err is None else round(max_err, 6)],
                        ["rigid_body_name", self.ot_body_name_var.get().strip()],
                        ["rigid_body_id_filter", self.ot_body_id_var.get().strip()],
                        ["saved_at", timestamp],
                    ],
                    columns=["field", "value"],
                )

                with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="FlightData")
                    summary_df.to_excel(writer, index=False, sheet_name="Summary")

                self._status(f"Saved Excel: {xlsx_path}")
            except Exception as e:
                self._status(f"CSV saved, but Excel save failed: {e}")
        else:
            self._status("CSV saved. Excel not saved because pandas/openpyxl is not installed.")

        self._reset_run_state()

    def _reset_run_state(self):
        self.current_shape_name = None
        self.current_shape_size = None
        self.shape_start_time = None

        self.prev_lh_timestamp = None
        self.prev_lh_xyz = None
        self.total_lh_loss_time_ms = 0.0
        self.lh_loss_event_count = 0

        self.ot_prev_loss_flag = 0
        self.total_ot_loss_time_ms = 0.0
        self.ot_loss_event_count = 0

    # -------------------------------------------------
    # Flight control
    # -------------------------------------------------

    def takeoff_clicked(self):
        self._run_threaded(self.takeoff)

    def land_clicked(self):
        self._run_threaded(self.land)

    def emergency_stop_clicked(self):
        self._run_threaded(self.emergency_stop)

    def goto_waypoint_clicked(self):
        self._run_threaded(self.goto_waypoint)

    def return_to_origin_clicked(self):
        self._run_threaded(self.return_to_origin)

    def shape_clicked(self, shape_name: str):
        self._run_threaded(lambda: self.fly_shape(shape_name))

    def takeoff(self):
        self.busy = True
        try:
            if not self.connected or not self.cf:
                self._status("Crazyflie not connected.")
                return

            if self.in_air:
                self._status("Already in the air.")
                return

            height = self._safe_float(self.height_var, "Takeoff Height", minimum=0.1)

            self.origin_z = height
            self.hover_x = self.origin_x
            self.hover_y = self.origin_y
            self.hover_z = height

            self.wp_z_var.set(f"{height:.2f}")

            self.pc = PositionHlCommander(
                self.scf,
                x=self.hover_x,
                y=self.hover_y,
                z=0.0,
                default_height=height,
                controller=PositionHlCommander.CONTROLLER_PID,
            )
            self.pc.__enter__()

            time.sleep(0.2)
            self.in_air = True
            self._status(f"Took off to {height:.2f} m.")
            time.sleep(1.0)

        except Exception as e:
            self._status(f"Takeoff failed: {e}")
            self.in_air = False
            try:
                if self.pc is not None:
                    self.pc.__exit__(None, None, None)
            except Exception:
                pass
            self.pc = None
        finally:
            self.busy = False

    def land(self):
        self.busy = True
        try:
            if not self.in_air:
                self._status("Drone is not in the air.")
                return

            if self.logging_active:
                self.stop_shape_logging()

            if self.pc is not None:
                self.pc.__exit__(None, None, None)
                self.pc = None

            time.sleep(1.0)
            self.in_air = False
            self._status("Landed.")

            try:
                self.cf.platform.send_arming_request(False)
            except Exception:
                pass

        except Exception as e:
            self._status(f"Land failed: {e}")
        finally:
            self.busy = False

    def emergency_stop(self):
        self.busy = True
        try:
            self._status("EMERGENCY STOP requested.")

            if self.logging_active:
                self.stop_shape_logging()

            if self.cf:
                try:
                    self.cf.high_level_commander.stop()
                except Exception:
                    pass
                try:
                    self.cf.commander.send_stop_setpoint()
                except Exception:
                    pass
                try:
                    self.cf.platform.send_arming_request(False)
                except Exception:
                    pass

            self.in_air = False
            self.pc = None
            self._status("Emergency stop command sent.")

        except Exception as e:
            self._status(f"Emergency stop failed: {e}")
        finally:
            self.busy = False

    def goto_waypoint(self):
        self.busy = True
        try:
            if not self.connected or not self.cf:
                self._status("Crazyflie not connected.")
                return

            if not self.in_air or self.pc is None:
                self._status("Take off first.")
                return

            x = self._safe_float(self.wp_x_var, "Waypoint X")
            y = self._safe_float(self.wp_y_var, "Waypoint Y")
            z = self._safe_float(self.wp_z_var, "Waypoint Z", minimum=0.05)
            speed = self._safe_float(self.speed_var, "Flight Speed", minimum=0.05)

            current = (self.hover_x, self.hover_y, self.hover_z)
            target = (x, y, z)
            distance = math.dist(current, target)
            duration = max(distance / speed, 0.8)

            self._status(f"Going to waypoint X={x:.2f}, Y={y:.2f}, Z={z:.2f}")
            self.pc.go_to(x, y, z, velocity=speed)
            time.sleep(duration + 0.3)

            self.hover_x = x
            self.hover_y = y
            self.hover_z = z

            self._status("Waypoint reached.")

        except Exception as e:
            self._status(f"Go to waypoint failed: {e}")
        finally:
            self.busy = False

    def return_to_origin(self):
        self.busy = True
        try:
            if not self.connected or not self.cf:
                self._status("Crazyflie not connected.")
                return

            if not self.in_air or self.pc is None:
                self._status("Take off first.")
                return

            speed = self._safe_float(self.speed_var, "Flight Speed", minimum=0.05)
            target = (self.origin_x, self.origin_y, self.origin_z)
            current = (self.hover_x, self.hover_y, self.hover_z)

            distance = math.dist(current, target)
            duration = max(distance / speed, 0.8)

            self._status(
                f"Returning to origin X={self.origin_x:.2f}, Y={self.origin_y:.2f}, Z={self.origin_z:.2f}"
            )
            self.pc.go_to(self.origin_x, self.origin_y, self.origin_z, velocity=speed)
            time.sleep(duration + 0.3)

            self.hover_x = self.origin_x
            self.hover_y = self.origin_y
            self.hover_z = self.origin_z

            self._status("Returned to origin.")

        except Exception as e:
            self._status(f"Return to origin failed: {e}")
        finally:
            self.busy = False

    # -------------------------------------------------
    # Shapes
    # -------------------------------------------------

    def fly_shape(self, shape_name: str):
        self.busy = True
        try:
            if not self.connected or not self.cf:
                self._status("Crazyflie not connected.")
                return

            if not self.ot_connected or not self.ot_streaming_ok or not self.ot_rigid_body_seen:
                self._status("OptiTrack is not ready. Press Test OptiTrack first.")
                self._update_ready_status()
                return

            if not self.in_air or self.pc is None:
                self._status("Take off first.")
                return

            size = self._safe_float(self.size_var, "Shape Size", minimum=0.05)
            speed = self._safe_float(self.speed_var, "Flight Speed", minimum=0.05)
            z = self.hover_z

            points = self.make_shape_points(shape_name, size, z)

            self._status(f"Flying {shape_name} with size={size:.2f} m, speed={speed:.2f} m/s")
            self.start_shape_logging(shape_name, size)

            time.sleep(0.2)
            prev = (self.hover_x, self.hover_y, z)

            for p in points:
                x, y, z = p
                distance = math.dist(prev, p)
                duration = max(distance / speed, 0.6)
                self.pc.go_to(x, y, z, velocity=speed)
                time.sleep(duration + 0.2)
                prev = p

            home = (self.hover_x, self.hover_y, self.hover_z)
            distance = math.dist(prev, home)
            duration = max(distance / speed, 0.6)
            self.pc.go_to(*home, velocity=speed)
            time.sleep(duration + 0.4)

            self.stop_shape_logging()
            self._status(f"{shape_name.capitalize()} complete. Hovering at current center.")

        except Exception as e:
            self._status(f"{shape_name.capitalize()} failed: {e}")
            try:
                if self.logging_active:
                    self.stop_shape_logging()
            except Exception:
                pass
        finally:
            self.busy = False

    def make_shape_points(self, shape_name: str, size: float, z: float):
        cx = self.hover_x
        cy = self.hover_y

        if shape_name == "square":
            half = size / 2.0
            return [
                (cx - half, cy - half, z),
                (cx + half, cy - half, z),
                (cx + half, cy + half, z),
                (cx - half, cy + half, z),
                (cx - half, cy - half, z),
            ]

        if shape_name == "triangle":
            h = (math.sqrt(3) / 2.0) * size
            return [
                (cx, cy + (2.0 / 3.0) * h, z),
                (cx - size / 2.0, cy - (1.0 / 3.0) * h, z),
                (cx + size / 2.0, cy - (1.0 / 3.0) * h, z),
                (cx, cy + (2.0 / 3.0) * h, z),
            ]

        if shape_name == "circle":
            radius = size / 2.0
            points = []
            num_points = 28
            for i in range(num_points + 1):
                angle = 2.0 * math.pi * i / num_points
                x = cx + radius * math.cos(angle)
                y = cy + radius * math.sin(angle)
                points.append((x, y, z))
            return points

        raise ValueError(f"Unknown shape: {shape_name}")

    # -------------------------------------------------
    # Shutdown
    # -------------------------------------------------

    def on_close(self):
        try:
            if self.logging_active:
                self.stop_shape_logging()
        except Exception:
            pass

        try:
            self.stop_live_logging()
        except Exception:
            pass

        try:
            self._stop_optitrack_client()
        except Exception:
            pass

        try:
            if self.in_air:
                if messagebox.askyesno("Drone in air", "The drone may still be in the air. Try to land before closing?"):
                    self.land()
        except Exception:
            pass

        try:
            if self.scf is not None:
                self.scf.close_link()
        except Exception:
            pass

        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = CrazyflieDualLoggerGUI(root)
    root.mainloop()