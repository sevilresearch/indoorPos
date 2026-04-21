"""
Crazyflie 2.1+ Lighthouse + OptiTrack Dual Logger GUI
-----------------------------------------------------
Features:
- Pre-flight navigation source selector:
    - Lighthouse
    - Motion Capture
- Connect / Disconnect Crazyflie
- Test / Reconnect / Stop OptiTrack
- Take off / Land / Emergency stop
- Go to waypoint
- Return to origin
- Fly square / circle / triangle
- Live XYZ display for Lighthouse and OptiTrack
- Independent timestamped logging for Lighthouse-side position and OptiTrack
- Aligned comparison CSV with error columns
- 3D comparison plot

Logging behavior:
- Lighthouse navigation mode:
    * Set lighthouse.method = 1
    * Fly using Lighthouse-driven estimator
    * Record stateEstimate.x/y/z as Lighthouse navigation estimate
    * Record OptiTrack independently with its own timestamps

- Motion Capture navigation mode:
    * Set lighthouse.method = 0
    * Fly using OptiTrack external position
    * Record raw lighthouse.x/y/z independently with its own timestamps
    * Record OptiTrack independently with its own timestamps

Saved files after each run:
- ..._lighthouse_raw.csv
- ..._optitrack_raw.csv
- ..._aligned_comparison.csv
- ..._summary.xlsx
- ..._3d_path.png

Install:
    pip install cflib pandas openpyxl matplotlib numpy
"""

import csv
import math
import os
import threading
import time
import traceback
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
    import numpy as np
except Exception:
    np = None

try:
    import matplotlib.pyplot as plt
except Exception:
    plt = None

try:
    from NatNetClient import NatNetClient
    NATNET_AVAILABLE = True
except Exception:
    NatNetClient = None
    NATNET_AVAILABLE = False


DEFAULT_URI = "radio://0/80/2M/E7E7E7E7C3"

NAV_LIGHTHOUSE = "Lighthouse"
NAV_MOCAP = "Motion Capture"

LIVE_PERIOD_MS = 100
RUN_PERIOD_MS = 10

OT_FRESHNESS_LIMIT_MS = 30
EXTPOS_SEND_PERIOD_S = 0.01

# Change these only if OptiTrack axes do not match your room axes
OT_AXIS_ORDER = ("x", "y", "z")
OT_AXIS_SIGN = (1.0, 1.0, 1.0)
OT_OFFSET_M = (0.0, 0.0, 0.0)


class CrazyflieDualLoggerGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Crazyflie Lighthouse + OptiTrack Dual Logger")
        self.root.geometry("1360x980")

        self.scf = None
        self.cf = None
        self.pc = None

        self.connected = False
        self.in_air = False
        self.busy = False

        # GUI/live log config
        self.live_lh_logconf = None

        # Run log config
        self.run_lh_logconf = None

        # Run state
        self.logging_active = False
        self.current_run_name = None
        self.current_run_size = None
        self.run_start_time = None

        # Separate sample buffers
        self.lh_samples = []
        self.ot_samples = []

        # Hover/origin state
        self.hover_x = 0.0
        self.hover_y = 0.0
        self.hover_z = 0.5
        self.origin_x = 0.0
        self.origin_y = 0.0
        self.origin_z = 0.6

        # Navigation source
        self.nav_source_selected = NAV_LIGHTHOUSE
        self.nav_source_in_use = NAV_LIGHTHOUSE

        # Live Lighthouse XYZ shown in GUI
        self.live_lh_x = None
        self.live_lh_y = None
        self.live_lh_z = None

        # OptiTrack state
        self.ot_client = None
        self.ot_connected = False
        self.ot_streaming_ok = False
        self.ot_rigid_body_seen = False

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

        # Motion capture external-position feed
        self.extpos_thread = None
        self.extpos_stop_event = threading.Event()
        self.extpos_running = False

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

        ttk.Label(settings_frame, text="Navigation Source").grid(row=0, column=2, sticky="w", **pad)
        self.nav_source_var = tk.StringVar(value=NAV_LIGHTHOUSE)
        self.nav_source_combo = ttk.Combobox(
            settings_frame,
            textvariable=self.nav_source_var,
            values=[NAV_LIGHTHOUSE, NAV_MOCAP],
            state="readonly",
            width=18,
        )
        self.nav_source_combo.grid(row=0, column=3, sticky="w", **pad)
        self.nav_source_combo.bind("<<ComboboxSelected>>", self._on_nav_source_changed)

        self.nav_status_var = tk.StringVar(value="Active nav source: Lighthouse")
        ttk.Label(settings_frame, textvariable=self.nav_status_var).grid(row=1, column=2, columnspan=2, sticky="w", **pad)

        ttk.Label(settings_frame, text="Rigid Body Name").grid(row=2, column=2, sticky="w", **pad)
        self.ot_body_name_var = tk.StringVar(value="Robot_3")
        ttk.Entry(settings_frame, textvariable=self.ot_body_name_var, width=18).grid(row=2, column=3, sticky="w", **pad)

        ttk.Label(settings_frame, text="Rigid Body ID (optional)").grid(row=3, column=2, sticky="w", **pad)
        self.ot_body_id_var = tk.StringVar(value="")
        ttk.Entry(settings_frame, textvariable=self.ot_body_id_var, width=18).grid(row=3, column=3, sticky="w", **pad)

        natnet_msg = "NatNetClient.py found" if NATNET_AVAILABLE else "NatNetClient.py NOT found"
        self.natnet_status_var = tk.StringVar(value=natnet_msg)
        ttk.Label(settings_frame, textvariable=self.natnet_status_var).grid(row=4, column=0, columnspan=2, sticky="w", **pad)

        self.ot_status_var = tk.StringVar(value="OptiTrack: not connected")
        ttk.Label(settings_frame, textvariable=self.ot_status_var).grid(row=4, column=2, columnspan=2, sticky="w", **pad)

        self.record_ready_var = tk.StringVar(value="Record Ready: No")
        ttk.Label(settings_frame, textvariable=self.record_ready_var).grid(row=5, column=2, columnspan=2, sticky="w", **pad)

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

        self.live_lh_x_var = tk.StringVar(value="nan")
        self.live_lh_y_var = tk.StringVar(value="nan")
        self.live_lh_z_var = tk.StringVar(value="nan")

        self.live_ot_x_var = tk.StringVar(value="nan")
        self.live_ot_y_var = tk.StringVar(value="nan")
        self.live_ot_z_var = tk.StringVar(value="nan")

        ttk.Label(live_frame, text="LH X").grid(row=0, column=0, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_lh_x_var, width=10).grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(live_frame, text="LH Y").grid(row=0, column=2, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_lh_y_var, width=10).grid(row=0, column=3, sticky="w", **pad)
        ttk.Label(live_frame, text="LH Z").grid(row=0, column=4, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_lh_z_var, width=10).grid(row=0, column=5, sticky="w", **pad)

        ttk.Label(live_frame, text="OT X").grid(row=1, column=0, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_ot_x_var, width=10).grid(row=1, column=1, sticky="w", **pad)
        ttk.Label(live_frame, text="OT Y").grid(row=1, column=2, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_ot_y_var, width=10).grid(row=1, column=3, sticky="w", **pad)
        ttk.Label(live_frame, text="OT Z").grid(row=1, column=4, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_ot_z_var, width=10).grid(row=1, column=5, sticky="w", **pad)

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

    # -------------------------------------------------
    # Button click wrappers
    # -------------------------------------------------

    def connect_clicked(self):
        self._status("Connect button pressed.")
        self._run_threaded(self.connect_cf)

    def disconnect_clicked(self):
        self._status("Disconnect button pressed.")
        self._run_threaded(self.disconnect_cf)

    def takeoff_clicked(self):
        self._status("Take Off button pressed.")
        self._run_threaded(self.takeoff)

    def land_clicked(self):
        self._status("Land button pressed.")
        self._run_threaded(self.land)

    def emergency_stop_clicked(self):
        self._status("Emergency Stop button pressed.")
        self._run_threaded(self.emergency_stop)

    def goto_waypoint_clicked(self):
        self._status("Go To Waypoint button pressed.")
        self._run_threaded(self.goto_waypoint)

    def return_to_origin_clicked(self):
        self._status("Return To Origin button pressed.")
        self._run_threaded(self.return_to_origin)

    def test_optitrack_clicked(self):
        self._status("Test OptiTrack button pressed.")
        self._run_threaded(self.test_optitrack)

    def reconnect_optitrack_clicked(self):
        self._status("Reconnect OptiTrack button pressed.")
        self._run_threaded(self.reconnect_optitrack)

    def stop_optitrack_clicked(self):
        self._status("Stop OptiTrack button pressed.")
        self._run_threaded(self.stop_optitrack_manual)

    def shape_clicked(self, shape_name: str):
        self._status(f"{shape_name.capitalize()} button pressed.")
        self._run_threaded(lambda: self.fly_shape(shape_name))

    # -------------------------------------------------
    # Thread helper
    # -------------------------------------------------

    def _run_threaded(self, target):
        if self.busy:
            self._status("Busy. Wait for the current action to finish.")
            return

        def worker():
            try:
                target()
            except Exception as e:
                tb = traceback.format_exc()
                print(tb)
                self._status(f"ERROR: {e}")

        threading.Thread(target=worker, daemon=True).start()

    # -------------------------------------------------
    # Utility
    # -------------------------------------------------

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

    def _format_filename_number(self, value: float) -> str:
        return f"{value:.2f}".replace(".", "p")

    def _sanitize_filename_part(self, text: str) -> str:
        return "".join(c if c.isalnum() or c in ("_", "-") else "_" for c in text)

    def _selected_nav_source(self):
        return self.nav_source_var.get().strip() or NAV_LIGHTHOUSE

    def _update_nav_status(self):
        selected = self._selected_nav_source()
        if self.in_air:
            self.nav_status_var.set(
                f"Active nav source: {self.nav_source_in_use} | Next takeoff: {selected}"
            )
        else:
            self.nav_status_var.set(f"Active nav source: {selected}")

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
            self.ot_status_var.set("OptiTrack: ready")

        self._update_nav_status()

    def _on_nav_source_changed(self, event=None):
        self.nav_source_selected = self._selected_nav_source()
        if self.in_air:
            self._status(
                f"Navigation source changed to '{self.nav_source_selected}'. It will apply on the next takeoff."
            )
        else:
            self._status(f"Navigation source set to '{self.nav_source_selected}' for the next takeoff.")
            if self.connected:
                self.restart_live_logging_for_selected_mode()
        self._update_nav_status()
        self._update_ready_status()

    # -------------------------------------------------
    # Live logging
    # -------------------------------------------------

    def restart_live_logging_for_selected_mode(self):
        try:
            self.stop_live_logging()
        except Exception:
            pass
        self.start_live_logging()

    def start_live_logging(self):
        if not self.cf:
            return

        selected_nav = self._selected_nav_source()

        try:
            if selected_nav == NAV_LIGHTHOUSE:
                self.live_lh_logconf = LogConfig(name="LiveLHDisplay", period_in_ms=LIVE_PERIOD_MS)
                self.live_lh_logconf.add_variable("stateEstimate.x", "float")
                self.live_lh_logconf.add_variable("stateEstimate.y", "float")
                self.live_lh_logconf.add_variable("stateEstimate.z", "float")
                self.live_lh_logconf.data_received_cb.add_callback(self._live_lh_est_callback)
            else:
                self.live_lh_logconf = LogConfig(name="LiveLHDisplay", period_in_ms=LIVE_PERIOD_MS)
                self.live_lh_logconf.add_variable("lighthouse.x", "float")
                self.live_lh_logconf.add_variable("lighthouse.y", "float")
                self.live_lh_logconf.add_variable("lighthouse.z", "float")
                self.live_lh_logconf.add_variable("lighthouse.status", "uint8_t")
                self.live_lh_logconf.add_variable("lighthouse.bsActive", "uint16_t")
                self.live_lh_logconf.data_received_cb.add_callback(self._live_lh_raw_callback)

            self.cf.log.add_config(self.live_lh_logconf)
            self.live_lh_logconf.start()
            self._status("Live Lighthouse display started.")
        except Exception as e:
            self.live_lh_logconf = None
            self._status(f"Could not start live Lighthouse display: {e}")

    def stop_live_logging(self):
        try:
            if self.live_lh_logconf is not None:
                try:
                    self.live_lh_logconf.stop()
                except Exception:
                    pass
                try:
                    self.cf.log.delete_config(self.live_lh_logconf)
                except Exception:
                    pass
        finally:
            self.live_lh_logconf = None

    def _live_lh_est_callback(self, timestamp, data, logconf):
        with self.data_lock:
            self.live_lh_x = float(data["stateEstimate.x"])
            self.live_lh_y = float(data["stateEstimate.y"])
            self.live_lh_z = float(data["stateEstimate.z"])
        self.root.after(0, self._update_live_labels)

    def _live_lh_raw_callback(self, timestamp, data, logconf):
        with self.data_lock:
            self.live_lh_x = float(data["lighthouse.x"])
            self.live_lh_y = float(data["lighthouse.y"])
            self.live_lh_z = float(data["lighthouse.z"])
        self.root.after(0, self._update_live_labels)

    def _update_live_labels(self):
        with self.data_lock:
            self.live_lh_x_var.set("nan" if self.live_lh_x is None else f"{self.live_lh_x:.3f}")
            self.live_lh_y_var.set("nan" if self.live_lh_y is None else f"{self.live_lh_y:.3f}")
            self.live_lh_z_var.set("nan" if self.live_lh_z is None else f"{self.live_lh_z:.3f}")

            self.live_ot_x_var.set("nan" if self.live_ot_x is None else f"{self.live_ot_x:.3f}")
            self.live_ot_y_var.set("nan" if self.live_ot_y is None else f"{self.live_ot_y:.3f}")
            self.live_ot_z_var.set("nan" if self.live_ot_z is None else f"{self.live_ot_z:.3f}")

    # -------------------------------------------------
    # OptiTrack
    # -------------------------------------------------

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

        self._status("OptiTrack client started.")
        self.root.after(0, self._update_live_labels)
        self.root.after(0, self._update_ready_status)

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
            self.root.after(0, self._update_live_labels)

    def _ot_new_frame_callback(self, frame_number):
        with self.data_lock:
            self.ot_streaming_ok = True
            self.ot_latest_frame_id = frame_number
        self.root.after(0, self._update_ready_status)

    def _ot_rigid_body_callback(self, rigid_body_id, position, rotation):
        try:
            target_id = self._safe_optional_int(self.ot_body_id_var)
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

                if self.logging_active:
                    self.ot_samples.append({
                        "pc_time_s": now_pc_time_s,
                        "rb_id": rigid_body_id,
                        "x_m": x_cf,
                        "y_m": y_cf,
                        "z_m": z_cf,
                        "tracked_flag": 1,
                    })

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

    # -------------------------------------------------
    # External position feed for mocap navigation
    # -------------------------------------------------

    def start_extpos_feed(self):
        if self.extpos_running:
            return
        self.extpos_stop_event.clear()
        self.extpos_thread = threading.Thread(target=self._extpos_worker, daemon=True)
        self.extpos_thread.start()
        self.extpos_running = True
        self._status("Motion capture external-position feed started.")

    def stop_extpos_feed(self):
        if not self.extpos_running:
            return
        self.extpos_stop_event.set()
        if self.extpos_thread is not None and self.extpos_thread.is_alive():
            self.extpos_thread.join(timeout=1.0)
        self.extpos_thread = None
        self.extpos_running = False
        self._status("Motion capture external-position feed stopped.")

    def _extpos_worker(self):
        while not self.extpos_stop_event.is_set():
            try:
                if self.connected and self.cf is not None and self.nav_source_in_use == NAV_MOCAP:
                    with self.data_lock:
                        ot_pc_time_s = self.ot_latest_pc_time_s
                        ot_x = self.ot_latest_x
                        ot_y = self.ot_latest_y
                        ot_z = self.ot_latest_z
                        ot_tracked = self.ot_latest_tracked

                    if ot_pc_time_s is not None and ot_tracked:
                        age_ms = (time.perf_counter() - ot_pc_time_s) * 1000.0
                        if age_ms <= OT_FRESHNESS_LIMIT_MS and ot_x is not None:
                            self.cf.extpos.send_extpos(ot_x, ot_y, ot_z)
            except Exception:
                pass
            time.sleep(EXTPOS_SEND_PERIOD_S)

    def _wait_for_fresh_optitrack(self, timeout_s=2.0):
        t0 = time.time()
        while time.time() - t0 < timeout_s:
            with self.data_lock:
                ot_pc_time_s = self.ot_latest_pc_time_s
                ot_tracked = self.ot_latest_tracked
            if ot_pc_time_s is not None and ot_tracked:
                age_ms = (time.perf_counter() - ot_pc_time_s) * 1000.0
                if age_ms <= OT_FRESHNESS_LIMIT_MS:
                    return True
            time.sleep(0.02)
        return False

    # -------------------------------------------------
    # Connection / setup
    # -------------------------------------------------

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

            try:
                initial_nav = self._selected_nav_source()
                if initial_nav == NAV_LIGHTHOUSE:
                    self.cf.param.set_value("lighthouse.method", "1")
                    self._status("Set lighthouse.method = 1 for Lighthouse navigation.")
                else:
                    self.cf.param.set_value("lighthouse.method", "0")
                    self._status("Set lighthouse.method = 0 for raw Lighthouse collection.")
                time.sleep(0.2)
            except Exception as e:
                self._status(f"Could not set lighthouse.method: {e}")

            self.reset_estimator()
            self.start_live_logging()

            self.connected = True
            self._status("Crazyflie connected successfully.")
            self._update_ready_status()

        finally:
            self.busy = False

    def disconnect_cf(self):
        self.busy = True
        try:
            if self.logging_active:
                self.stop_run_logging()

            self.stop_live_logging()
            self.stop_extpos_feed()

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
    # Run logging
    # -------------------------------------------------

    def start_run_logging(self, run_name: str, run_size: float):
        if not self.cf:
            return

        self.current_run_name = run_name
        self.current_run_size = run_size
        self.run_start_time = time.time()

        self.lh_samples = []
        self.ot_samples = []

        try:
            if self.nav_source_in_use == NAV_LIGHTHOUSE:
                self.run_lh_logconf = LogConfig(name="RunLH", period_in_ms=RUN_PERIOD_MS)
                self.run_lh_logconf.add_variable("stateEstimate.x", "float")
                self.run_lh_logconf.add_variable("stateEstimate.y", "float")
                self.run_lh_logconf.add_variable("stateEstimate.z", "float")
                self.run_lh_logconf.data_received_cb.add_callback(self._run_lh_est_callback)
            else:
                self.run_lh_logconf = LogConfig(name="RunLH", period_in_ms=RUN_PERIOD_MS)
                self.run_lh_logconf.add_variable("lighthouse.x", "float")
                self.run_lh_logconf.add_variable("lighthouse.y", "float")
                self.run_lh_logconf.add_variable("lighthouse.z", "float")
                self.run_lh_logconf.add_variable("lighthouse.status", "uint8_t")
                self.run_lh_logconf.add_variable("lighthouse.bsActive", "uint16_t")
                self.run_lh_logconf.data_received_cb.add_callback(self._run_lh_raw_callback)

            self.cf.log.add_config(self.run_lh_logconf)
            self.run_lh_logconf.start()

            self.logging_active = True
            self._status(
                f"Started independent logging for {run_name} using {self.nav_source_in_use} navigation."
            )
        except Exception as e:
            self.run_lh_logconf = None
            self.logging_active = False
            self._status(f"Could not start run logger: {e}")

    def _run_lh_est_callback(self, timestamp, data, logconf):
        if not self.logging_active:
            return
        now_s = time.perf_counter()
        with self.data_lock:
            self.lh_samples.append({
                "pc_time_s": now_s,
                "source_label": "lh_nav_est",
                "x_m": float(data["stateEstimate.x"]),
                "y_m": float(data["stateEstimate.y"]),
                "z_m": float(data["stateEstimate.z"]),
                "status": "",
                "bs_active": "",
            })

    def _run_lh_raw_callback(self, timestamp, data, logconf):
        if not self.logging_active:
            return
        now_s = time.perf_counter()
        with self.data_lock:
            self.lh_samples.append({
                "pc_time_s": now_s,
                "source_label": "lh_raw",
                "x_m": float(data["lighthouse.x"]),
                "y_m": float(data["lighthouse.y"]),
                "z_m": float(data["lighthouse.z"]),
                "status": int(data["lighthouse.status"]),
                "bs_active": int(data["lighthouse.bsActive"]),
            })

    def _interp_series(self, target_t, src_t, src_v):
        if np is None:
            return None
        if len(src_t) < 2 or len(src_v) < 2:
            return None
        return np.interp(target_t, src_t, src_v)

    def _build_aligned_dataframe(self):
        if np is None or pd is None:
            return None

        if len(self.lh_samples) < 2 or len(self.ot_samples) < 2:
            return None

        lh_df = pd.DataFrame(self.lh_samples).sort_values("pc_time_s").reset_index(drop=True)
        ot_df = pd.DataFrame(self.ot_samples).sort_values("pc_time_s").reset_index(drop=True)

        lh_t = lh_df["pc_time_s"].to_numpy(dtype=float)
        ot_t = ot_df["pc_time_s"].to_numpy(dtype=float)

        t_start = max(lh_t[0], ot_t[0])
        t_end = min(lh_t[-1], ot_t[-1])

        if t_end <= t_start:
            return None

        dt = 0.01
        target_t = np.arange(t_start, t_end + 1e-9, dt)
        if len(target_t) < 2:
            return None

        lh_x = self._interp_series(target_t, lh_t, lh_df["x_m"].to_numpy(dtype=float))
        lh_y = self._interp_series(target_t, lh_t, lh_df["y_m"].to_numpy(dtype=float))
        lh_z = self._interp_series(target_t, lh_t, lh_df["z_m"].to_numpy(dtype=float))

        ot_x = self._interp_series(target_t, ot_t, ot_df["x_m"].to_numpy(dtype=float))
        ot_y = self._interp_series(target_t, ot_t, ot_df["y_m"].to_numpy(dtype=float))
        ot_z = self._interp_series(target_t, ot_t, ot_df["z_m"].to_numpy(dtype=float))

        if any(v is None for v in [lh_x, lh_y, lh_z, ot_x, ot_y, ot_z]):
            return None

        err_x = lh_x - ot_x
        err_y = lh_y - ot_y
        err_z = lh_z - ot_z
        err_3d = np.sqrt(err_x**2 + err_y**2 + err_z**2)

        aligned_df = pd.DataFrame({
            "pc_time_s": target_t,
            "lh_x_m": lh_x,
            "lh_y_m": lh_y,
            "lh_z_m": lh_z,
            "ot_x_m": ot_x,
            "ot_y_m": ot_y,
            "ot_z_m": ot_z,
            "error_x_m": err_x,
            "error_y_m": err_y,
            "error_z_m": err_z,
            "error_3d_m": err_3d,
        })
        return aligned_df

    def _save_3d_flight_path_plot(self, png_path: str, title: str):
        if plt is None or not self.lh_samples or not self.ot_samples:
            return

        lh_x = [float(s["x_m"]) for s in self.lh_samples if s["x_m"] != ""]
        lh_y = [float(s["y_m"]) for s in self.lh_samples if s["y_m"] != ""]
        lh_z = [float(s["z_m"]) for s in self.lh_samples if s["z_m"] != ""]

        ot_x = [float(s["x_m"]) for s in self.ot_samples if s["x_m"] != ""]
        ot_y = [float(s["y_m"]) for s in self.ot_samples if s["y_m"] != ""]
        ot_z = [float(s["z_m"]) for s in self.ot_samples if s["z_m"] != ""]

        if not lh_x or not ot_x:
            return

        fig = plt.figure(figsize=(9, 7))
        ax = fig.add_subplot(111, projection="3d")
        ax.plot(lh_x, lh_y, lh_z, label="Lighthouse")
        ax.plot(ot_x, ot_y, ot_z, label="OptiTrack")
        ax.set_title(title)
        ax.set_xlabel("X (m)")
        ax.set_ylabel("Y (m)")
        ax.set_zlabel("Z (m)")
        ax.legend()
        plt.tight_layout()
        plt.savefig(png_path, dpi=200)
        plt.close(fig)

    def stop_run_logging(self):
        run_time_s = 0.0
        if self.run_start_time is not None:
            run_time_s = time.time() - self.run_start_time

        try:
            if self.run_lh_logconf is not None:
                try:
                    self.run_lh_logconf.stop()
                except Exception:
                    pass
                try:
                    self.cf.log.delete_config(self.run_lh_logconf)
                except Exception:
                    pass
        finally:
            self.run_lh_logconf = None
            self.logging_active = False

        if not self.lh_samples or not self.ot_samples:
            self._status("Run stopped, but one of the data streams is missing.")
            return

        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        os.makedirs(desktop, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        run_name = self._sanitize_filename_part(self.current_run_name or "run")
        run_size = self.current_run_size if self.current_run_size is not None else 0.0
        size_str = self._format_filename_number(run_size)
        time_str = self._format_filename_number(run_time_s)
        nav_str = "OptiTrack" if self.nav_source_in_use == NAV_MOCAP else "Lighthouse"

        base_name = f"crazyflie_{run_name}_using_{nav_str}_size{size_str}m_time{time_str}s_{timestamp}"

        lh_csv = os.path.join(desktop, base_name + "_lighthouse_raw.csv")
        ot_csv = os.path.join(desktop, base_name + "_optitrack_raw.csv")
        aligned_csv = os.path.join(desktop, base_name + "_aligned_comparison.csv")
        xlsx_path = os.path.join(desktop, base_name + "_summary.xlsx")
        png_path = os.path.join(desktop, base_name + "_3d_path.png")

        lh_columns = ["pc_time_s", "source_label", "x_m", "y_m", "z_m", "status", "bs_active"]
        with open(lh_csv, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=lh_columns)
            writer.writeheader()
            writer.writerows(self.lh_samples)

        ot_columns = ["pc_time_s", "rb_id", "x_m", "y_m", "z_m", "tracked_flag"]
        with open(ot_csv, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=ot_columns)
            writer.writeheader()
            writer.writerows(self.ot_samples)

        self._status(f"Saved Lighthouse raw CSV: {lh_csv}")
        self._status(f"Saved OptiTrack raw CSV: {ot_csv}")

        aligned_df = self._build_aligned_dataframe()
        mean_3d = rmse_3d = max_3d = None

        if aligned_df is not None:
            aligned_df.to_csv(aligned_csv, index=False)
            self._status(f"Saved aligned comparison CSV: {aligned_csv}")

            if len(aligned_df) > 0:
                mean_3d = float(aligned_df["error_3d_m"].mean())
                rmse_3d = float(np.sqrt((aligned_df["error_3d_m"] ** 2).mean()))
                max_3d = float(aligned_df["error_3d_m"].max())
        else:
            self._status("Could not create aligned comparison CSV.")

        if pd is not None:
            try:
                lh_df = pd.DataFrame(self.lh_samples)
                ot_df = pd.DataFrame(self.ot_samples)

                summary_rows = [
                    ["navigation_source_used", self.nav_source_in_use],
                    ["run_name", self.current_run_name],
                    ["shape_size_m", run_size],
                    ["run_time_s", round(run_time_s, 3)],
                    ["lh_sample_count", len(self.lh_samples)],
                    ["ot_sample_count", len(self.ot_samples)],
                    ["mean_3d_error_m", "" if mean_3d is None else round(mean_3d, 6)],
                    ["rmse_3d_error_m", "" if rmse_3d is None else round(rmse_3d, 6)],
                    ["max_3d_error_m", "" if max_3d is None else round(max_3d, 6)],
                    ["rigid_body_name", self.ot_body_name_var.get().strip()],
                    ["rigid_body_id_filter", self.ot_body_id_var.get().strip()],
                    ["saved_at", timestamp],
                ]
                summary_df = pd.DataFrame(summary_rows, columns=["field", "value"])

                with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
                    lh_df.to_excel(writer, index=False, sheet_name="LighthouseRaw")
                    ot_df.to_excel(writer, index=False, sheet_name="OptiTrackRaw")
                    if aligned_df is not None:
                        aligned_df.to_excel(writer, index=False, sheet_name="AlignedComparison")
                    summary_df.to_excel(writer, index=False, sheet_name="Summary")

                self._status(f"Saved summary Excel: {xlsx_path}")
            except Exception as e:
                self._status(f"Could not save Excel summary: {e}")

        try:
            self._save_3d_flight_path_plot(
                png_path,
                f"{self.current_run_name} - {self.nav_source_in_use} Navigation"
            )
            if os.path.isfile(png_path):
                self._status(f"Saved 3D flight path plot: {png_path}")
        except Exception as e:
            self._status(f"Could not save 3D flight path plot: {e}")

        if mean_3d is not None:
            self._status(
                f"Aligned errors: mean={mean_3d:.4f} m, rmse={rmse_3d:.4f} m, max={max_3d:.4f} m"
            )

    # -------------------------------------------------
    # Flight control
    # -------------------------------------------------

    def takeoff(self):
        self.busy = True
        try:
            if not self.connected or not self.cf:
                self._status("Crazyflie not connected.")
                return

            if self.in_air:
                self._status("Already in the air.")
                return

            selected_nav = self._selected_nav_source()
            self.nav_source_in_use = selected_nav
            self._update_nav_status()

            try:
                if selected_nav == NAV_LIGHTHOUSE:
                    self.cf.param.set_value("lighthouse.method", "1")
                    self._status("Set lighthouse.method = 1 for Lighthouse navigation.")
                else:
                    self.cf.param.set_value("lighthouse.method", "0")
                    self._status("Set lighthouse.method = 0 so raw Lighthouse data can be collected.")
                time.sleep(0.2)
            except Exception as e:
                self._status(f"Could not set lighthouse.method: {e}")

            self.restart_live_logging_for_selected_mode()

            if selected_nav == NAV_MOCAP:
                if not self.ot_connected or not self.ot_streaming_ok or not self.ot_rigid_body_seen:
                    self._status("Motion Capture navigation selected, but OptiTrack is not ready.")
                    self._update_ready_status()
                    return

                self.start_extpos_feed()
                if not self._wait_for_fresh_optitrack(timeout_s=2.0):
                    self._status("No fresh OptiTrack sample available for Motion Capture navigation.")
                    return
            else:
                self.stop_extpos_feed()

            self.reset_estimator()

            if selected_nav == NAV_MOCAP:
                if not self._wait_for_fresh_optitrack(timeout_s=1.0):
                    self._status("OptiTrack sample became stale after estimator reset.")
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
            self._status(f"Took off to {height:.2f} m using {selected_nav} navigation.")
            time.sleep(1.0)
            self._update_nav_status()

        finally:
            self.busy = False

    def land(self):
        self.busy = True
        try:
            if not self.in_air:
                self._status("Drone is not in the air.")
                return

            if self.logging_active:
                self.stop_run_logging()

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

            self.stop_extpos_feed()
            self._update_nav_status()

        finally:
            self.busy = False

    def emergency_stop(self):
        self.busy = True
        try:
            self._status("EMERGENCY STOP requested.")

            if self.logging_active:
                self.stop_run_logging()

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
            self.stop_extpos_feed()
            self._status("Emergency stop command sent.")
            self._update_nav_status()

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

            if self.nav_source_in_use == NAV_MOCAP and not self._wait_for_fresh_optitrack(timeout_s=0.5):
                self._status("Motion Capture navigation is active, but OptiTrack is stale.")
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

            if self.nav_source_in_use == NAV_MOCAP and not self._wait_for_fresh_optitrack(timeout_s=0.5):
                self._status("Motion Capture navigation is active, but OptiTrack is stale.")
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

        finally:
            self.busy = False

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

            if self.nav_source_in_use == NAV_MOCAP and not self._wait_for_fresh_optitrack(timeout_s=0.5):
                self._status("Motion Capture navigation is active, but OptiTrack is stale.")
                return

            size = self._safe_float(self.size_var, "Shape Size", minimum=0.05)
            speed = self._safe_float(self.speed_var, "Flight Speed", minimum=0.05)
            z = self.hover_z

            points = self.make_shape_points(shape_name, size, z)

            self._status(
                f"Flying {shape_name} with size={size:.2f} m, speed={speed:.2f} m/s "
                f"using {self.nav_source_in_use} navigation"
            )
            self.start_run_logging(shape_name, size)

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

            self.stop_run_logging()
            self._status(f"{shape_name.capitalize()} complete. Hovering at current center.")

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
                self.stop_run_logging()
        except Exception:
            pass

        try:
            self.stop_live_logging()
        except Exception:
            pass

        try:
            self.stop_extpos_feed()
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