"""
Crazyflie 2.1+ Lighthouse GUI Flight Controller + Shape Logger
--------------------------------------------------------------
Features:
- Connect / Disconnect
- Take off to specific height
- Land
- Fly Square / Circle / Triangle
- Set shape size in meters
- Set flight speed in m/s
- Go to an X, Y, Z waypoint from the GUI
- Return to origin
- Live X, Y, Z readout in GUI
- Log ONLY while flying a shape
- Save BOTH CSV and Excel after each shape run
- File name includes:
    - shape
    - size
    - flight time
    - timestamp
- Loss-of-track / data-gap detection added:
    - gap_ms
    - loss_of_track
    - estimated_lost_time_ms
    - loss_reason
    - summary percent of run with loss

IMPORTANT:
1) Configure Lighthouse in cfclient first and write the Lighthouse config to the Crazyflie.
2) Disconnect cfclient before running this script.
3) Start with small shapes and low speed.
4) Fly safely.
5) Excel output requires:
       pip install pandas openpyxl

Recommended Python:
- Python 3.10

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


DEFAULT_URI = "radio://0/80/2M/E7E7E7E7C3"

# Logging / loss detection settings
LIVE_LOG_PERIOD_MS = 100          # live GUI display = 10 Hz
SHAPE_LOG_PERIOD_MS = 10          # shape logging = 100 Hz
EXPECTED_LOG_PERIOD_MS = 10       # expected gap at 100 Hz
LOSS_GAP_THRESHOLD_MS = 30        # if gap is bigger than this, flag it
FREEZE_EPSILON_M = 0.0005         # nearly no motion across a large gap


class CrazyflieShapeGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Crazyflie Lighthouse Shape Flight GUI")
        self.root.geometry("940x720")

        self.scf = None
        self.cf = None
        self.pc = None

        self.connected = False
        self.in_air = False
        self.busy = False

        # Logging configs
        self.shape_logconf = None
        self.live_logconf = None

        # Shape logging state
        self.logging_active = False
        self.log_rows = []
        self.current_shape_name = None
        self.current_shape_size = None
        self.shape_start_time = None

        # Loss tracking
        self.prev_shape_timestamp = None
        self.prev_logged_xyz = None
        self.total_loss_time_ms = 0.0
        self.loss_event_count = 0

        # Current commanded hover/home point
        self.hover_x = 0.0
        self.hover_y = 0.0
        self.hover_z = 0.5

        # Origin
        self.origin_x = 0.0
        self.origin_y = 0.0
        self.origin_z = 0.6

        # Live estimated position
        self.live_x = 0.0
        self.live_y = 0.0
        self.live_z = 0.0

        self.data_lock = threading.Lock()

        self._build_gui()
        cflib.crtp.init_drivers(enable_debug_driver=False)

    # ---------------- GUI ----------------

    def _build_gui(self):
        pad = {"padx": 8, "pady": 6}

        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True, padx=12, pady=12)

        # Flight settings
        flight_frame = ttk.LabelFrame(main, text="Flight Settings")
        flight_frame.pack(fill="x", **pad)

        ttk.Label(flight_frame, text="URI").grid(row=0, column=0, sticky="w", **pad)
        self.uri_var = tk.StringVar(value=DEFAULT_URI)
        ttk.Entry(flight_frame, textvariable=self.uri_var, width=35).grid(row=0, column=1, sticky="w", **pad)

        ttk.Label(flight_frame, text="Takeoff Height (m)").grid(row=1, column=0, sticky="w", **pad)
        self.height_var = tk.StringVar(value="0.6")
        ttk.Entry(flight_frame, textvariable=self.height_var, width=12).grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(flight_frame, text="Shape Size (m)").grid(row=2, column=0, sticky="w", **pad)
        self.size_var = tk.StringVar(value="0.6")
        ttk.Entry(flight_frame, textvariable=self.size_var, width=12).grid(row=2, column=1, sticky="w", **pad)

        ttk.Label(flight_frame, text="Flight Speed (m/s)").grid(row=3, column=0, sticky="w", **pad)
        self.speed_var = tk.StringVar(value="0.3")
        ttk.Entry(flight_frame, textvariable=self.speed_var, width=12).grid(row=3, column=1, sticky="w", **pad)

        # Waypoint
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

        ttk.Button(waypoint_frame, text="Go To Waypoint", command=self.goto_waypoint_clicked).grid(
            row=0, column=6, **pad
        )
        ttk.Button(waypoint_frame, text="Return To Origin", command=self.return_to_origin_clicked).grid(
            row=0, column=7, **pad
        )

        # Live position display
        live_frame = ttk.LabelFrame(main, text="Live Estimated Position")
        live_frame.pack(fill="x", **pad)

        self.live_x_var = tk.StringVar(value="0.000")
        self.live_y_var = tk.StringVar(value="0.000")
        self.live_z_var = tk.StringVar(value="0.000")

        ttk.Label(live_frame, text="X (m):").grid(row=0, column=0, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_x_var, width=12).grid(row=0, column=1, sticky="w", **pad)

        ttk.Label(live_frame, text="Y (m):").grid(row=0, column=2, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_y_var, width=12).grid(row=0, column=3, sticky="w", **pad)

        ttk.Label(live_frame, text="Z (m):").grid(row=0, column=4, sticky="w", **pad)
        ttk.Label(live_frame, textvariable=self.live_z_var, width=12).grid(row=0, column=5, sticky="w", **pad)

        # Controls
        button_frame = ttk.LabelFrame(main, text="Controls")
        button_frame.pack(fill="x", **pad)

        ttk.Button(button_frame, text="Connect", command=self.connect_clicked).grid(row=0, column=0, **pad)
        ttk.Button(button_frame, text="Disconnect", command=self.disconnect_clicked).grid(row=0, column=1, **pad)
        ttk.Button(button_frame, text="Take Off", command=self.takeoff_clicked).grid(row=0, column=2, **pad)
        ttk.Button(button_frame, text="Land", command=self.land_clicked).grid(row=0, column=3, **pad)

        ttk.Button(button_frame, text="Fly Square", command=lambda: self.shape_clicked("square")).grid(row=1, column=0, **pad)
        ttk.Button(button_frame, text="Fly Circle", command=lambda: self.shape_clicked("circle")).grid(row=1, column=1, **pad)
        ttk.Button(button_frame, text="Fly Triangle", command=lambda: self.shape_clicked("triangle")).grid(row=1, column=2, **pad)
        ttk.Button(button_frame, text="EMERGENCY STOP", command=self.emergency_stop_clicked).grid(row=1, column=3, **pad)

        # Status
        status_frame = ttk.LabelFrame(main, text="Status")
        status_frame.pack(fill="both", expand=True, **pad)

        self.status_text = tk.Text(status_frame, height=18, wrap="word")
        self.status_text.pack(fill="both", expand=True, padx=8, pady=8)
        self._status("Ready.")

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _status(self, msg: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {msg}\n"

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

    def _run_threaded(self, target):
        if self.busy:
            self._status("Busy. Wait for the current action to finish.")
            return
        threading.Thread(target=target, daemon=True).start()

    def _format_filename_number(self, value: float) -> str:
        return f"{value:.2f}".replace(".", "p")

    def _sanitize_filename_part(self, text: str) -> str:
        return "".join(c if c.isalnum() or c in ("_", "-") else "_" for c in text)

    # ---------------- Connection ----------------

    def connect_clicked(self):
        self._run_threaded(self.connect_cf)

    def disconnect_clicked(self):
        self._run_threaded(self.disconnect_cf)

    def connect_cf(self):
        self.busy = True
        try:
            if self.connected:
                self._status("Already connected.")
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
            self._status("Connected successfully.")

        except Exception as e:
            self._status(f"Connection failed: {e}")
            self.connected = False
            self.scf = None
            self.cf = None
            self.pc = None
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
                self._status("Disconnected.")
            else:
                self._status("Not connected.")

            self.connected = False
            self.scf = None
            self.cf = None
            self.pc = None

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

    # ---------------- Live position logging ----------------

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
            self._status("Live position display started.")
        except Exception as e:
            self.live_logconf = None
            self._status(f"Could not start live position display: {e}")

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
            self.live_x = float(data["stateEstimate.x"])
            self.live_y = float(data["stateEstimate.y"])
            self.live_z = float(data["stateEstimate.z"])

        self.root.after(0, self._update_live_labels)

    def _update_live_labels(self):
        with self.data_lock:
            self.live_x_var.set(f"{self.live_x:.3f}")
            self.live_y_var.set(f"{self.live_y:.3f}")
            self.live_z_var.set(f"{self.live_z:.3f}")

    # ---------------- Shape logging ----------------

    def start_shape_logging(self, shape_name: str, shape_size: float):
        if not self.cf:
            return

        self.log_rows = []
        self.current_shape_name = shape_name
        self.current_shape_size = shape_size
        self.shape_start_time = time.time()

        self.prev_shape_timestamp = None
        self.prev_logged_xyz = None
        self.total_loss_time_ms = 0.0
        self.loss_event_count = 0

        self.shape_logconf = LogConfig(name="ShapePositionLog", period_in_ms=SHAPE_LOG_PERIOD_MS)
        self.shape_logconf.add_variable("stateEstimate.x", "float")
        self.shape_logconf.add_variable("stateEstimate.y", "float")
        self.shape_logconf.add_variable("stateEstimate.z", "float")

        self.shape_logconf.data_received_cb.add_callback(self._shape_log_callback)
        self.cf.log.add_config(self.shape_logconf)
        self.shape_logconf.start()
        self.logging_active = True
        self._status(f"Started shape logging for {shape_name} at 100 Hz.")

    def _shape_log_callback(self, timestamp, data, logconf):
        if not self.logging_active:
            return

        x = float(data["stateEstimate.x"])
        y = float(data["stateEstimate.y"])
        z = float(data["stateEstimate.z"])

        gap_ms = 0
        loss_flag = 0
        loss_reason = ""
        estimated_lost_time_ms = 0

        if self.prev_shape_timestamp is not None:
            gap_ms = int(timestamp - self.prev_shape_timestamp)

            if gap_ms > LOSS_GAP_THRESHOLD_MS:
                loss_flag = 1
                estimated_lost_time_ms = max(0, gap_ms - EXPECTED_LOG_PERIOD_MS)
                self.total_loss_time_ms += estimated_lost_time_ms
                self.loss_event_count += 1
                loss_reason = "timestamp_gap"

            if self.prev_logged_xyz is not None:
                px, py, pz = self.prev_logged_xyz
                dist = math.sqrt((x - px) ** 2 + (y - py) ** 2 + (z - pz) ** 2)

                if gap_ms > LOSS_GAP_THRESHOLD_MS and dist < FREEZE_EPSILON_M:
                    loss_flag = 1
                    if loss_reason:
                        loss_reason += "+frozen_estimate"
                    else:
                        loss_reason = "frozen_estimate"

        with self.data_lock:
            self.log_rows.append([
                int(timestamp),              # timestamp_ms
                round(x, 5),                # x_m
                round(y, 5),                # y_m
                round(z, 5),                # z_m
                int(gap_ms),                # gap_ms
                int(loss_flag),             # loss_of_track
                int(estimated_lost_time_ms),# estimated_lost_time_ms
                loss_reason                 # loss_reason
            ])

        self.prev_shape_timestamp = timestamp
        self.prev_logged_xyz = (x, y, z)

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

        if self.log_rows:
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
            total_loss_time_ms = float(self.total_loss_time_ms)
            loss_percent = 0.0
            if total_run_time_ms > 0:
                loss_percent = 100.0 * total_loss_time_ms / total_run_time_ms

            # Save CSV
            with open(csv_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow([
                    "timestamp_ms",
                    "x_m",
                    "y_m",
                    "z_m",
                    "gap_ms",
                    "loss_of_track",
                    "estimated_lost_time_ms",
                    "loss_reason"
                ])
                writer.writerows(self.log_rows)

            self._status(f"Saved CSV: {csv_path}")
            self._status(
                f"Loss summary: events={self.loss_event_count}, "
                f"loss_time={total_loss_time_ms:.1f} ms, "
                f"loss_percent={loss_percent:.2f}%"
            )

            # Save Excel
            if pd is not None:
                try:
                    df = pd.DataFrame(self.log_rows, columns=[
                        "timestamp_ms",
                        "x_m",
                        "y_m",
                        "z_m",
                        "gap_ms",
                        "loss_of_track",
                        "estimated_lost_time_ms",
                        "loss_reason"
                    ])

                    info_df = pd.DataFrame([
                        ["shape", self.current_shape_name],
                        ["shape_size_m", shape_size],
                        ["flight_time_s", round(flight_time_s, 3)],
                        ["loss_event_count", self.loss_event_count],
                        ["estimated_total_loss_time_ms", round(total_loss_time_ms, 1)],
                        ["estimated_loss_percent", round(loss_percent, 3)],
                        ["saved_at", timestamp],
                    ], columns=["field", "value"])

                    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="FlightData")
                        info_df.to_excel(writer, index=False, sheet_name="LossSummary")

                    self._status(f"Saved Excel: {xlsx_path}")
                except Exception as e:
                    self._status(f"CSV saved, but Excel save failed: {e}")
            else:
                self._status("CSV saved. Excel not saved because pandas/openpyxl is not installed.")
        else:
            self._status("No shape data was logged.")

        self.current_shape_name = None
        self.current_shape_size = None
        self.shape_start_time = None
        self.prev_shape_timestamp = None
        self.prev_logged_xyz = None
        self.total_loss_time_ms = 0.0
        self.loss_event_count = 0

    # ---------------- Flight controls ----------------

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
                self._status("Not connected.")
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
                self._status("Not connected.")
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
                self._status("Not connected.")
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

    # ---------------- Shapes ----------------

    def fly_shape(self, shape_name: str):
        self.busy = True
        try:
            if not self.connected or not self.cf:
                self._status("Not connected.")
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

            # Return to current hover center after shape
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

    # ---------------- Close ----------------

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
    app = CrazyflieShapeGUI(root)
    root.mainloop()