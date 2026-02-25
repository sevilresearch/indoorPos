# Crazyflie Lighthouse GUI Controller with Adjustable Speed + Takeoff Height
# -------------------------------------------------------------------------
# Restores waypoint + shape controls and keeps speed/height features.

import time
import tkinter as tk
from tkinter import ttk

import cflib.crtp
from cflib.crazyflie import Crazyflie
from cflib.crazyflie.syncCrazyflie import SyncCrazyflie

URI = 'radio://0/80/2M/E7E7E7E7C3'


class DroneController:
    """Handles all Crazyflie flight logic."""

    def __init__(self):
        self.scf = None
        self.connected = False
        self.speed = 0.6
        self.takeoff_height = 0.6

    # -----------------------------
    # Connection
    # -----------------------------
    def connect(self):
        if not self.connected:
            self.scf = SyncCrazyflie(URI, cf=Crazyflie(rw_cache='./cache'))
            self.scf.open_link()

            self.scf.cf.param.set_value('commander.enHighLevel', '1')
            time.sleep(0.1)

            self.scf.cf.param.set_value('stabilizer.estimator', '2')
            time.sleep(0.1)

            self.connected = True
            print("Connected")

    def disconnect(self):
        if self.connected and self.scf is not None:
            self.scf.close_link()
            self.connected = False
            print("Disconnected")

    # -----------------------------
    # User settings
    # -----------------------------
    def set_speed(self, speed_mps):
        try:
            speed_mps = float(speed_mps)
            self.speed = max(0.05, min(speed_mps, 2.0))
            print(f"Speed set to {self.speed:.2f} m/s")
        except ValueError:
            print("Invalid speed value")

    def set_takeoff_height(self, height_m):
        try:
            height_m = float(height_m)
            self.takeoff_height = max(0.2, min(height_m, 2.5))
            print(f"Takeoff height set to {self.takeoff_height:.2f} m")
        except ValueError:
            print("Invalid height value")

    # -----------------------------
    # Flight primitives
    # -----------------------------
    def takeoff(self, duration=2.0):
        commander = self.scf.cf.high_level_commander
        commander.takeoff(self.takeoff_height, duration)
        time.sleep(duration + 0.5)

    def land(self, duration=2.0):
        commander = self.scf.cf.high_level_commander
        commander.land(0.0, duration)
        time.sleep(duration + 0.5)
        self.disconnect()

    # -----------------------------
    # Motion helper
    # -----------------------------
    def _compute_duration(self, x0, y0, z0, x1, y1, z1):
        dist = ((x1 - x0) ** 2 + (y1 - y0) ** 2 + (z1 - z0) ** 2) ** 0.5
        return max(dist / self.speed, 0.5)

    def go_to(self, x, y, z, current_pos=(0, 0, 0)):
        commander = self.scf.cf.high_level_commander
        duration = self._compute_duration(
            current_pos[0], current_pos[1], current_pos[2], x, y, z
        )
        commander.go_to(x, y, z, 0.0, duration, relative=False)
        time.sleep(duration + 0.3)

    # -----------------------------
    # Shapes
    # -----------------------------
    def fly_square(self, size):
        h = self.takeoff_height
        pts = [(0, 0, h), (size, 0, h), (size, size, h), (0, size, h), (0, 0, h)]
        cur = pts[0]
        for p in pts:
            self.go_to(*p, current_pos=cur)
            cur = p

    def fly_triangle(self, size):
        h = self.takeoff_height
        pts = [(0, 0, h), (size, 0, h), (size / 2, size, h), (0, 0, h)]
        cur = pts[0]
        for p in pts:
            self.go_to(*p, current_pos=cur)
            cur = p

    def fly_circle(self, radius, points=24):
        import math

        h = self.takeoff_height
        pts = []
        for i in range(points + 1):
            ang = 2 * math.pi * i / points
            pts.append((radius * math.cos(ang), radius * math.sin(ang), h))

        cur = pts[0]
        for p in pts:
            self.go_to(*p, current_pos=cur)
            cur = p


# ============================================================
# GUI
# ============================================================
class DroneGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Crazyflie Lighthouse Controller")

        self.drone = DroneController()

        main = ttk.Frame(root, padding=10)
        main.grid()

        # -----------------------------
        # Speed control
        # -----------------------------
        ttk.Label(main, text="Speed (m/s)").grid(column=0, row=0)
        self.speed_entry = ttk.Entry(main, width=7)
        self.speed_entry.insert(0, "0.6")
        self.speed_entry.grid(column=1, row=0)

        ttk.Button(
            main,
            text="Set Speed",
            command=lambda: self.drone.set_speed(self.speed_entry.get()),
        ).grid(column=2, row=0)

        # -----------------------------
        # Takeoff height
        # -----------------------------
        ttk.Label(main, text="Takeoff Height (m)").grid(column=0, row=1)
        self.height_entry = ttk.Entry(main, width=7)
        self.height_entry.insert(0, "0.6")
        self.height_entry.grid(column=1, row=1)

        ttk.Button(
            main,
            text="Set Height",
            command=lambda: self.drone.set_takeoff_height(self.height_entry.get()),
        ).grid(column=2, row=1)

        # -----------------------------
        # Waypoint controls
        # -----------------------------
        ttk.Label(main, text="Waypoint X").grid(column=0, row=2)
        ttk.Label(main, text="Y").grid(column=1, row=2)
        ttk.Label(main, text="Z").grid(column=2, row=2)

        self.wp_x = ttk.Entry(main, width=7)
        self.wp_y = ttk.Entry(main, width=7)
        self.wp_z = ttk.Entry(main, width=7)
        self.wp_z.insert(0, "0.6")

        self.wp_x.grid(column=0, row=3)
        self.wp_y.grid(column=1, row=3)
        self.wp_z.grid(column=2, row=3)

        ttk.Button(
            main,
            text="Go To Waypoint",
            command=self.goto_waypoint,
        ).grid(column=0, row=4, columnspan=3, sticky="ew")

        # -----------------------------
        # Shape controls
        # -----------------------------
        ttk.Label(main, text="Shape Size (m)").grid(column=0, row=5)
        self.shape_size = ttk.Entry(main, width=7)
        self.shape_size.insert(0, "1.0")
        self.shape_size.grid(column=1, row=5)

        ttk.Button(main, text="Square", command=self.fly_square).grid(column=0, row=6)
        ttk.Button(main, text="Triangle", command=self.fly_triangle).grid(column=1, row=6)
        ttk.Button(main, text="Circle", command=self.fly_circle).grid(column=2, row=6)

        # -----------------------------
        # Basic flight
        # -----------------------------
        ttk.Button(main, text="Connect", command=self.drone.connect).grid(column=0, row=7)
        ttk.Button(main, text="Takeoff", command=self.drone.takeoff).grid(column=1, row=7)
        ttk.Button(main, text="Land", command=self.drone.land).grid(column=2, row=7)

    # -----------------------------
    # GUI callbacks
    # -----------------------------
    def goto_waypoint(self):
        try:
            x = float(self.wp_x.get())
            y = float(self.wp_y.get())
            z = float(self.wp_z.get())
            self.drone.go_to(x, y, z)
        except ValueError:
            print("Invalid waypoint")

    def fly_square(self):
        try:
            size = float(self.shape_size.get())
            self.drone.fly_square(size)
        except ValueError:
            print("Invalid size")

    def fly_triangle(self):
        try:
            size = float(self.shape_size.get())
            self.drone.fly_triangle(size)
        except ValueError:
            print("Invalid size")

    def fly_circle(self):
        try:
            radius = float(self.shape_size.get())
            self.drone.fly_circle(radius)
        except ValueError:
            print("Invalid size")


# ============================================================
# Main
# ============================================================
if __name__ == '__main__':
    cflib.crtp.init_drivers()
    root = tk.Tk()
    app = DroneGUI(root)
    root.mainloop()
