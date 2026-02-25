import time
import csv
import os
import matplotlib.pyplot as plt

import cflib.crtp
from cflib.crazyflie import Crazyflie
from cflib.crazyflie.log import LogConfig
from cflib.crazyflie.syncCrazyflie import SyncCrazyflie



URI = 'radio://0/80/2M/E7E7E7E7C3'



desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
csv_filename = os.path.join(desktop_path, "Carroll_Random 2_LH.csv")

log_data = []


def log_callback(timestamp, data, logconf):
    x = data['stateEstimate.x']
    y = data['stateEstimate.y']
    z = data['stateEstimate.z']

    print(f"x={x:.2f}, y={y:.2f}, z={z:.2f}")

    log_data.append((timestamp, x, y, z))


def save_csv():
    with open(csv_filename, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["timestamp", "x", "y", "z"])
        writer.writerows(log_data)

    print(f"\nCSV saved to: {csv_filename}")


def plot_data():
    timestamps = [row[0] for row in log_data]
    x_vals = [row[1] for row in log_data]
    y_vals = [row[2] for row in log_data]
    z_vals = [row[3] for row in log_data]

    t0 = timestamps[0]
    time_sec = [(t - t0) / 1000 for t in timestamps]

    plt.figure(figsize=(12, 6))

    # Position vs time
    plt.subplot(2, 1, 1)
    plt.plot(time_sec, x_vals, label="X")
    plt.plot(time_sec, y_vals, label="Y")
    plt.plot(time_sec, z_vals, label="Z")
    plt.xlabel("Time (s)")
    plt.ylabel("Position (m)")
    plt.title("Position vs Time")
    plt.legend()
    plt.grid(True)

    # 2D Flight path
    plt.subplot(2, 1, 2)
    plt.plot(x_vals, y_vals)
    plt.xlabel("X (m)")
    plt.ylabel("Y (m)")
    plt.title("2D Flight Path (Top View)")
    plt.axis("equal")
    plt.grid(True)

    plt.tight_layout()
    plt.show()


if __name__ == "__main__":

    print("Initializing drivers...")
    cflib.crtp.init_drivers()

    with SyncCrazyflie(URI, cf=Crazyflie(rw_cache='./cache')) as scf:
        cf = scf.cf

        # Reset Kalman estimator
        cf.param.set_value('kalman.resetEstimation', '1')
        time.sleep(2)

        log_conf = LogConfig(name='Position', period_in_ms=12)
        log_conf.add_variable('stateEstimate.x', 'float')
        log_conf.add_variable('stateEstimate.y', 'float')
        log_conf.add_variable('stateEstimate.z', 'float')

        cf.log.add_config(log_conf)
        log_conf.data_received_cb.add_callback(log_callback)

        log_conf.start()

        print("\nLogging started...")
        print("Fly the drone.")
        print("Press Ctrl+C to stop logging.\n")

        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\nStopping logging...")

        log_conf.stop()

    save_csv()
    plot_data()
