import os
import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# ============================================================
# SET YOUR CSV FILE PATH HERE
# ============================================================
csv_file = r"C:\Users\ISR-Lab\Carroll\Project 2 runs\crazyflie_triangle_size1p50m_time32p67s_20260326_120329.csv"


def sanitize_filename(name):
    """
    Make a safe Windows filename from the plot title.
    """
    name = name.strip()
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    name = re.sub(r'\s+', '_', name)
    return name


def compute_3d_error(df):
    """
    Computes 3D distance between Lighthouse and MoCap points.
    Uses path_error_3d_m if already present, otherwise computes it.
    """
    required_cols = ["lh_x_m", "lh_y_m", "lh_z_m", "ot_x_m", "ot_y_m", "ot_z_m"]

    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Missing required column: {col}")

    if "path_error_3d_m" in df.columns:
        err = pd.to_numeric(df["path_error_3d_m"], errors="coerce")
    else:
        dx = pd.to_numeric(df["lh_x_m"], errors="coerce") - pd.to_numeric(df["ot_x_m"], errors="coerce")
        dy = pd.to_numeric(df["lh_y_m"], errors="coerce") - pd.to_numeric(df["ot_y_m"], errors="coerce")
        dz = pd.to_numeric(df["lh_z_m"], errors="coerce") - pd.to_numeric(df["ot_z_m"], errors="coerce")
        err = np.sqrt(dx**2 + dy**2 + dz**2)

    return err


def load_and_clean_data(csv_path):
    """
    Loads CSV and keeps only rows where both tracking systems have valid position data.
    """
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Could not find file:\n{csv_path}")

    df = pd.read_csv(csv_path)

    needed = ["lh_x_m", "lh_y_m", "lh_z_m", "ot_x_m", "ot_y_m", "ot_z_m"]
    for col in needed:
        if col not in df.columns:
            raise ValueError(f"CSV is missing required column: {col}")

    for col in needed:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    if "ot_tracked_flag" in df.columns:
        df["ot_tracked_flag"] = pd.to_numeric(df["ot_tracked_flag"], errors="coerce")

    df = df.dropna(subset=needed).copy()

    if "ot_tracked_flag" in df.columns:
        df = df[df["ot_tracked_flag"] == 1].copy()

    df["computed_error_3d_m"] = compute_3d_error(df)
    df = df.dropna(subset=["computed_error_3d_m"]).copy()

    if len(df) == 0:
        raise ValueError("No valid overlapping data points found after cleaning.")

    return df


def calculate_error_stats(error_series):
    """
    Returns mean, RMSE, and max error.
    """
    error_array = np.asarray(error_series, dtype=float)
    mean_error = np.mean(error_array)
    rmse = np.sqrt(np.mean(error_array**2))
    max_error = np.max(error_array)
    return mean_error, rmse, max_error


def set_equal_axes_3d(ax, x1, y1, z1, x2, y2, z2):
    """
    Makes 3D axes use equal scale.
    """
    xs = np.concatenate([np.asarray(x1), np.asarray(x2)])
    ys = np.concatenate([np.asarray(y1), np.asarray(y2)])
    zs = np.concatenate([np.asarray(z1), np.asarray(z2)])

    x_mid = (xs.max() + xs.min()) / 2
    y_mid = (ys.max() + ys.min()) / 2
    z_mid = (zs.max() + zs.min()) / 2

    max_range = max(
        xs.max() - xs.min(),
        ys.max() - ys.min(),
        zs.max() - zs.min()
    ) / 2

    pad = max(0.05, max_range * 0.05)
    max_range += pad

    ax.set_xlim(x_mid - max_range, x_mid + max_range)
    ax.set_ylim(y_mid - max_range, y_mid + max_range)
    ax.set_zlim(z_mid - max_range, z_mid + max_range)


def plot_flight_paths(df, plot_title, show_connector_lines=True, connector_step=50):
    """
    Plots Lighthouse and MoCap paths in 3D and saves image using the plot title.
    """
    lh_x = df["lh_x_m"].to_numpy()
    lh_y = df["lh_y_m"].to_numpy()
    lh_z = df["lh_z_m"].to_numpy()

    ot_x = df["ot_x_m"].to_numpy()
    ot_y = df["ot_y_m"].to_numpy()
    ot_z = df["ot_z_m"].to_numpy()

    err = df["computed_error_3d_m"].to_numpy()
    mean_error, rmse, max_error = calculate_error_stats(err)

    fig = plt.figure(figsize=(12, 9))
    ax = fig.add_subplot(111, projection="3d")

    ax.plot(lh_x, lh_y, lh_z, linewidth=2.5, label="Lighthouse Path")
    ax.plot(ot_x, ot_y, ot_z, linewidth=2.5, label="MoCap / OptiTrack Path")

    ax.scatter(lh_x[0], lh_y[0], lh_z[0], s=60, marker="o", label="Lighthouse Start")
    ax.scatter(ot_x[0], ot_y[0], ot_z[0], s=60, marker="o", label="MoCap Start")

    ax.scatter(lh_x[-1], lh_y[-1], lh_z[-1], s=80, marker="x", label="Lighthouse End")
    ax.scatter(ot_x[-1], ot_y[-1], ot_z[-1], s=80, marker="x", label="MoCap End")

    if show_connector_lines:
        for i in range(0, len(df), connector_step):
            ax.plot(
                [lh_x[i], ot_x[i]],
                [lh_y[i], ot_y[i]],
                [lh_z[i], ot_z[i]],
                linewidth=0.8,
                alpha=0.35
            )

    ax.set_title(plot_title, fontsize=14, pad=20)
    ax.set_xlabel("X Position (m)")
    ax.set_ylabel("Y Position (m)")
    ax.set_zlabel("Z Position (m)")
    ax.legend(loc="upper right")
    ax.grid(True)

    set_equal_axes_3d(ax, lh_x, lh_y, lh_z, ot_x, ot_y, ot_z)

    stats_text = (
        f"Samples: {len(df)}\n"
        f"Mean Error: {mean_error:.4f} m\n"
        f"RMSE: {rmse:.4f} m\n"
        f"Max Error: {max_error:.4f} m"
    )

    ax.text2D(
        0.02,
        0.98,
        stats_text,
        transform=ax.transAxes,
        fontsize=11,
        verticalalignment="top",
        bbox=dict(boxstyle="round", facecolor="white", alpha=0.85)
    )

    plt.tight_layout()

    # Save image with same title
    safe_title = sanitize_filename(plot_title)
    save_folder = os.path.dirname(csv_file)
    image_path = os.path.join(save_folder, f"{safe_title}.png")
    plt.savefig(image_path, dpi=300, bbox_inches="tight")

    print(f"Plot saved as:\n{image_path}")

    plt.show()


def main():
    try:
        df = load_and_clean_data(csv_file)

        plot_title = input("Enter plot title: ").strip()
        if plot_title == "":
            plot_title = "3D_Flight_Path_Comparison"

        plot_flight_paths(df, plot_title, show_connector_lines=True, connector_step=50)

    except Exception as e:
        print(f"ERROR: {e}")


if __name__ == "__main__":
    main()