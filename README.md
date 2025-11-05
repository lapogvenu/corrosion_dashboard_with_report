import os
import shutil
import platform 
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime, timedelta
import pandas as pd
import tempfile
from PIL import Image, ImageTk
from docx import Document
from docx.shared import Inches
import calendar
import json
from config import load_config, save_config, get_country_code, is_authorized_location

# Load configuration
config = load_config()
import calendar
if not config['country_code']:
    country_code = simpledialog.askstring("Country Code", "Please enter your country code (e.g., DE, IN, US):")
    if country_code:
        config['country_code'] = country_code.upper()
        save_config(config)

# Basic app/window and project defaults
root = tk.Tk()
root.title(f"Corrosion Test Dashboard - {config['country_code']}")
root.geometry("1200x800")

# File system and UI globals
BASE_DIR = os.path.join(config.get('local_drive_path', '.'), 'parts')
SHARED_DIR = os.path.join(config.get('shared_folder_path', '.'), 'shared_reports')
breakdown_history_file = os.path.join(config.get('local_drive_path', '.'), f'breakdown_history_{config["country_code"]}.xlsx')

# Create directories if they don't exist
os.makedirs(BASE_DIR, exist_ok=True)
os.makedirs(SHARED_DIR, exist_ok=True)

# Columns for the Treeview and report
columns = [
    "Project ID", "Part Name", "Part Number", "vLims ID", "Type of chamber",
    "Test standard", "Background details", "Start Date", "End Date",
    "Spec Hours", "Actual Hours", "Status", "Result"
]

# Image/report defaults
report_canvas_w = 900
report_canvas_h = 600
report_embed_width = 2.6
report_max_images = 0
report_include_timestamp = True

# Runtime lists
photos = []
img_paths = []

# Initialize auto-save
def init_auto_save():
    """Initialize auto-save functionality"""
    # Load existing data if available
    if os.path.exists(breakdown_history_file):
        try:
            df = pd.read_excel(breakdown_history_file)
            for _, row in df.iterrows():
                values = row.values.tolist()
                tag = get_result_tag(values[-1], values[-2])
                tree.insert("", tk.END, values=values, tags=(tag,))
        except Exception as e:
            messagebox.showwarning("Data Load Error", f"Failed to load existing data: {str(e)}")
    
    # Start auto-save timer
    auto_save_data()

# Initialize shared paths
def init_shared_paths():
    """Initialize shared paths for multi-location access"""
    if not config.get('shared_folder_path'):
        shared_path = filedialog.askdirectory(title="Select Shared Network Folder")
        if shared_path:
            config['shared_folder_path'] = shared_path
            save_config(config)
    
    if not config.get('local_drive_path'):
        local_path = filedialog.askdirectory(title="Select Local Drive Folder")
        if local_path:
            config['local_drive_path'] = local_path
            save_config(config)

def get_result_tag(result_text, status_text):
    """Return a tree tag name based on result text and status."""
    try:
        res = (result_text or "").strip().lower()
        status = (status_text or "").strip().lower()
    except Exception:
        return "inprogress"
    if status == "in-progress":
        return "inprogress"
    if status == "completed":
        if res in ("", None):
            return "completed_blink"
        if any(k in res for k in ("pass", "ok", "success")):
            return "ok_static"
        if any(k in res for k in ("fail", "ng", "reject")):
            return "fail_blink"
        return "completed_blink"
    # default
    if any(k in res for k in ("pass", "ok", "success")):
        return "ok_static"
    if any(k in res for k in ("fail", "ng", "reject")):
        return "fail_blink"
    return "inprogress"

def parse_datetime(date_str):
    try:
        return datetime.strptime(date_str + " 09:00", "%d-%m-%Y %H:%M")
    except:
        return None

def calculate_actual_hours(start_dt, end_dt):
    now = datetime.now()
    if now < start_dt:
        return 0
    effective_end = min(now, end_dt)
    delta_days = (effective_end.date() - start_dt.date()).days
    hours = delta_days * 23
    td_start = timedelta(hours=start_dt.hour, minutes=start_dt.minute, seconds=start_dt.second)
    td_end = timedelta(hours=effective_end.hour, minutes=effective_end.minute, seconds=effective_end.second)
    partial_hours = (td_end - td_start).total_seconds() / 3600
    if partial_hours < 0:
        partial_hours += 23
    partial_hours = min(partial_hours, 23)
    total_hours = hours + partial_hours
    return round(total_hours, 2)

# Define functions for dashboard operations
def apply_row_colors():
    for iid in tree.get_children():
        vals = list(tree.item(iid, "values"))
        spec_hours = float(vals[9])
        actual_hours = float(vals[10])
        status = "Completed" if round(actual_hours, 2) >= spec_hours else "In-Progress"
        if vals[11] != status:
            vals[11] = status
            tree.item(iid, values=vals)
        tag = get_result_tag(vals[12], vals[11])
        tree.item(iid, tags=(tag,))
    tree.tag_configure("ok_static", background="#9FE6A8")
    tree.tag_configure("completed_blink", background="#9FE6A8")
    tree.tag_configure("fail_blink", background="#ff9090")
    tree.tag_configure("inprogress", background="#fff9d6")

def auto_save_data():
    """Automatically save data to Excel file"""
    try:
        data = []
        for item in tree.get_children():
            values = tree.item(item)['values']
            data.append(dict(zip(columns, values)))
        
        df = pd.DataFrame(data)
        df.to_excel(breakdown_history_file, index=False)
        
        # Schedule next auto-save
        root.after(config['auto_save_interval'] * 1000, auto_save_data)
    except Exception as e:
        messagebox.showerror("Auto-save Error", f"Failed to auto-save: {str(e)}")

def update_chamber_counts():
    nss_count = 0
    cct_count = 0
    for iid in tree.get_children():
        values = tree.item(iid, "values")
        if len(values) > 4:
            chamber_type = values[4]
            if chamber_type == "NSS": nss_count += 1
            elif chamber_type == "CCT": cct_count += 1
    lbl_nss_count.config(text=str(nss_count))
    lbl_cct_count.config(text=str(cct_count))

def add_record():
    project_id = entry_project_id.get()
    part_name = entry_part_name.get()
    part_number = entry_part_number.get()
    vlims_id = entry_vlims_id.get()
    type_chamber = type_chamber_var.get()
    test_standard = entry_test_standard.get()
    background_details = entry_background_details.get()
    date_str = entry_start_date.get()
    spec_hours_str = entry_spec_hours.get()
    result = entry_result_update.get()
    if not all([project_id, part_name, part_number, date_str, spec_hours_str]):
        messagebox.showwarning("Missing Data", "Please fill all required fields.")
        return
    start_date = parse_datetime(date_str)
    if not start_date:
        messagebox.showerror("Date Error", "Invalid start date format. Use DD-MM-YYYY.")
        return
    try:
        spec_hours = float(spec_hours_str)
    except:
        messagebox.showerror("Input Error", "Spec Hours must be a number.")
        return
    end_date = start_date + timedelta(hours=spec_hours)
    actual_hours = calculate_actual_hours(start_date, end_date)
    status = "Completed" if round(actual_hours, 2) >= spec_hours else "In-Progress"
    tag = get_result_tag(result, status)
    tree.insert("", tk.END, values=(
        project_id, part_name, part_number, vlims_id, type_chamber,
        test_standard, background_details, start_date.strftime("%d-%m-%Y"),
        end_date.strftime("%d-%m-%Y"), spec_hours, actual_hours, status, result),
        tags=(tag,))
    for entry in entries:
        entry.delete(0, tk.END)
    entry_result_update.delete(0, tk.END)
    apply_row_colors()
    update_chamber_counts()

def update_actual_hours_all():
    for iid in tree.get_children():
        vals = list(tree.item(iid, "values"))
        start_dt = parse_datetime(vals[7])
        end_dt = parse_datetime(vals[8])
        if not (start_dt and end_dt):
            continue
        vals[10] = calculate_actual_hours(start_dt, end_dt)
        tree.item(iid, values=vals)

def update_result():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Selection Required", "Select a row to update result.")
        return
    iid = selected[0]
    vals = list(tree.item(iid, "values"))
    if vals[11] == "In-Progress":
        messagebox.showinfo("Update Blocked", "Cannot update result when test is In-Progress.")
        return
    res_text = entry_result_update.get()
    vals[12] = res_text
    tag = get_result_tag(res_text, vals[11])
    tree.item(iid, values=vals, tags=(tag,))
    entry_result_update.delete(0, tk.END)
    apply_row_colors()

def delete_selected():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Delete Row", "Select a row to delete.")
        return
    tree.delete(selected[0])
    update_chamber_counts()

def update_selected_details():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Update Row", "Select a row first.")
        return
    iid = selected[0]
    vals = list(tree.item(iid, "values"))
    new_vals = [
        entry_project_id.get(), entry_part_name.get(), entry_part_number.get(),
        entry_vlims_id.get(), type_chamber_var.get(),
        entry_test_standard.get(), entry_background_details.get(),
        entry_start_date.get(), vals[8],
        entry_spec_hours.get(), vals[10], vals[11], vals[12]
    ]
    start_date = parse_datetime(new_vals[7])
    try:
        spec_hours = float(new_vals[9])
    except:
        spec_hours = 0
    if start_date and spec_hours > 0:
        end_date = start_date + timedelta(hours=spec_hours)
        new_vals[8] = end_date.strftime("%d-%m-%Y")
        new_vals[10] = calculate_actual_hours(start_date, end_date)
        new_vals[11] = "Completed" if round(new_vals[10], 2) >= spec_hours else "In-Progress"
    tree.item(iid, values=new_vals)
    apply_row_colors()
    update_chamber_counts()
    messagebox.showinfo("Detail Updated", "Selected row has been updated with new values from entry fields.")

def open_breakdown_popup():
    popup = tk.Toplevel(root)
    popup.title("Breakdown Duration Input")
    popup.geometry("480x185")
    frame_dates = tk.Frame(popup)
    frame_dates.pack(pady=3)
    tk.Label(frame_dates, text="From Date (DD-MM-YYYY):").grid(row=0, column=0, padx=4)
    from_date_entry = tk.Entry(frame_dates, width=14)
    from_date_entry.grid(row=0, column=1, padx=4)
    tk.Label(frame_dates, text="From Time (HH:MM, 24h):").grid(row=0, column=2, padx=4)
    from_time_entry = tk.Entry(frame_dates, width=10)
    from_time_entry.grid(row=0, column=3, padx=4)
    tk.Label(frame_dates, text="To Date (DD-MM-YYYY):").grid(row=1, column=0, padx=4)
    to_date_entry = tk.Entry(frame_dates, width=14)
    to_date_entry.grid(row=1, column=1, padx=4)
    tk.Label(frame_dates, text="To Time (HH:MM, 24h):").grid(row=1, column=2, padx=4)
    to_time_entry = tk.Entry(frame_dates, width=10)
    to_time_entry.grid(row=1, column=3, padx=4)
    tk.Label(popup, text="Breakdown Reason:").pack(pady=2)
    breakdown_reason_entry = tk.Entry(popup, width=50)
    breakdown_reason_entry.pack(pady=2)
    def apply_breakdown_and_save():
        from_date_str = from_date_entry.get()
        from_time_str = from_time_entry.get()
        to_date_str = to_date_entry.get()
        to_time_str = to_time_entry.get()
        reason = breakdown_reason_entry.get().strip()
        try:
            from_dt = datetime.strptime(from_date_str + " " + from_time_str, "%d-%m-%Y %H:%M")
            to_dt = datetime.strptime(to_date_str + " " + to_time_str, "%d-%m-%Y %H:%M")
        except Exception:
            messagebox.showerror("Invalid input", "Please enter valid dates and times in required format.")
            return
        if to_dt <= from_dt:
            messagebox.showerror("Input Error", "To datetime must be after From datetime!")
            return
        breakdown_hours = (to_dt - from_dt).total_seconds() / 3600
        new_row = {
            "From": from_dt.strftime("%d-%m-%Y %H:%M"),
            "To": to_dt.strftime("%d-%m-%Y %H:%M"),
            "Hours": round(breakdown_hours, 2),
            "Reason": reason
        }
        if os.path.exists(breakdown_history_file):
            df = pd.read_excel(breakdown_history_file)
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        else:
            df = pd.DataFrame([new_row])
        df.to_excel(breakdown_history_file, index=False)
        for iid in tree.get_children():
            vals = list(tree.item(iid, "values"))
            part_start = parse_datetime(vals[7])
            part_end = parse_datetime(vals[8])
            if not part_start or not part_end:
                continue
            latest_start = max(part_start, from_dt)
            earliest_end = min(part_end, to_dt)
            overlap = (earliest_end - latest_start).total_seconds() / 3600
            overlap = max(0, overlap)
            actual_hours = float(vals[10])
            actual_hours = max(0, actual_hours - overlap)
            vals[10] = round(actual_hours, 2)
            tree.item(iid, values=vals)
        apply_row_colors()
        update_chamber_counts()
        draw_monthly_breakdown_boxes(frame_calendar, breakdown_history_file)
        messagebox.showinfo("Breakdown Applied", "Breakdown saved & subtracted from actual hours for affected parts.")
        popup.destroy()
    btns = tk.Frame(popup)
    btns.pack(pady=6)
    tk.Button(btns, text="Apply & Save", bg="#D35400", fg="white", width=16, command=apply_breakdown_and_save).pack(side="left", padx=8)
    tk.Button(btns, text="Cancel", width=16, command=popup.destroy).pack(side="left", padx=8)

def open_breakdown_history():
    if not os.path.exists(breakdown_history_file):
        messagebox.showinfo("History", "No breakdown history found.")
        return
    try:
        if platform.system() == "Windows":
            os.startfile(breakdown_history_file)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", breakdown_history_file])
        else:
            subprocess.Popen(["xdg-open", breakdown_history_file])
    except Exception as e:
        messagebox.showerror("Error", f"Could not open breakdown history:\n{e}")

def export_data():
    data = [tree.item(iid, "values") for iid in tree.get_children()]
    if not data:
        messagebox.showinfo("No Data", "No records to export.")
        return
    df = pd.DataFrame(data, columns=columns)
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel Files", "*.xlsx")],
                                             title="Save Dashboard Data As")
    if not file_path:
        return
    try:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Export Success", f"Dashboard exported to\n{file_path}")
    except Exception as e:
        messagebox.showerror("Export Error", f"Failed to export data:\n{e}")

def upload_photo_for_selected():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Select Row", "Select a row first.")
        return
    vals = tree.item(selected[0], "values")
    # values: (Project ID, Part Name, Part Number, vLims ID, ...)
    # use Part Number (index 2) and Part Name (index 1)
    part_number, part_name, actual_hours = vals[2], vals[1], float(vals[10])
    top = tk.Toplevel(root)
    top.title(f"Upload Photos for {part_number} {part_name}")
    top.geometry("400x450")
    tk.Label(top, text="Select Timeline or enter custom:").pack(pady=5)
    timeline_vars = ["Received", "24h", "96h", "240h", "480h", "720h"]
    timeline_box = ttk.Combobox(top, values=timeline_vars, width=25, state="readonly")
    timeline_box.pack()
    timeline_box.set("")
    custom_timeline = tk.Entry(top, width=25)
    tk.Label(top, text="Custom Timeline (optional):").pack(pady=5)
    custom_timeline.pack()
    images_listbox = tk.Listbox(top, selectmode='multiple', width=50, height=8)
    images_listbox.pack(pady=10)
    selected_paths = []
    def select_images():
        nonlocal selected_paths
        paths = filedialog.askopenfilenames(filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp")])
        if paths:
            selected_paths = list(paths)
            images_listbox.delete(0, tk.END)
            for p in selected_paths:
                images_listbox.insert(tk.END, os.path.basename(p))
    def timeline_allowed(timeline_val):
        try:
            h = int(''.join(filter(str.isdigit, timeline_val)))
            return h <= actual_hours
        except:
            return True
    tk.Button(top, text="Browse Images", command=select_images).pack(pady=7)
    def save_photos():
        timeline_val = timeline_box.get().strip()
        custom_val = custom_timeline.get().strip()
        timeline_folder = custom_val if custom_val else timeline_val
        if timeline_folder == "":
            messagebox.showwarning("Timeline", "Please specify a timeline.")
            return
        if not selected_paths:
            messagebox.showwarning("No Images", "Please select images to upload.")
            return
        if not timeline_allowed(timeline_folder):
            messagebox.showerror("Error", "Timeline hours exceed actual hours.")
            return
        save_dir = os.path.join(BASE_DIR, f"{part_number}_{part_name}", timeline_folder)
        os.makedirs(save_dir, exist_ok=True)
        try:
            for img_path in selected_paths:
                shutil.copy(img_path, os.path.join(save_dir, os.path.basename(img_path)))
            messagebox.showinfo("Success", f"{len(selected_paths)} photos saved.")
            top.destroy()
        except Exception as e:
            messagebox.showerror("Save Error", f"Error saving images:\n{e}")
    tk.Button(top, text="Save", command=save_photos, bg="#4CAF50", fg="white", width=15).pack(pady=10)
    tk.Button(top, text="Cancel", command=top.destroy).pack()

def view_images_for_selected():
    global photos, img_paths
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Select Row", "Select a row to view images.")
        return
    vals = tree.item(selected[0], "values")
    # values: (Project ID, Part Name, Part Number, ...)
    part_number, part_name = vals[2], vals[1]
    part_folder = os.path.join(BASE_DIR, f"{part_number}_{part_name}")
    if not os.path.exists(part_folder):
        messagebox.showinfo("No Images", "No images folders found.")
        return
    timelines = [d for d in os.listdir(part_folder) if os.path.isdir(os.path.join(part_folder, d))]
    if not timelines:
        messagebox.showinfo("No Timeline Folders", "No timeline folders found.")
        return
    top = tk.Toplevel(root)
    top.title(f"View Images for {part_number} {part_name}")
    top.geometry("1000x600")
    timeline_var = tk.StringVar()
    timeline_box = ttk.Combobox(top, values=timelines, textvariable=timeline_var, state="readonly", width=30)
    timeline_box.pack(pady=10)
    timeline_box.set("Select Timeline")
    submit_btn = tk.Button(top, text="Load", width=14)
    submit_btn.pack(pady=5)
    container = tk.Frame(top)
    container.pack(fill="both", expand=True)
    canvas = tk.Canvas(container, bg="white")
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")
    scrollable_frame = tk.Frame(canvas)
    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    photos.clear()
    img_paths.clear()

    def open_viewer(start_index):
        viewer = tk.Toplevel(root)
        viewer.title(f"Image Viewer ({start_index+1})")
        viewer.geometry("900x700")

        canvas2 = tk.Canvas(viewer, bg="white")
        hscroll2 = tk.Scrollbar(viewer, orient="horizontal", command=canvas2.xview)
        vscroll2 = tk.Scrollbar(viewer, orient="vertical", command=canvas2.yview)
        canvas2.config(xscrollcommand=hscroll2.set, yscrollcommand=vscroll2.set)
        canvas2.pack(side="top", fill="both", expand=True)
        hscroll2.pack(side="bottom", fill="x")
        vscroll2.pack(side="right", fill="y")

        button_frame2 = tk.Frame(viewer)
        button_frame2.pack(fill="x", side="bottom", pady=8)

        idx = [start_index]
        scale = [1.0]

        lbl_index = tk.Label(button_frame2, font=("Arial", 11))
        lbl_index.pack(side="top", pady=2)

        def show_image():
            img = Image.open(img_paths[idx[0]])
            img_w, img_h = img.size
            disp_w = int(img_w * scale[0])
            disp_h = int(img_h * scale[0])
            win_w = canvas2.winfo_width()
            win_h = canvas2.winfo_height()
            img_disp = img.resize((disp_w, disp_h), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img_disp)
            canvas2.delete("all")
            cx = max((win_w - disp_w) // 2, 0)
            cy = max((win_h - disp_h) // 2, 0)
            canvas2.create_image(cx, cy, anchor="nw", image=photo)
            canvas2.image = photo
            canvas2.config(scrollregion=(0, 0, max(win_w, disp_w), max(win_h, disp_h)))
            lbl_index.config(text=f"Image {idx[0]+1} of {len(img_paths)}: {os.path.basename(img_paths[idx[0]])}   |   Zoom: {int(scale[0]*100)}%")

        def fit_to_window(event=None):
            img = Image.open(img_paths[idx[0]])
            win_w = canvas2.winfo_width()
            win_h = canvas2.winfo_height()
            if win_w <= 20 or win_h <= 20:
                return
            img_w, img_h = img.size
            s = min(win_w / img_w, win_h / img_h, 1.0)
            scale[0] = s
            show_image()

        def zoom_in():
            scale[0] *= 1.15
            show_image()

        def zoom_out():
            scale[0] /= 1.15
            show_image()

        def next_image(event=None):
            idx[0] = (idx[0] + 1) % len(img_paths)
            fit_to_window()

        def prev_image(event=None):
            idx[0] = (idx[0] - 1) % len(img_paths)
            fit_to_window()

        def on_mouse_wheel(event):
            if getattr(event, "delta", 0) < 0:
                next_image()
            else:
                prev_image()

        btn_prev = tk.Button(button_frame2, text="<< Previous", command=prev_image, width=12)
        btn_next = tk.Button(button_frame2, text="Next >>", command=next_image, width=12)
        btn_zoomin = tk.Button(button_frame2, text="Zoom In (+)", command=zoom_in, width=10)
        btn_zoomout = tk.Button(button_frame2, text="Zoom Out (-)", command=zoom_out, width=10)
        btn_prev.pack(side="left", padx=15)
        btn_zoomout.pack(side="left", padx=15)
        btn_zoomin.pack(side="left", padx=0)
        btn_next.pack(side="right", padx=15)

        viewer.bind("<Left>", prev_image)
        viewer.bind("<Right>", next_image)
        viewer.bind("<plus>", lambda event: zoom_in())
        viewer.bind("<minus>", lambda event: zoom_out())
        viewer.bind("<MouseWheel>", on_mouse_wheel)
        viewer.bind("<Configure>", lambda e: fit_to_window())

        fit_to_window()

    def load_thumbnails():
        photos.clear()
        img_paths.clear()
        for widget in scrollable_frame.winfo_children():
            widget.destroy()
        timeline = timeline_var.get()
        if timeline == "" or timeline == "Select Timeline":
            messagebox.showwarning("Select Timeline", "Select timeline to view photos.")
            return
        folder = os.path.join(part_folder, timeline)
        if not os.path.exists(folder):
            messagebox.showinfo("No Photos", "No photos found in this timeline.")
            return
        files = [f for f in os.listdir(folder) if f.lower().endswith((".jpg", ".jpeg", ".png", ".bmp"))]
        if not files:
            messagebox.showinfo("No Photos", "No photos found.")
            return
        for i, file in enumerate(files):
            path = os.path.join(folder, file)
            try:
                img = Image.open(path)
                img.thumbnail((150, 150))
                photo = ImageTk.PhotoImage(img)
            except Exception:
                continue
            photos.append(photo)
            img_paths.append(path)
            btn = tk.Button(scrollable_frame, image=photo, width=155, height=155,
                            command=lambda idx=i: open_viewer(idx))
            btn.grid(row=i // 4, column=i % 4, padx=5, pady=5)

    submit_btn.config(command=load_thumbnails)

# UI Layout (header, inputs, buttons, table, calendar)
frame_header = tk.Frame(root, bg="white", pady=12)
frame_header.pack(fill="x")

header_left = tk.Frame(frame_header, bg="white")
header_left.pack(side="left", anchor="nw", padx=14)

header_center = tk.Frame(frame_header, bg="white")
header_center.pack(side="left", expand=True)

header_right = tk.Frame(frame_header, bg="white")
header_right.pack(side="right", anchor="ne", padx=12)

count_frame = tk.Frame(header_left, bg="white")
count_frame.pack(anchor="nw", pady=10)

lbl_nss_frame = tk.LabelFrame(count_frame, text="NSS Count", bg="#E8F5E9", fg="#1B5E20",
                              font=("Arial", 12, "bold"), padx=12, pady=4)
lbl_nss_frame.pack(side="top", pady=3)

lbl_nss_count = tk.Label(lbl_nss_frame, text="0", font=("Arial", 16, "bold"),
                         fg="#1B5E20", bg="#E8F5E9")
lbl_nss_count.pack()

lbl_cct_frame = tk.LabelFrame(count_frame, text="CCT Count", bg="#E3F2FD", fg="#0D47A1",
                              font=("Arial", 12, "bold"), padx=12, pady=4)
lbl_cct_frame.pack(side="top", pady=3)

lbl_cct_count = tk.Label(lbl_cct_frame, text="0", font=("Arial", 16, "bold"),
                         fg="#0D47A1", bg="#E3F2FD")
lbl_cct_count.pack()

lbl_title = tk.Label(header_center, text="Corrosion Test Dashboard",
                     font=("Arial", 30, "bold"), bg="white")
lbl_title.pack(pady=(2, 2))

lbl_date = tk.Label(header_center, text="", font=("Arial", 15, "bold"), bg="white")
lbl_date.pack()

def update_date():
    now = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    lbl_date.config(text=now)
    root.after(1000, update_date)

update_date()

frame_calendar = tk.Frame(header_right, bg="white")
frame_calendar.pack()

def draw_monthly_breakdown_boxes(parent, breakdown_file):
    for widget in parent.winfo_children():
        widget.destroy()
    month_hours = [0] * 12
    if os.path.exists(breakdown_file):
        df = pd.read_excel(breakdown_file)
        if 'From' in df:
            df['From'] = pd.to_datetime(df['From'], format="%d-%m-%Y %H:%M")
            for idx, row in df.iterrows():
                m = row['From'].month-1
                month_hours[m] += row['Hours']
    box_w, box_h = 110, 70
    pad_x, pad_y = 11, 10
    months = list(calendar.month_abbr)[1:]
    c = tk.Canvas(parent, width=(box_w+pad_x)*4, height=(box_h+pad_y)*3, bg="white", highlightthickness=0)
    c.pack()
    for i in range(12):
        r, col = divmod(i, 4)
        bx = pad_x + col*(box_w+pad_x)
        by = pad_y + r*(box_h+pad_y)
        fillcolor = "#C8E6C9" if month_hours[i] == 0 else "#FFCDD2"
        htext = "0" if month_hours[i] == 0 else f"{month_hours[i]:.0f}h"
        c.create_rectangle(bx, by, bx+box_w, by+box_h, fill=fillcolor, outline="black", width=2)
        c.create_text(bx+box_w/2, by+box_h/2-10, text=months[i], font=("Arial", 16, "bold"))
        c.create_text(bx+box_w/2, by+box_h/2+15, text=htext, font=("Arial", 15, "bold"), fill="#232323")

draw_monthly_breakdown_boxes(frame_calendar, breakdown_history_file)

def open_images_folder_selected():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Select Row", "Please select a part row.")
        return
    vals = tree.item(selected[0], "values")
    # values: (Project ID, Part Name, Part Number, ...)
    part_number, part_name = vals[2], vals[1]
    folder = os.path.join(BASE_DIR, f"{part_number}_{part_name}")
    if not os.path.isdir(folder):
        messagebox.showinfo("No Folder", "No images folder exists for selected part.")
        return
    try:
        if platform.system() == "Windows":
            os.startfile(folder)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])
    except Exception as e:
        messagebox.showerror("Error", f"Could not open folder:\n{e}")


# Report settings removed per user request

def generate_report():
    selected = tree.selection()
    if not selected:
        print("Please select a part row.")
        return

    vals = tree.item(selected[0], "values")
    if len(vals) < len(columns):
        print("Selected row does not contain all required data.")
        return

    data = dict(zip(columns, vals))
    part_num, part_name = data["Part Number"], data["Part Name"]
    folder = os.path.join(BASE_DIR, f"{part_num}_{part_name}")

    if not os.path.isdir(folder):
        os.makedirs(folder)
        print(f"Created folder for part: {folder}")

    try:
        # If a report template path is configured and exists, use it as the base document
        tpl_path = config.get('report_template_path', '')
        if not tpl_path or not os.path.exists(tpl_path):
            # Prompt user to select a template if not configured or missing
            tpl = filedialog.askopenfilename(title='Select Report Template (.docx)', filetypes=[('Word Documents','*.docx')])
            if not tpl:
                messagebox.showinfo("Template", "No template selected. Report generation cancelled.")
                return
            tpl_path = tpl
            config['report_template_path'] = tpl_path
            try:
                save_config(config)
            except Exception:
                pass
        try:
            doc = Document(tpl_path)
        except Exception as e:
            messagebox.showerror("Template Error", f"Could not open template: {e}")
            return

        # Find the tables in the template
        tables = doc.tables
        if not tables:
            messagebox.showerror("Template Error", "Template must contain tables for data")
            return

        # Find and update Part Number and Background details in the first table
        try:
            part_details_table = tables[0]  # First table should be for part details
            for row in part_details_table.rows:
                if len(row.cells) >= 2:
                    cell_text = row.cells[0].text.strip().lower()
                    if "rough part drawing" in cell_text or "part number" in cell_text:
                        row.cells[1].text = data["Part Number"]
                    elif "test reason" in cell_text or "background" in cell_text:
                        row.cells[1].text = data["Background details"]
        except Exception as e:
            print(f"Error updating part details: {e}")

        # Find the photos table (usually the last table) and prepare to insert timeline photos
        photos_table = tables[-1]  # Last table should be for photos
        if photos_table:
            # Clear any existing content (keep header row if present)
            while len(photos_table.rows) > 1:
                photos_table._element.remove(photos_table.rows[-1]._element)

        # Replace placeholder tokens in the template with data values.
        # Template should contain placeholders like {{Part Name}}, {{Part Number}}, {{Background details}}, etc.
        placeholders = {
            "{{Project ID}}": str(data.get("Project ID", "")),
            "{{Part Name}}": str(data.get("Part Name", "")),
            "{{Part Number}}": str(data.get("Part Number", "")),
            "{{vLims ID}}": str(data.get("vLims ID", "")),
            "{{Type of chamber}}": str(data.get("Type of chamber", "")),
            "{{Test standard}}": str(data.get("Test standard", "")),
            "{{Background details}}": str(data.get("Background details", "")),
            "{{Start Date}}": str(data.get("Start Date", "")),
            "{{End Date}}": str(data.get("End Date", "")),
            "{{Spec Hours}}": str(data.get("Spec Hours", "")),
            "{{Actual Hours}}": str(data.get("Actual Hours", "")),
            "{{Status}}": str(data.get("Status", "")),
            "{{Result}}": str(data.get("Result", ""))
        }

        def replace_in_paragraph(paragraph, ph, val):
            if ph in paragraph.text:
                # Replace inside runs to preserve styling
                for run in paragraph.runs:
                    if ph in run.text:
                        run.text = run.text.replace(ph, val)

        # Replace in paragraphs
        for p in doc.paragraphs:
            for ph, val in placeholders.items():
                replace_in_paragraph(p, ph, val)

        # Replace in tables (cells)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for ph, val in placeholders.items():
                        if ph in cell.text:
                            cell.text = cell.text.replace(ph, val)

        # Prepare timelines (subfolders under the part folder)
        if not os.path.isdir(folder):
            timelines = []
        else:
            try:
                timelines = [d for d in os.listdir(folder) if os.path.isdir(os.path.join(folder, d))]
            except Exception:
                timelines = []

        def timeline_sort_key(name):
            low = name.strip().lower()
            if low == 'received':
                return (0, name)
            digits = ''.join(ch for ch in name if ch.isdigit())
            try:
                return (1, int(digits))
            except Exception:
                return (2, name)

        timelines = sorted(timelines, key=timeline_sort_key)

        temp_files = []
        # Insert photos into the photos_table using the template's structure
        if timelines and photos_table is not None:
            for tl in timelines:
                tl_folder = os.path.join(folder, tl)
                imgs = []
                for r, _, files in os.walk(tl_folder):
                    for f in files:
                        if f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp')):
                            imgs.append(os.path.join(r, f))
                if not imgs:
                    row = photos_table.add_row().cells
                    row[0].text = tl
                    row[1].text = 'No photo found'
                    continue
                imgs = sorted(imgs)
                if report_max_images and report_max_images > 0:
                    imgs = imgs[:report_max_images]

                row = photos_table.add_row().cells
                row[0].text = tl
                # For the photos cell, add each image as its own paragraph
                for img_path in imgs:
                    try:
                        img = Image.open(img_path)
                        target_w, target_h = report_canvas_w, report_canvas_h
                        w, h = img.size
                        scale = min(float(target_w) / w, float(target_h) / h, 1.0)
                        new_w = int(w * scale)
                        new_h = int(h * scale)
                        img_resized = img.resize((new_w, new_h), Image.LANCZOS)
                        canvas_img = Image.new('RGB', (target_w, target_h), (255, 255, 255))
                        paste_x = (target_w - new_w) // 2
                        paste_y = (target_h - new_h) // 2
                        canvas_img.paste(img_resized.convert('RGB'), (paste_x, paste_y))
                        tf = tempfile.NamedTemporaryFile(delete=False, suffix='.jpg')
                        tmp_path = tf.name
                        try:
                            canvas_img.save(tmp_path, format='JPEG', quality=85)
                        finally:
                            tf.close()
                        temp_files.append(tmp_path)

                        para = row[1].add_paragraph()
                        run = para.add_run()
                        run.add_picture(tmp_path, width=Inches(report_embed_width))
                    except Exception as e:
                        row[1].add_paragraph(f'Error adding image {os.path.basename(img_path)}: {e}')

        # Prompt user where to save the report (Save As dialog)
        initial_name = f"{part_name}_Part_Report.docx"
        # Suggest the part folder as the initial directory if it exists, otherwise use cwd
        start_dir = folder if os.path.isdir(folder) else os.getcwd()
        save_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                                 initialdir=start_dir,
                                                 initialfile=initial_name,
                                                 filetypes=[("Word Documents", "*.docx")],
                                                 title="Save Report As")
        if not save_path:
            print("Report save cancelled by user.")
            # cleanup temporary files if any
            try:
                for p in locals().get('temp_files', []) or []:
                    try:
                        os.remove(p)
                    except Exception:
                        pass
            except Exception:
                pass
            return

        doc.save(save_path)
        print(f"Report saved at: {save_path}")

        # cleanup temporary files
        try:
            for p in locals().get('temp_files', []) or []:
                try:
                    os.remove(p)
                except Exception:
                    pass
        except Exception:
            pass

    except Exception as e:
        print(f"Failed to generate report: {e}")


def generate_report_default():
    """Generate the report using the old/default layout (no template)."""
    selected = tree.selection()
    if not selected:
        print("Please select a part row.")
        return

    vals = tree.item(selected[0], "values")
    if len(vals) < len(columns):
        print("Selected row does not contain all required data.")
        return

    data = dict(zip(columns, vals))
    part_num, part_name = data["Part Number"], data["Part Name"]
    folder = os.path.join(BASE_DIR, f"{part_num}_{part_name}")

    if not os.path.isdir(folder):
        os.makedirs(folder, exist_ok=True)

    try:
        doc = Document()

        # Add Part Information as a 2-column table
        doc.add_heading("Part Information", level=1)
        info_table = doc.add_table(rows=0, cols=2)
        info_table.style = 'Table Grid'
        for key in columns[:7]:
            row_cells = info_table.add_row().cells
            row_cells[0].text = str(key)
            row_cells[1].text = str(data.get(key, ''))

        # Table for test results
        doc.add_heading("Test Results", level=1)
        table = doc.add_table(rows=1, cols=6)
        table.style = 'Table Grid'
        headers = columns[7:]
        hdr_cells = table.rows[0].cells
        for i, h in enumerate(headers):
            hdr_cells[i].text = h

        row_cells = table.add_row().cells
        for i, h in enumerate(headers):
            row_cells[i].text = str(data.get(h, ''))

        # Photos by timeline: 2-column table (Hours | Photo)
        doc.add_heading("Photos by Timeline", level=1)

        if not os.path.isdir(folder):
            doc.add_paragraph('No photos folder found for this part.')
            timelines = []
        else:
            try:
                timelines = [d for d in os.listdir(folder) if os.path.isdir(os.path.join(folder, d))]
            except Exception as e:
                doc.add_paragraph(f'Error reading timeline folders: {e}')
                timelines = []

        def timeline_sort_key(name):
            low = name.strip().lower()
            if low == 'received':
                return (0, name)
            digits = ''.join(ch for ch in name if ch.isdigit())
            try:
                return (1, int(digits))
            except Exception:
                return (2, name)

        timelines = sorted(timelines, key=timeline_sort_key)

        if not timelines:
            doc.add_paragraph('No timeline photo folders found in: ' + folder)
        else:
            photo_table = doc.add_table(rows=1, cols=2)
            photo_table.style = 'Table Grid'
            hdr = photo_table.rows[0].cells
            hdr[0].text = 'Hours'
            hdr[1].text = 'Photos'

            temp_files = []

            for tl in timelines:
                tl_folder = os.path.join(folder, tl)
                imgs = []
                for r, _, files in os.walk(tl_folder):
                    for f in files:
                        if f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp')):
                            imgs.append(os.path.join(r, f))
                if not imgs:
                    row = photo_table.add_row().cells
                    row[0].text = tl
                    row[1].text = 'No photo found'
                    continue
                imgs = sorted(imgs)
                if report_max_images and report_max_images > 0:
                    imgs = imgs[:report_max_images]

                row = photo_table.add_row().cells
                row[0].text = tl

                for img_path in imgs:
                    try:
                        img = Image.open(img_path)
                        target_w, target_h = report_canvas_w, report_canvas_h
                        w, h = img.size
                        scale = min(float(target_w) / w, float(target_h) / h, 1.0)
                        new_w = int(w * scale)
                        new_h = int(h * scale)
                        img_resized = img.resize((new_w, new_h), Image.LANCZOS)
                        canvas_img = Image.new('RGB', (target_w, target_h), (255, 255, 255))
                        paste_x = (target_w - new_w) // 2
                        paste_y = (target_h - new_h) // 2
                        canvas_img.paste(img_resized.convert('RGB'), (paste_x, paste_y))
                        tf = tempfile.NamedTemporaryFile(delete=False, suffix='.jpg')
                        tmp_path = tf.name
                        try:
                            canvas_img.save(tmp_path, format='JPEG', quality=85)
                        finally:
                            tf.close()
                        temp_files.append(tmp_path)

                        para = row[1].add_paragraph()
                        run = para.add_run()
                        run.add_picture(tmp_path, width=Inches(report_embed_width))
                    except Exception as e:
                        row[1].add_paragraph(f'Error adding image {os.path.basename(img_path)}: {e}')

        # Prompt user where to save the report (Save As dialog)
        initial_name = f"{part_name}_Part_Report.docx"
        start_dir = folder if os.path.isdir(folder) else os.getcwd()
        save_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                                 initialdir=start_dir,
                                                 initialfile=initial_name,
                                                 filetypes=[("Word Documents", "*.docx")],
                                                 title="Save Report As")
        if not save_path:
            print("Report save cancelled by user.")
            try:
                for p in locals().get('temp_files', []) or []:
                    try:
                        os.remove(p)
                    except Exception:
                        pass
            except Exception:
                pass
            return

        doc.save(save_path)
        print(f"Report saved at: {save_path}")

        try:
            for p in locals().get('temp_files', []) or []:
                try:
                    os.remove(p)
                except Exception:
                    pass
        except Exception:
            pass

    except Exception as e:
        print(f"Failed to generate default report: {e}")

# Note: do not auto-run generate_report() here.
# Reports will be generated only when the user clicks the "Generate Report" button
# which always opens a Save As prompt. Do not change this behavior.


# Scrollable button bar so buttons adapt to small screens
button_bar_container = tk.Frame(root, bg="white", pady=8)
button_bar_container.pack(fill="x")

canvas_btn = tk.Canvas(button_bar_container, height=52, bg="white", highlightthickness=0)
hscroll = ttk.Scrollbar(button_bar_container, orient="horizontal", command=canvas_btn.xview)
canvas_btn.configure(xscrollcommand=hscroll.set)
hscroll.pack(side="bottom", fill="x")
canvas_btn.pack(side="top", fill="x", expand=True)

# This frame will hold the buttons and be scrolled horizontally when needed
button_frame = tk.Frame(canvas_btn, bg="white")
canvas_btn.create_window((0, 0), window=button_frame, anchor="nw")

def _on_button_frame_config(event=None):
    try:
        canvas_btn.configure(scrollregion=canvas_btn.bbox("all"))
    except Exception:
        pass

button_frame.bind("<Configure>", _on_button_frame_config)
frame_input = tk.Frame(root, bg="white", padx=10, pady=5)
frame_input.pack(fill="x")
labels1 = ["Project ID", "Part Name", "Part Number", "vLims ID", "Type of chamber"]
labels2 = ["Test standard", "Background details", "Start Date (DD-MM-YYYY)", "Spec Hours"]

for i, lab in enumerate(labels1):
    tk.Label(frame_input, text=lab, width=18, anchor="w", bg="white", font=("Arial", 10, "bold")).grid(row=0, column=i, padx=5)
for i, lab in enumerate(labels2):
    tk.Label(frame_input, text=lab, width=18, anchor="w", bg="white", font=("Arial", 10, "bold")).grid(row=2, column=i, padx=5)

entry_project_id = tk.Entry(frame_input, width=15)
entry_part_name = tk.Entry(frame_input, width=15)
entry_part_number = tk.Entry(frame_input, width=15)
entry_vlims_id = tk.Entry(frame_input, width=15)
type_chamber_var = tk.StringVar()
entry_type_chamber = ttk.Combobox(frame_input, textvariable=type_chamber_var, width=15, state="readonly")
entry_type_chamber['values'] = ("NSS", "CCT")
entry_type_chamber.current(0)
entry_test_standard = tk.Entry(frame_input, width=15)
entry_background_details = tk.Entry(frame_input, width=15)
entry_start_date = tk.Entry(frame_input, width=15)
entry_spec_hours = tk.Entry(frame_input, width=10)

entry_project_id.grid(row=1, column=0)
entry_part_name.grid(row=1, column=1)
entry_part_number.grid(row=1, column=2)
entry_vlims_id.grid(row=1, column=3)
entry_type_chamber.grid(row=1, column=4)
entry_test_standard.grid(row=3, column=0)
entry_background_details.grid(row=3, column=1)
entry_start_date.grid(row=3, column=2)
entry_spec_hours.grid(row=3, column=3)

entries = [entry_project_id, entry_part_name, entry_part_number,
           entry_vlims_id, entry_test_standard, entry_background_details,
           entry_start_date, entry_spec_hours]

result_frame = tk.Frame(root, bg="white")
result_frame.pack(fill="x", padx=10, pady=5)

tk.Label(result_frame, text="Result (for update):", bg="white", font=("Arial", 10, "bold")).pack(side="left", padx=5)
entry_result_update = tk.Entry(result_frame, width=15)
entry_result_update.pack(side="left", padx=5)
tk.Button(result_frame, text="Update Result", bg="#F39C12", fg="white", font=("Arial", 10, "bold"), command=update_result).pack(side="left", padx=10)

tree_frame = tk.Frame(root, bg="white", highlightbackground="black", highlightthickness=3)
tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=15)
scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
scroll_y.pack(side="right", fill="y")
scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
scroll_x.pack(side="bottom", fill="x")

tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=120, anchor="center")

tree.pack(fill="both", expand=True)
btn_add = tk.Button(button_frame, text="Add to Board", bg="#4CAF50", fg="white", width=13, font=("Arial", 10, "bold"), command=add_record)
btn_delete = tk.Button(button_frame, text="Delete Selected", bg="#e74c3c", fg="white", width=15, font=("Arial", 10, "bold"), command=delete_selected)
btn_update = tk.Button(button_frame, text="Update Selected", bg="#8E44AD", fg="white", width=15, font=("Arial", 10, "bold"), command=update_selected_details)
btn_breakdown = tk.Button(button_frame, text="Break Down", bg="#D35400", fg="white", width=13, font=("Arial", 10, "bold"), command=open_breakdown_popup)
btn_breakdown_history = tk.Button(button_frame, text="Breakdown History", bg="#3498DB", fg="white", width=16, font=("Arial", 10, "bold"), command=open_breakdown_history)
btn_export = tk.Button(button_frame, text="Export Excel Report", bg="#2196F3", fg="white", width=18, font=("Arial", 10, "bold"), command=export_data)
btn_upload_photo = tk.Button(button_frame, text="Upload Photo", bg="#F39C12", fg="white", width=14, font=("Arial", 10, "bold"), command=upload_photo_for_selected)
btn_view_images = tk.Button(button_frame, text="View Images", bg="#2196F3", fg="white", width=14, font=("Arial", 10, "bold"), command=view_images_for_selected)
btn_view_folder = tk.Button(button_frame, text="View Images Folder", bg="#3F51B5", fg="white", width=16, font=("Arial", 10, "bold"), command=open_images_folder_selected)
btn_generate_report = tk.Button(button_frame, text="Generate Report", bg="#8EB53F", fg="white", width=16, font=("Arial", 10, "bold"), command=generate_report)
btn_generate_report.pack(side="left", pady=6)
btn_add.pack(side="left", pady=6)
btn_delete.pack(side="left", pady=6)
btn_update.pack(side="left", pady=6)
btn_breakdown.pack(side="left", pady=6)
btn_breakdown_history.pack(side="left", pady=6)
btn_export.pack(side="left", pady=6)
btn_upload_photo.pack(side="left", pady=6)
btn_view_images.pack(side="left", pady=6)
btn_view_folder.pack(side="left", pady=6)
# Report Settings button removed per user request

# Add a menu bar with a File menu so Generate Report is always reachable
menubar = tk.Menu(root)
file_menu = tk.Menu(menubar, tearoff=0)

# Handler to let user select and persist a report template (.docx)
def set_report_template():
    tpl = filedialog.askopenfilename(title='Select Report Template (.docx)', filetypes=[('Word Documents','*.docx')])
    if not tpl:
        return
    config['report_template_path'] = tpl
    try:
        save_config(config)
        messagebox.showinfo('Template Saved', f'Report template saved:\n{tpl}')
    except Exception as e:
        messagebox.showerror('Save Error', f'Could not save template path:\n{e}')

def inspect_template():
    """Quickly inspect the configured template and report placeholders/tables."""
    tpl = config.get('report_template_path', '')
    if not tpl or not os.path.exists(tpl):
        messagebox.showinfo('Inspect Template', 'No template configured. Use File > Set Report Template to choose a .docx file.')
        return
    try:
        doc = Document(tpl)
    except Exception as e:
        messagebox.showerror('Inspect Template', f'Failed to open template:\n{e}')
        return

    # collect placeholders of form {{...}}
    import re
    ph_set = set()
    pattern = re.compile(r"\{\{\s*([^}]+?)\s*\}\}")
    for p in doc.paragraphs:
        for m in pattern.findall(p.text):
            ph_set.add(m.strip())
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for m in pattern.findall(cell.text):
                    ph_set.add(m.strip())

    info = []
    info.append(f'Template: {tpl}')
    info.append(f'Tables found: {len(doc.tables)}')
    if ph_set:
        info.append('Placeholders found:')
        for ph in sorted(ph_set):
            info.append('  - ' + ph)
    else:
        info.append('No {{placeholders}} found.')

    messagebox.showinfo('Inspect Template', '\n'.join(info))

file_menu.add_command(label='Set Report Template', command=set_report_template)
file_menu.add_command(label='Inspect Template', command=inspect_template)
file_menu.add_command(label='Generate Report (Template)', command=generate_report)
file_menu.add_command(label='Generate Report (Default)', command=generate_report_default)
file_menu.add_separator()
file_menu.add_command(label='Exit', command=root.quit)
menubar.add_cascade(label='File', menu=file_menu)
root.config(menu=menubar)

# Keyboard shortcut for generate report
def _on_ctrl_g(event=None):
    generate_report()

root.bind('<Control-g>', _on_ctrl_g)
now = datetime.now()
tree.insert("", tk.END, values=("P001", "Part A", "PA01", "VL123", "NSS",
                                "Standard1", "Background #1", now.strftime("%d-%m-%Y"),
                                (now + timedelta(days=1)).strftime("%d-%m-%Y"), 24, 0.0, "In-Progress", ""))
tree.insert("", tk.END, values=("P002", "Part B", "PB02", "VL456", "CCT",
                                "Standard2", "Background #2", (now - timedelta(days=3)).strftime("%d-%m-%Y"),
                                (now - timedelta(days=1)).strftime("%d-%m-%Y"), 48, 0.0, "Completed", "Pass"))

apply_row_colors()
update_actual_hours_all()
update_chamber_counts()
draw_monthly_breakdown_boxes(frame_calendar, breakdown_history_file)
def periodic_actual_hours_update():
    update_actual_hours_all()
    apply_row_colors()
    root.after(60*1000, periodic_actual_hours_update)  # Update every 60 seconds

periodic_actual_hours_update()
root.mainloop()
