import os
import sys
import datetime
import time
import base64

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:
    print("Tkinter not found. Please ensure Python is fully installed with Tkinter support.")
    sys.exit()

# --- INITIALIZE RESOLVE API ---
def get_bmd_objects():
    res = globals().get("resolve")
    if not res:
        try:
            import DaVinciResolveScript as dvr
            res = dvr.scriptapp("Resolve")
        except: pass
    return res

resolve = get_bmd_objects()
if not resolve:
    print("Error: Could not connect to DaVinci Resolve API.")
    sys.exit()

# --- HELPER FUNCTIONS ---
def frames_to_tc(frames, fps):
    fps_int = int(round(fps))
    if fps_int == 0: fps_int = 24
    frames = int(frames)
    hh = frames // (fps_int * 3600)
    mm = (frames // (fps_int * 60)) % 60
    ss = (frames // fps_int) % 60
    ff = frames % fps_int
    return "{:02d}:{:02d}:{:02d}:{:02d}".format(int(hh), int(mm), int(ss), int(ff))

def get_color_hex(color_name):
    colors = {
        "Orange": "#f7aa67", "Apricot": "#f7aa67", "Yellow": "#e0c242", "Lime": "#66b574",
        "Green": "#589f66", "Jade": "#37b79a", "Cyan": "#4ad1b7", "Sky": "#89aef0",
        "Blue": "#418dec", "Sand": "#d8a071", "Brown": "#b07e59", "Tan": "#d8a071",
        "Violet": "#ea8ff3", "Pink": "#f68dae", "Rose": "#ec659c", "Lavender": "#cfb3ed",
        "Purple": "#bf90ed", "Cerulean": "#74a0e6", "None": "#ffffff"
    }
    return colors.get(color_name, "#ffffff")

# --- UI CLASS ---
class ExporterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Timeline Reporter - Resolve Free/Studio")
        self.root.geometry("550x700")
        
        style = ttk.Style()
        try: style.theme_use('clam')
        except: pass
        
        # Container
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        ttk.Label(main_frame, text="Timeline Reporter", font=("Helvetica", 16, "bold")).pack(pady=(0, 15))
        
        # Filters
        ttk.Label(main_frame, text="Search Filters (One per line)").pack(anchor=tk.W)
        self.text_filters = tk.Text(main_frame, height=4, width=50)
        self.text_filters.pack(fill=tk.X, pady=(0, 15))
        
        # Options Frame
        opt_frame = ttk.Frame(main_frame)
        opt_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(opt_frame, text="Scan Options:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.cb_scan = ttk.Combobox(opt_frame, values=["Video Only", "Audio Only", "Full Video & Audio"], state="readonly")
        self.cb_scan.current(0)
        self.cb_scan.grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)
        
        ttk.Label(opt_frame, text="Filter by Color:").grid(row=1, column=0, sticky=tk.W, pady=5)
        colors = ["Any Color", "Orange", "Apricot", "Yellow", "Lime", "Green", "Jade", "Cyan", "Sky", "Blue", "Sand", "Brown", "Tan", "Violet", "Pink", "Rose", "Lavender", "Purple", "Cerulean"]
        self.cb_color = ttk.Combobox(opt_frame, values=colors, state="readonly")
        self.cb_color.current(0)
        self.cb_color.grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)
        
        # Columns
        ttk.Label(main_frame, text="Columns to Include", font=("Helvetica", 10, "bold")).pack(anchor=tk.W, pady=(5,5))
        col_frame = ttk.Frame(main_frame)
        col_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.chk_vars = {
            "label": tk.BooleanVar(value=True),
            "comments": tk.BooleanVar(value=True),
            "frames": tk.BooleanVar(value=True),
            "codecs": tk.BooleanVar(value=True),
            "frameRate": tk.BooleanVar(value=True),
            "size": tk.BooleanVar(value=True),
            "filePath": tk.BooleanVar(value=True)
        }
        
        ttk.Checkbutton(col_frame, text="Color Label", variable=self.chk_vars["label"]).grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Checkbutton(col_frame, text="Comments", variable=self.chk_vars["comments"]).grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Checkbutton(col_frame, text="Duration (Frames)", variable=self.chk_vars["frames"]).grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        
        ttk.Checkbutton(col_frame, text="Codecs", variable=self.chk_vars["codecs"]).grid(row=0, column=1, sticky=tk.W, padx=20, pady=2)
        ttk.Checkbutton(col_frame, text="Frame Rate", variable=self.chk_vars["frameRate"]).grid(row=1, column=1, sticky=tk.W, padx=20, pady=2)
        ttk.Checkbutton(col_frame, text="Aspect / Size", variable=self.chk_vars["size"]).grid(row=2, column=1, sticky=tk.W, padx=20, pady=2)
        ttk.Checkbutton(col_frame, text="File Path", variable=self.chk_vars["filePath"]).grid(row=3, column=1, sticky=tk.W, padx=20, pady=2)
        
        # Output folder
        ttk.Label(main_frame, text="Output Folder").pack(anchor=tk.W)
        out_frame = ttk.Frame(main_frame)
        out_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.out_path = tk.StringVar()
        ttk.Entry(out_frame, textvariable=self.out_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(out_frame, text="Browse", command=self.browse_folder).pack(side=tk.RIGHT)
        
        # Export Button
        self.btn_export = ttk.Button(main_frame, text="GENERATE REPORT", command=self.export_report)
        self.btn_export.pack(fill=tk.X, pady="10")
        
        self.lbl_status = ttk.Label(main_frame, text="Ready", foreground="gray")
        self.lbl_status.pack(pady=5)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.out_path.set(folder)

    def export_report(self):
        out_dir = self.out_path.get()
        if not out_dir or not os.path.exists(out_dir):
            messagebox.showerror("Error", "Select a valid output folder")
            return
            
        pm = resolve.GetProjectManager()
        proj = pm.GetCurrentProject()
        if not proj:
            messagebox.showerror("Error", "No project open")
            return
            
        tl = proj.GetCurrentTimeline()
        if not tl:
            messagebox.showerror("Error", "No timeline active")
            return
            
        self.lbl_status.config(text="Processing timeline...", foreground="blue")
        self.btn_export.config(state=tk.DISABLED)
        self.root.update()

        try:
            self._process_timeline(proj, tl, out_dir)
            self.lbl_status.config(text="SUCCESS: Report generated!", foreground="green")
            messagebox.showinfo("Success", "Timeline report generated successfully in:\n" + out_dir)
        except Exception as e:
            self.lbl_status.config(text=f"Error: {str(e)}", foreground="red")
            messagebox.showerror("Error", str(e))
        finally:
            self.btn_export.config(state=tk.NORMAL)

    def _process_timeline(self, proj, tl, out_dir):
        # Settings
        target_names = [n.strip() for n in self.text_filters.get("1.0", tk.END).split('\n') if n.strip()]
        scan_type = self.cb_scan.get()
        color_filter = self.cb_color.get()
        cols = {k: v.get() for k, v in self.chk_vars.items()}

        thumbs_dir = os.path.join(out_dir, "Thumbnails")
        if not os.path.exists(thumbs_dir):
            os.makedirs(thumbs_dir)

        all_clips = []
        fps = float(proj.GetSetting("timelineFrameRate"))
        resolve.OpenPage("color")
        
        num_v = int(tl.GetTrackCount("video")) if tl.GetTrackCount("video") else 0
        original_video_tracks = {}
        for i in range(1, num_v + 1):
            original_video_tracks[i] = tl.GetIsTrackEnabled("video", i)
        
        tracks_to_scan = []
        if scan_type in ["Video Only", "Full Video & Audio"] and num_v > 0:
            for i in range(1, num_v + 1): tracks_to_scan.append(("video", i))
        
        num_a = int(tl.GetTrackCount("audio")) if tl.GetTrackCount("audio") else 0
        if scan_type in ["Audio Only", "Full Video & Audio"] and num_a > 0:
            for i in range(1, num_a + 1): tracks_to_scan.append(("audio", i))

        gallery = proj.GetGallery()
        gallery_album = gallery.GetCurrentStillAlbum() if gallery else None
        
        for track_type, track_idx in tracks_to_scan:
            items = tl.GetItemListInTrack(track_type, track_idx)
            if not items: continue
            
            batch_prefix = f"thumb_v{track_idx}"
            if track_type == "video":
                for i in range(1, num_v + 1):
                    tl.SetTrackEnable("video", i, (i == track_idx))
                
                if gallery_album:
                    stills = tl.GrabAllStills(2)
                    if stills:
                        gallery_album.ExportStills(stills, thumbs_dir, batch_prefix, "jpg")
                        gallery_album.DeleteStills(stills)
                    
            for clip_idx, item in enumerate(items, 1):
                name = item.GetName()
                color = item.GetClipColor()
                
                if target_names and not any(t.lower() in name.lower() for t in target_names):
                    continue
                if color_filter != "Any Color" and color != color_filter:
                    continue

                mp_item = item.GetMediaPoolItem()
                if not mp_item: continue
                
                t_in = int(item.GetStart())
                t_out = int(item.GetEnd())
                duration_frames = t_out - t_in
                s_in = mp_item.GetClipProperty("Start TC") if mp_item else "-"
                
                thumb_html = "🎵" if track_type == "audio" else "🎞️"
                
                if track_type == "video" and gallery_album:
                    expected_start = f"{batch_prefix}_{track_idx}.{clip_idx}."
                    try:
                        for f in os.listdir(thumbs_dir):
                            if f.startswith(expected_start) and f.endswith(".jpg"):
                                thumb_path = os.path.join(thumbs_dir, f)
                                with open(thumb_path, "rb") as image_file:
                                    encoded_string = base64.b64encode(image_file.read()).decode()
                                    thumb_html = f"<img src='data:image/jpeg;base64,{encoded_string}' style='max-width:80px; max-height:45px; border-radius:4px;'/>"
                                os.remove(thumb_path)
                                drx_path = thumb_path.replace(".jpg", ".drx")
                                if os.path.exists(drx_path): os.remove(drx_path)
                                break
                    except Exception as e:
                        pass
                
                clip_data = {
                    "thumbnail": thumb_html,
                    "color": get_color_hex(color),
                    "name": name,
                    "filePath": mp_item.GetClipProperty("File Path") if mp_item else "-",
                    "track": ("V" if track_type == "video" else "A") + str(track_idx),
                    "fps": mp_item.GetClipProperty("FPS") if mp_item else "-",
                    "size": mp_item.GetClipProperty("Resolution") if mp_item else "-",
                    "codec": mp_item.GetClipProperty("Video Codec") if mp_item else "-", 
                    "tIn": frames_to_tc(t_in, fps),
                    "tOut": frames_to_tc(t_out, fps),
                    "sIn": s_in,
                    "sOut": "-", 
                    "duration": frames_to_tc(duration_frames, fps),
                    "frames": duration_frames,
                    "comments": mp_item.GetMetadata("Description") or mp_item.GetMetadata("Comments") or "-",
                    "start_frame": t_in
                }
                all_clips.append(clip_data)

        if num_v > 0:
            for i in range(1, num_v + 1):
                tl.SetTrackEnable("video", i, original_video_tracks.get(i, True))

        all_clips.sort(key=lambda x: x["start_frame"])

        # HTML Generation
        html = f"""<!DOCTYPE html><html><head><meta charset='utf-8'><style>
            body{{font-family:'Inter', Arial, sans-serif; padding:20px; background:#f4f4f4; color: #333;}} 
            table{{border-collapse:collapse; width:100%; background:white; box-shadow:0 4px 12px rgba(0,0,0,0.1); border-radius: 8px; overflow: hidden;}} 
            th,td{{border:1px solid #eee; padding:12px; text-align:center; vertical-align:middle; font-size: 13px;}} 
            td.color-cell{{padding:0 !important; width:12px;}} 
            th{{background:#1a1a1a; color:white; font-size:11px; text-transform:uppercase; letter-spacing: 1px;}} 
            tr:nth-child(even){{background: #fafafa;}} 
            h2{{text-align: center; color: #1a1a1a; margin-bottom: 30px; border-bottom: 2px solid #6366f1; display: inline-block; padding-bottom: 5px;}}
            .header-container {{ text-align: center; }}
        </style></head><body>
        <div class='header-container'><h2>{tl.GetName()} - Timeline Report</h2></div>
        <table><tr><th>Thumb</th>"""
        
        if cols["label"]: html += "<th>Color</th>"
        html += "<th>File</th>"
        if cols["filePath"]: html += "<th>File Path</th>"
        html += "<th>Track</th>"
        if cols["frameRate"]: html += "<th>FPS</th>"
        if cols["size"]: html += "<th>Resolution</th>"
        if cols["codecs"]: html += "<th>Codec</th>"
        html += "<th>Timeline IN</th><th>Timeline OUT</th><th>Source IN</th><th>Duration</th>"
        if cols["frames"]: html += "<th>Frames</th>"
        if cols["comments"]: html += "<th>Comments</th>"
        html += "</tr>"

        for c in all_clips:
            html += f"<tr><td>{c['thumbnail']}</td>"
            if cols["label"]: html += f"<td class='color-cell' style='background-color: {c['color']};'></td>"
            html += f"<td>{c['name']}</td>"
            if cols["filePath"]: html += f"<td>{c['filePath']}</td>"
            html += f"<td>{c['track']}</td>"
            if cols["frameRate"]: html += f"<td>{c['fps']}</td>"
            if cols["size"]: html += f"<td>{c['size']}</td>"
            if cols["codecs"]: html += f"<td>{c['codec']}</td>"
            html += f"<td>{c['tIn']}</td><td>{c['tOut']}</td><td>{c['sIn']}</td><td>{c['duration']}</td>"
            if cols["frames"]: html += f"<td>{c['frames']}</td>"
            if cols["comments"]: html += f"<td>{c['comments']}</td>"
            html += "</tr>"

        html += "</table><div style='text-align:center; margin-top:20px; font-size:10px; color:#999;'>Generated by Timeline Reporter for DaVinci Resolve</div></body></html>"

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Resolve_Report_{timestamp}.html"
        full_path = os.path.join(out_dir, filename)
        
        with open(full_path, "w", encoding="utf-8") as f:
            f.write(html)

if __name__ == "__main__":
    root = tk.Tk()
    # Force tk window to top
    root.lift()
    root.attributes('-topmost', True)
    root.after_idle(root.attributes, '-topmost', False)
    
    app = ExporterApp(root)
    root.mainloop()
