import os
import sys
import datetime
import json

# --- INITIALIZE RESOLVE API ---
def get_resolve():
    try:
        # Check if already in DaVinci Resolve
        if 'resolve' in globals():
            return globals()['resolve']
        
        # Try importing the script API
        import DaVinciResolveScript as dvr
        return dvr.scriptapp("Resolve")
    except ImportError:
        return None

resolve = get_resolve()
if not resolve:
    # Try common Windows path as fallback
    try:
        sys.path.append(r"C:\Program Files\Blackmagic Design\DaVinci Resolve")
        import DaVinciResolveScript as dvr
        resolve = dvr.scriptapp("Resolve")
    except:
        pass

if not resolve:
    print("Error: Could not find DaVinci Resolve. Please run from Workspace > Scripts.")
    sys.exit()

fusion = resolve.Fusion()
ui = fusion.UIManager
dispatcher = fusion.FindEventDispatcher(ui)

# --- UI CONSTANTS ---
WIN_ID = "com.timeline.exporter.resolve"
WINDOW_TITLE = "Timeline Reporter - Resolve Edition"

# --- HELPER FUNCTIONS ---
def frames_to_tc(frames, fps):
    frames = int(frames)
    hh = frames // (fps * 3600)
    mm = (frames // (fps * 60)) % 60
    ss = (frames // fps) % 60
    ff = frames % fps
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

# --- MAIN WINDOW DEFINITION ---
win = ui.NewWindow(WIN_ID, WINDOW_TITLE, [
    ui.VGroup([
        # Header
        ui.HGroup({"Weight": 0}, [
            ui.Label({"ID": "HeaderLabel", "Text": WINDOW_TITLE, "Font": ui.Font({"PixelSize": 18, "Bold": True})}),
        ]),
        
        ui.Label({"Text": "Search Filters (One per line)"}),
        ui.TextEdit({"ID": "FileNamesList", "PlaceholderText": "Ex: Clip_01\nScene_A", "Weight": 1}),
        
        ui.Label({"Text": "Scan Options"}),
        ui.ComboBox({"ID": "ScanType", "Text": "Full Video & Audio"}), # Add items later
        
        ui.Label({"Text": "Filter by Color"}),
        ui.ComboBox({"ID": "ColorFilter", "Text": "Any Color"}),
        
        ui.Label({"Text": "Columns to Include"}),
        ui.HGroup([
            ui.VGroup([
                ui.CheckBox({"ID": "ChkLabel", "Text": "Color Label", "Checked": True}),
                ui.CheckBox({"ID": "ChkComments", "Text": "Comments", "Checked": True}),
                ui.CheckBox({"ID": "ChkFrames", "Text": "Duration (Frames)", "Checked": True}),
            ]),
            ui.VGroup([
                ui.CheckBox({"ID": "ChkCodecs", "Text": "Codecs", "Checked": True}),
                ui.CheckBox({"ID": "ChkFrameRate", "Text": "Frame Rate", "Checked": True}),
                ui.CheckBox({"ID": "ChkSize", "Text": "Aspect / Size", "Checked": True}),
                ui.CheckBox({"ID": "ChkFilePath", "Text": "File Path", "Checked": True}),
            ])
        ]),

        ui.Label({"Text": "Output Folder"}),
        ui.HGroup({"Weight": 0}, [
            ui.LineEdit({"ID": "OutputPath", "PlaceholderText": "Select folder..."}),
            ui.Button({"ID": "BrowseBtn", "Text": "Browse", "Weight": 0}),
        ]),
        
        ui.Button({"ID": "ExportBtn", "Text": "GENERATE REPORT", "MinimumSize": [0, 40]}),
        
        ui.Label({"ID": "StatusLabel", "Text": "Ready", "Alignment": {"H": "Center"}}),
    ])
])

# Initialize Dropdowns
win.Find("ScanType").AddItems(["Video Only", "Audio Only", "Full Video & Audio"])
color_options = ["Any Color", "Orange", "Apricot", "Yellow", "Lime", "Green", "Jade", "Cyan", "Sky", "Blue", "Sand", "Brown", "Tan", "Violet", "Pink", "Rose", "Lavender", "Purple", "Cerulean"]
win.Find("ColorFilter").AddItems(color_options)

# --- EVENT HANDLERS ---
def on_close(ev):
    dispatcher.ExitLoop()

def on_browse(ev):
    path = fusion.RequestDir()
    if path:
        win.Find("OutputPath").Text = path

def on_export(ev):
    # Setup Resolve objects
    pm = resolve.GetProjectManager()
    proj = pm.GetCurrentProject()
    if not proj:
        win.Find("StatusLabel").Text = "Error: No project open"
        return
    
    tl = proj.GetCurrentTimeline()
    if not tl:
        win.Find("StatusLabel").Text = "Error: No timeline active"
        return
    
    out_dir = win.Find("OutputPath").Text
    if not out_dir or not os.path.exists(out_dir):
        win.Find("StatusLabel").Text = "Error: Select a valid output folder"
        return

    win.Find("StatusLabel").Text = "Processing timeline..."
    win.Find("ExportBtn").Enabled = False
    
    # Gather settings from UI
    target_names = win.Find("FileNamesList").PlainText.split('\n')
    target_names = [n.strip() for n in target_names if n.strip()]
    
    scan_type = win.Find("ScanType").CurrentText
    color_filter = win.Find("ColorFilter").CurrentText
    
    cols = {
        "label": win.Find("ChkLabel").Checked,
        "comments": win.Find("ChkComments").Checked,
        "frames": win.Find("ChkFrames").Checked,
        "codecs": win.Find("ChkCodecs").Checked,
        "frameRate": win.Find("ChkFrameRate").Checked,
        "size": win.Find("ChkSize").Checked,
        "filePath": win.Find("ChkFilePath").Checked,
    }

    # Process Clipes
    all_clips = []
    fps = float(proj.GetSetting("timelineFrameRate"))
    
    # Track mapping
    tracks_to_scan = []
    if scan_type in ["Video Only", "Full Video & Audio"]:
        num_v = tl.GetTrackCount("video")
        for i in range(1, int(num_v) + 1):
            tracks_to_scan.append(("video", i))
            
    if scan_type in ["Audio Only", "Full Video & Audio"]:
        num_a = tl.GetTrackCount("audio")
        for i in range(1, int(num_a) + 1):
            tracks_to_scan.append(("audio", i))

    for track_type, track_idx in tracks_to_scan:
        items = tl.GetItemListInTrack(track_type, track_idx)
        if not items: continue
        
        for item in items:
            name = item.GetName()
            color = item.GetClipColor()
            
            # Filter by name
            if target_names:
                match = False
                for t in target_names:
                    if t.lower() in name.lower():
                        match = True
                        break
                if not match: continue
            
            # Filter by color
            if color_filter != "Any Color" and color != color_filter:
                continue

            mp_item = item.GetMediaPoolItem()
            if not mp_item: continue
            
            # Extract data
            duration_frames = int(item.GetDuration())
            t_in = int(item.GetStart())
            t_out = int(item.GetEnd())
            s_in = mp_item.GetClipProperty("Start TC") if mp_item else "-"
            # DaVinci API doesn't have direct Source Out as easily as Start, 
            # so we approximate or use Source Duration
            
            clip_data = {
                "thumbnail": "🎵" if track_type == "audio" else "🎞️",
                "color": get_color_hex(color),
                "name": name,
                "filePath": mp_item.GetClipProperty("File Path") if mp_item else "-",
                "track": ("V" if track_type == "video" else "A") + str(track_idx),
                "fps": mp_item.GetClipProperty("FPS") if mp_item else "-",
                "size": mp_item.GetClipProperty("Resolution") if mp_item else "-",
                "codec": mp_item.GetClipProperty("Format") if mp_item else "-",
                "tIn": frames_to_tc(t_in, fps),
                "tOut": frames_to_tc(t_out, fps),
                "sIn": s_in,
                "sOut": "-", # Source out requires more math in Resolve API
                "duration": frames_to_tc(duration_frames, fps),
                "frames": duration_frames,
                "comments": mp_item.GetMetadata("Description") or mp_item.GetMetadata("Comments") or "-",
                "start_frame": t_in
            }
            all_clips.append(clip_data)

    # Sort by start frame
    all_clips.sort(key=lambda x: x["start_frame"])

    # --- HTML GENERATION ---
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
        
    win.Find("StatusLabel").Text = f"SUCCESS: Report saved to {filename}"
    win.Find("ExportBtn").Enabled = True

# --- BIND EVENTS ---
win.On.MainWindow.Close = on_close
win.On.BrowseBtn.Clicked = on_browse
win.On.ExportBtn.Clicked = on_export

# --- SHOW WINDOW ---
win.Show()
dispatcher.RunLoop()
win.Hide()
